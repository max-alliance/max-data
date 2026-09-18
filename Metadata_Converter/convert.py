#!/usr/bin/env python3
"""template.xlsx -> Croissant metadata.jsonld 변환기.

사용법:
    python convert.py                          # 기본값: -i template.xlsx -o <데이터셋ID>.jsonld
    python convert.py -i template.xlsx -o metadata.jsonld
    python convert.py -i MAX_TR_DS01.xlsx      # 데이터 폴더가 있는 위치에서 실행하면 자동 검증
    python convert.py -i MAX_TR_DS01.xlsx -d /data/root   # 데이터 상위 폴더를 직접 지정(선택)
    python convert.py -i MAX_TR_DS01.xlsx --no-verify     # 데이터 없이 메타데이터만 생성

변환 전에 필수값/형식/데이터 경로를 검사한다. 데이터 파일 존재 여부는 엑셀의
'데이터 경로'(01/02 시트)와 '스플릿 정보'(03 시트)를 합쳐 만든 실제 경로
(예: ./MAX_SECOM_DS01/data/train/)를 검사한다. 기준 폴더는 현재 작업 폴더 또는
엑셀 파일이 있는 폴더에서 자동으로 찾으며, -d/--base-dir 로 직접 지정할 수도 있다.
문제(ERROR)가 하나라도 있으면 출력 파일을 만들지 않고 "이렇게 고치세요" 안내만
출력하고 종료 코드 1을 반환한다(--no-verify 로 파일 검사를 건너뛸 수 있음).

------------------------------------------------------------------------------
데이터-라벨 매핑(02_라벨 정보 시트)
------------------------------------------------------------------------------
이 변환기는 라벨이 데이터와 "어떻게 짝지어지는가"(매핑 기준)에 따라 Croissant 구조를
다르게 생성한다. 매핑 기준은 4가지다:

  column   : 데이터와 같은 행(레코드) 안에 라벨이 들어 있음(조인 불필요).
             예) SECOM CSV 의 마지막 컬럼 Pass/Fail.
             → 별도 라벨 FileSet 을 만들지 않고, 데이터 파일(FileObject)에서
               '라벨 컬럼'을 추출하는 Field 로 표현한다.
  filename : 데이터 파일명으로 라벨 파일을 찾음(예: 이미지 a.jpg ↔ 라벨 a.txt).
             → 데이터/라벨을 각각 FileSet 으로 만들고, 파일명+내용을 그대로 내주는
               RecordSet 을 만든다(로더가 파일명으로 매핑).
  id/key   : 명시적 식별자 컬럼으로 두 테이블을 조인. 라벨은 테이블의 한 컬럼.
             → column 과 동일하게 '라벨 컬럼'을 추출하되, 조인 키가 필요함을 비고로 남긴다.
  index    : 명시적 키 없이 저장 순서(i번째↔i번째)로 매칭.
             → filename 과 동일하게 파일 기반으로 두되, 재현성 주의를 남긴다.

'라벨 컬럼'(02 시트 신설 행)은 column / id/key 일 때 필수, filename / index 이면 공란.
"""
import argparse
import csv
import datetime
import glob
import json
import os
import re
import sys

import openpyxl
from openpyxl.utils import column_index_from_string

# ---------------------------------------------------------------------------
# 상수/설정
# ---------------------------------------------------------------------------
DEFAULT_INPUT = "template.xlsx"
# 출력은 기본적으로 데이터셋 ID 로 파일명을 만든다(예: MAX_TR_DS01.jsonld).
# -o 로 명시하면 그 경로를 그대로 쓴다.

# 엑셀에 없는 값들(필요 시 이 상수만 바꾸면 됨)
LICENSE_URL = "https://maxalliance.kr/license/license.html"
# 데이터셋을 실제로 호스팅하는 URL 이 생기면 아래에 "https://.../{ds_id}" 형태로 넣는다.
# 없으면 None 으로 두어 url 필드를 아예 생성하지 않는다(존재하지 않는 링크를 넣지 않음).
URL_TEMPLATE = None

# 폴더(디렉토리)를 가리키는 FileObject 의 encodingFormat.
DIRECTORY_MIME = "application/x-directory"

EXT2MIME = {
    "jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
    "gif": "image/gif", "bmp": "image/bmp", "webp": "image/webp",
    "tif": "image/tiff", "tiff": "image/tiff",
    "txt": "text/plain", "csv": "text/csv", "tsv": "text/tab-separated-values",
    "json": "application/json", "jsonl": "application/jsonlines",
    "parquet": "application/x-parquet",
    "wav": "audio/wav", "mp3": "audio/mpeg", "flac": "audio/flac",
    "npy": "application/octet-stream", "npz": "application/octet-stream",
}

# "표(테이블)"로 읽어 컬럼을 추출할 수 있는 확장자(= 라벨이 컬럼으로 들어갈 수 있는 형식).
TABULAR_EXTS = {"csv", "tsv", "parquet"}
# 이미지로 읽을 수 있는 MIME(파일 기반 매핑에서 image RecordSet 을 만들지 판단).
IMAGE_MIMES = {"image/jpeg", "image/png", "image/gif", "image/bmp",
               "image/webp", "image/tiff"}

ALLOWED_MAPPINGS = {"column", "filename", "id/key", "index"}

# mlcroissant 1.1 이 요구하는 정식 @context.
# 주의: containedIn 은 반드시 cr:containedIn 이어야 한다. (schema.org 로 해석되면
#       FileSet 의 containedIn 링크가 무시되어 파일을 하나도 못 찾는다 — glob 이 0건이 됨)
CANONICAL_CONTEXT = {
    "@language": "en", "@vocab": "https://schema.org/", "citeAs": "cr:citeAs",
    "column": "cr:column", "conformsTo": "dct:conformsTo",
    "containedIn": "cr:containedIn",
    "cr": "http://mlcommons.org/croissant/", "rai": "http://mlcommons.org/croissant/RAI/",
    "data": {"@id": "cr:data", "@type": "@json"},
    "dataType": {"@id": "cr:dataType", "@type": "@vocab"}, "dct": "http://purl.org/dc/terms/",
    "equivalentProperty": "cr:equivalentProperty",
    "examples": {"@id": "cr:examples", "@type": "@json"}, "extract": "cr:extract",
    "field": "cr:field", "fileProperty": "cr:fileProperty", "fileObject": "cr:fileObject",
    "fileSet": "cr:fileSet", "format": "cr:format", "includes": "cr:includes",
    "isLiveDataset": "cr:isLiveDataset", "jsonPath": "cr:jsonPath", "key": "cr:key",
    "md5": "cr:md5", "parentField": "cr:parentField", "path": "cr:path",
    "recordSet": "cr:recordSet", "references": "cr:references", "regex": "cr:regex",
    "repeated": "cr:repeated", "replace": "cr:replace", "samplingRate": "cr:samplingRate",
    "sc": "https://schema.org/", "separator": "cr:separator", "source": "cr:source",
    "subField": "cr:subField", "transform": "cr:transform",
}

SPLIT_KEYS = [
    ("train", "학습 데이터 경로", "학습 데이터 비율"),
    ("validation", "검증 데이터 경로", "검증 데이터 비율"),
    ("test", "테스트 데이터 경로", "테스트 데이터 비율"),
]


# ---------------------------------------------------------------------------
# 작은 헬퍼
# ---------------------------------------------------------------------------
def is_empty(v):
    return v is None or str(v).strip() in ("", "-")


def clean_path(p):
    return str(p).lstrip("./").rstrip("/") if p else ""


def ext_of(pattern):
    """'*.jpg' -> 'jpg'. 확장자만 소문자로."""
    return str(pattern).split(".")[-1].lower().strip()


def mime_from_pattern(pattern):
    return EXT2MIME.get(ext_of(pattern), "application/octet-stream")


def parse_patterns(raw):
    """'파일 패턴' 문자열을 glob 패턴 리스트로 파싱.

    한 칸에 여러 형식이 올 수 있다: '*.jpg, *.png' / '*.jpg; *.png' / '*.jpg *.png'
    구분자는 콤마/세미콜론/파이프/공백. 순서는 유지하되 중복은 제거한다.
    """
    if is_empty(raw):
        return []
    parts = re.split(r"[\s,;|]+", str(raw).strip())
    out = []
    for p in parts:
        p = p.strip()
        if p and p not in out:
            out.append(p)
    return out


def mimes_for_patterns(patterns):
    """패턴 리스트 -> 중복 제거된 MIME 리스트(순서 유지)."""
    out = []
    for p in patterns:
        m = mime_from_pattern(p)
        if m not in out:
            out.append(m)
    return out


def is_tabular_patterns(patterns):
    """패턴 중 하나라도 표(csv/tsv/parquet)면 True."""
    return any(ext_of(p) in TABULAR_EXTS for p in patterns)


def iso(v):
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    return str(v) if v else None


# ---------------------------------------------------------------------------
# 1) 엑셀 읽기
# ---------------------------------------------------------------------------
def read_sheet(ws):
    """워크시트를 (values, required) 두 딕셔너리로 읽는다. 1행=제목, 2행=헤더."""
    values, required = {}, {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        key = row[0]
        if key is None:
            continue
        key = str(key).strip()
        values[key] = row[3]  # '입력' 열
        required[key] = (str(row[2]).strip().upper() == "Y") if row[2] else False  # '필수 여부' 열
    return values, required


def load_sheets(xlsx_path):
    """엑셀 파일에서 필요한 시트를 읽어 구조화한 dict 로 돌려준다."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    base, base_req = read_sheet(wb["기본정보"])
    data_i, data_req = read_sheet(wb["01_데이터 정보"])
    label, label_req = read_sheet(wb["02_라벨 정보"])
    split, split_req = read_sheet(wb["03_데이터 스플릿 정보"])
    # 04_AI 모델 정보 시트는 모델 설명이라 Croissant(데이터셋 기술 포맷)에 매핑하지 않고 생략한다.
    return {
        "base": base, "base_req": base_req,
        "data": data_i, "data_req": data_req,
        "label": label, "label_req": label_req,
        "split": split, "split_req": split_req,
        # (시트명, 값, 필수여부) 목록 — 필수값 검사에 사용
        "list": [
            ("기본정보", base, base_req),
            ("01_데이터 정보", data_i, data_req),
            ("02_라벨 정보", label, label_req),
            ("03_데이터 스플릿 정보", split, split_req),
        ],
    }


# ---------------------------------------------------------------------------
# 2) 파싱: 스플릿 / 클래스 / 매핑
# ---------------------------------------------------------------------------
def build_splits(split):
    """03 시트에서 값이 있는 스플릿만 리스트로 만든다(validation 은 optional)."""
    splits = []
    for name, path_key, ratio_key in SPLIT_KEYS:
        p = split.get(path_key)
        if not is_empty(p):
            splits.append({"name": name, "sub": clean_path(p), "ratio": split.get(ratio_key)})
    return splits


def parse_classes(label):
    """02 시트의 '클래스 분류'("0: 정상 / 1: 비정상")를 [(0,'정상'), ...] 로 파싱.

    - 음수 라벨(-1) 도 허용한다(예: SECOM 의 -1=Pass).
    - 정수가 아닌 키(예: 'cat')도 허용하며, 그 경우 id 를 문자열로 둔다.
    """
    cls_raw = label.get("클래스 분류")
    classes = []
    if cls_raw:
        for line in str(cls_raw).splitlines():
            line = line.strip().strip(",")
            if ":" not in line:
                continue
            k, v = line.split(":", 1)
            k, v = k.strip(), v.strip()
            if not k:
                continue
            try:
                cid = int(k)          # '-1', '0', '1' ... 정수 우선
            except ValueError:
                cid = k               # 'cat' 같은 문자열 라벨
            classes.append((cid, v))
    return classes


def get_mapping(label):
    """'데이터-라벨 매핑 기준' 값을 소문자로 정규화해서 돌려준다(없으면 '')."""
    m = label.get("데이터-라벨 매핑 기준")
    return "" if is_empty(m) else str(m).strip().lower()


def get_label_column(label):
    """'라벨 컬럼' 값(컬럼명). 없으면 ''. 값은 있는 그대로 쓰며(주석 등 임의 가공 안 함),
    실제 CSV 헤더와 일치하는지는 검증 단계에서 확인해 불일치 시 반려한다."""
    c = label.get("라벨 컬럼")
    return "" if is_empty(c) else str(c).strip()


# ---------------------------------------------------------------------------
# 2.5) 입력(Feature) 컬럼 명세 파싱 — 엑셀 열문자 셀 참조(B1:VS1 등) → 컬럼명
#      값을 관대하게 보정하지 않는다. 형식이 어긋나면 검증 단계에서 반려한다.
# ---------------------------------------------------------------------------
def _strict_cell_index(tok):
    """'B1' 같은 '1행 셀 참조'만 열 인덱스로 변환한다.

    반드시 열문자 + 행번호 '1' 형태여야 한다(헤더 강제 규격).
    'B'(행 없음)·'B2'(다른 행)·'B:VS'(열 전체)·기타 잡음은 None(=형식 오류).
    """
    m = re.match(r"^([A-Za-z]+)1$", tok.strip())
    if not m:
        return None
    try:
        return column_index_from_string(m.group(1).upper())
    except ValueError:
        return None


def parse_feature_spec(spec):
    """입력 피처 셀 참조 문자열 -> (정렬된 1-based 열 인덱스 리스트, 오류 리스트).

    허용 문법(합집합)은 1행 셀 참조뿐이다: 'B1:VS1' / 'A1,C1,E1' / 'A1:C1,E1:F1'.
    형식이 어긋난 토큰은 '보정'하지 않고 오류로 수집한다(→ 검증 단계에서 반려).
    값이 비어 있으면 (None, []) — '명세 없음'을 뜻한다(→ 라벨 외 전체가 피처).
    """
    if is_empty(spec):
        return None, []
    spec = str(spec).strip()
    idx, errs = [], []
    for tok in spec.split(","):
        tok = tok.strip()
        if not tok:
            continue
        if ":" in tok:
            a, b = tok.split(":", 1)
            ia, ib = _strict_cell_index(a), _strict_cell_index(b)
            if ia is None or ib is None:
                errs.append(f"셀 참조 '{tok}' 형식이 올바르지 않습니다. "
                            "'B1:VS1'처럼 1행 셀 참조 범위로 적으세요.")
                continue
            lo, hi = sorted((ia, ib))
            idx.extend(range(lo, hi + 1))
        else:
            i = _strict_cell_index(tok)
            if i is None:
                errs.append(f"셀 참조 '{tok}' 형식이 올바르지 않습니다. "
                            "'A1'처럼 1행 셀 참조로 적으세요(열 전체 'A:A'·다른 행 'A2' 불가).")
                continue
            idx.append(i)
    seen, out = set(), []
    for i in sorted(idx):
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out, errs


def resolve_feature_columns(idxs, header, label_col):
    """(열 인덱스 + CSV 헤더) -> (피처 컬럼명 리스트, 오류 리스트).

    - 위치 i ↔ 헤더 i-1 로 변환.
    - 범위를 벗어난 인덱스, 라벨 컬럼이 피처 범위에 섞인 경우는 '오류'로 수집한다
      (→ 검증 단계에서 반려). 관대하게 무시하지 않는다.
    """
    names, errs, oob = [], [], []
    for i in idxs:
        if 1 <= i <= len(header):
            names.append(header[i - 1])
        else:
            oob.append(i)
    if oob:
        shown = ", ".join(map(str, oob[:10])) + (f" 외 {len(oob) - 10}개" if len(oob) > 10 else "")
        errs.append(f"열 인덱스 {shown} 가 실제 컬럼 수({len(header)})를 벗어났습니다. "
                    "셀 참조 범위를 파일에 맞게 고치세요.")
    if label_col and label_col in names:
        errs.append(f"라벨 컬럼 '{label_col}' 이 입력 피처 범위에 포함돼 있습니다. "
                    "라벨은 피처 범위에서 제외하세요.")
        names = [n for n in names if str(n) != str(label_col)]
    return names, errs


# ---------------------------------------------------------------------------
# 3) 배치 계획(layout) — distribution 과 recordSet 을 한 기준으로 만들기 위한 중간 표현
# ---------------------------------------------------------------------------
def resolve_base_dir(input_path, data_root, override):
    """실제 파일 검증에 쓸 '기준 폴더'를 정한다.

    엑셀에 적힌 데이터 경로(예: ./MAX_SECOM_DS01/data)는 상대경로이므로, 어느 폴더를
    기준으로 삼을지 정해야 한다. 별도 옵션 없이도 동작하도록 다음 순서로 찾는다:
      1) override(-d/--base-dir)가 있으면 그 폴더만 사용
      2) 없으면 현재 작업 폴더(CWD) → 엑셀 파일이 있는 폴더 순으로,
         엑셀의 데이터 경로가 실제로 존재하는 첫 폴더를 기준으로 삼는다.
    반환: (base_dir, tried) — base_dir 는 최종 기준 폴더, tried 는 확인한 후보 목록.
    """
    if override:
        return override, [override]
    candidates = []
    for c in [os.getcwd(), os.path.dirname(os.path.abspath(input_path))]:
        if c and c not in candidates:
            candidates.append(c)
    for c in candidates:
        # data_root 는 폴더일 수도, 단일 파일일 수도 있으므로 존재 여부로 판단.
        if data_root and os.path.exists(os.path.join(c, data_root)):
            return c, candidates
    # 어디서도 못 찾으면 CWD 를 기준으로 두고(검증에서 '폴더 없음'으로 안내)
    return (candidates[0] if candidates else os.getcwd()), candidates


def _resolve_tabular_file(base_dir, data_root, sub, patterns):
    """`데이터 경로 + 스플릿 하위폴더` 안에서 표(csv 등) 파일 하나를 찾는다.

    예) data_root='MAX_SECOM_DS01/data', sub='train' → 'MAX_SECOM_DS01/data/train' 를 본다.
    실제 파일을 찾으면 그 상대경로를, 못 찾으면 '<폴더>/<첫 패턴>' 자리표시자를 쓴다.
    반환: (content_url, mime, note) — note 는 경고 문구(없으면 None).
    """
    rel_dir = f"{data_root}/{sub}" if sub else data_root
    abs_dir = os.path.join(base_dir, rel_dir) if base_dir else rel_dir
    if base_dir and os.path.isdir(abs_dir):
        matches = []
        for pat in patterns:
            matches += sorted(glob.glob(os.path.join(abs_dir, pat)))
        seen, uniq = set(), []
        for m in matches:
            if m not in seen:
                seen.add(m); uniq.append(m)
        if uniq:
            rel = os.path.relpath(uniq[0], base_dir).replace(os.sep, "/")
            note = None
            if len(uniq) > 1:
                note = (f"'{rel_dir}' 아래 표 파일이 {len(uniq)}개 있어 첫 파일('{rel}')만 "
                        "사용했습니다. 한 스플릿에는 표 파일 하나를 권장합니다.")
            return rel, mime_from_pattern(uniq[0]), note

    # 못 찾으면 자리표시자(첫 패턴). 로드 전에 실제 파일이 있어야 한다.
    first = patterns[0] if patterns else "*.csv"
    url = f"{rel_dir}/{first}" if rel_dir else first
    note = (f"'{rel_dir}' 에서 표 파일을 찾지 못해 경로를 패턴 그대로('{url}') 넣었습니다. "
            "데이터 경로/스플릿 폴더와 실제 파일 위치가 맞는지 확인하세요.")
    return url, mime_from_pattern(first), note


def _looks_like_file(base_dir, rel):
    """data_root 가 '폴더'가 아니라 '단일 파일'을 가리키는지 판단한다.

    실제로 파일로 존재하거나, 존재 여부와 무관하게 표 확장자(csv/tsv/parquet …)로
    끝나면 파일 경로로 본다. (폴더명에는 보통 확장자가 없다.)
    """
    if not rel:
        return False
    if base_dir and os.path.isfile(os.path.join(base_dir, rel)):
        return True
    if base_dir and os.path.isdir(os.path.join(base_dir, rel)):
        return False
    return ext_of(rel) in TABULAR_EXTS


def _resolve_tabular_set(base_dir, root, eff_splits, patterns, notes, tag):
    """root 가 단일 파일이면 [파일 하나], 폴더면 스플릿별 표 파일 목록을 만든다.

    반환: (files, single_file). files 원소 = {name, sub, url, mime}.
    """
    if _looks_like_file(base_dir, root):
        return [{"name": "all", "sub": "", "url": root,
                 "mime": mime_from_pattern(root)}], True
    files = []
    for s in eff_splits:
        url, mime, note = _resolve_tabular_file(base_dir, root, s["sub"], patterns)
        if note:
            notes.append(f"[{tag}] " + note)
        files.append({"name": s["name"], "sub": s["sub"], "url": url, "mime": mime})
    return files, (len(files) == 1 and files[0]["sub"] == "")


def plan_layout(base, data, label, splits, base_dir):
    """매핑 기준에 따라 distribution/recordSet 을 만들 '계획'을 세운다.

    반환 dict(layout):
      strategy        : 'tabular_embedded' | 'file_sidecar'
      mapping         : 원본 매핑 값
      label_column    : 라벨 컬럼명
      eff_splits      : 실제 사용할 스플릿 목록(없으면 단일 세트)
      notes           : 경고 문구 리스트
      (strategy 별 세부 필드)
    """
    mapping = get_mapping(label)
    label_column = get_label_column(label)
    data_root = clean_path(data.get("데이터 경로"))
    data_format = "" if is_empty(data.get("파일 형식")) else str(data.get("파일 형식")).strip()
    feature_spec = "" if is_empty(data.get("입력(Feature) 컬럼")) else str(data.get("입력(Feature) 컬럼")).strip()
    feature_idxs, feature_parse_errors = parse_feature_spec(feature_spec)
    data_patterns = parse_patterns(data.get("파일 패턴")) or ["*.*"]
    data_mimes = mimes_for_patterns(data_patterns)
    data_tabular = is_tabular_patterns(data_patterns)
    notes = []

    # 매핑별 전략:
    #   column → 라벨이 데이터 파일의 한 컬럼(tabular_embedded)
    #   id/key → 라벨이 '별도 파일'에 있고 공통 키로 조인(tabular_join)
    #   filename/index/(미지정) → 파일 기반(file_sidecar)
    if mapping == "column":
        strategy = "tabular_embedded"
        if not data_tabular:
            notes.append(f"[02_라벨 정보] 매핑이 '{mapping}'(컬럼 기반)인데 데이터 파일 패턴"
                         f"({', '.join(data_patterns)})이 표(csv/tsv/parquet) 형식이 아닙니다. "
                         "컬럼 추출이 가능한 형식인지 확인하세요.")
    elif mapping == "id/key":
        strategy = "tabular_join"
    else:
        # filename / index / (미지정) → 파일 기반
        strategy = "file_sidecar"

    # 스플릿(하위폴더)이 하나도 없으면 하위폴더 없는 단일 세트('all')로.
    eff_splits = splits if splits else [{"name": "all", "sub": "", "ratio": None}]

    layout = {
        "strategy": strategy, "mapping": mapping, "label_column": label_column,
        "data_root": data_root, "data_patterns": data_patterns, "data_mimes": data_mimes,
        "eff_splits": eff_splits, "notes": notes, "base_dir": base_dir,
        "feature_spec": feature_spec, "feature_idxs": feature_idxs,
        "feature_parse_errors": feature_parse_errors, "feature_resolve_errors": [],
        "data_format": data_format,
    }

    if strategy == "tabular_embedded":
        tabular_files, single = _resolve_tabular_set(
            base_dir, data_root, eff_splits, data_patterns, notes, "01_데이터 정보")
        layout["tabular_files"] = tabular_files
        layout["single_file"] = single
        if single and any(s.get("sub") for s in eff_splits):
            notes.append("[01/03] 데이터 경로가 단일 파일이라, 스플릿 하위폴더 대신 "
                         "비율(splits)로만 처리합니다. 학습 코드에서 그 비율로 나누세요.")

    elif strategy == "tabular_join":
        # 데이터 파일(피처) + 별도 라벨 파일을 공통 키로 조인한다.
        label_root = clean_path(label.get("데이터 경로"))
        label_patterns = parse_patterns(label.get("파일 패턴")) or ["*.*"]
        data_files, single = _resolve_tabular_set(
            base_dir, data_root, eff_splits, data_patterns, notes, "01_데이터 정보")
        label_files, _ = _resolve_tabular_set(
            base_dir, label_root, eff_splits, label_patterns, notes, "02_라벨 정보")
        lf_by_name = {f["name"]: f for f in label_files}
        join_files = []
        for df in data_files:
            lf = lf_by_name.get(df["name"]) or (label_files[0] if len(label_files) == 1 else None)
            join_files.append({
                "name": df["name"], "sub": df["sub"],
                "data_url": df["url"], "data_mime": df["mime"],
                "label_url": lf["url"] if lf else None,
                "label_mime": lf["mime"] if lf else None,
            })
        layout.update({"join_files": join_files, "single_file": single,
                       "label_root": label_root, "label_patterns": label_patterns})
        if single and any(s.get("sub") for s in eff_splits):
            notes.append("[01/03] 데이터/라벨이 단일 파일이라, 스플릿은 비율(splits)로만 처리합니다.")

    else:  # file_sidecar
        label_root = clean_path(label.get("데이터 경로"))
        label_patterns = parse_patterns(label.get("파일 패턴")) or ["*.*"]
        label_mimes = mimes_for_patterns(label_patterns)
        layout.update({
            "label_root": label_root, "label_patterns": label_patterns,
            "label_mimes": label_mimes, "label_fmt": label.get("라벨 파일 형식"),
            "has_label_files": bool(label_root),
        })
        if mapping == "index":
            notes.append("[02_라벨 정보] 매핑이 'index'(저장 순서 기반)입니다. 데이터/라벨을 "
                         "재정렬하면 매핑이 깨지므로, 로더에서 정렬 기준을 고정하세요(재현성 주의).")
        elif mapping == "" :
            notes.append("[02_라벨 정보] '데이터-라벨 매핑 기준'이 비어 있어 파일명(filename) 기반으로 "
                         "간주했습니다. column/filename/id/key/index 중 하나를 지정하세요.")
    return layout


# ---------------------------------------------------------------------------
# 4) Croissant 구성요소 빌드
# ---------------------------------------------------------------------------
def _dir_fileobject(fid, content_url, desc):
    """로컬 폴더를 가리키는 디렉토리 FileObject. FileSet 의 containedIn 대상이 된다."""
    return {
        "@type": "cr:FileObject", "@id": fid, "name": fid,
        "description": desc,
        "contentUrl": content_url, "encodingFormat": DIRECTORY_MIME,
    }


def _fileset(fid, root_id, includes, mimes, desc):
    """디렉토리(root_id) 안에서 glob 으로 파일을 모으는 FileSet.

    includes 는 root 기준 상대 glob 리스트(예: ['train/*.jpg', 'train/*.png']).
    encodingFormat 은 여러 개 올 수 있다(예: image/jpeg, image/png).
    """
    return {
        "@type": "cr:FileSet", "@id": fid, "name": fid,
        "description": desc,
        "containedIn": {"@id": root_id},
        "includes": includes if len(includes) > 1 else includes[0],
        "encodingFormat": mimes if len(mimes) > 1 else mimes[0],
    }


def build_distribution(layout):
    """layout 에 맞는 distribution(파일/파일셋 목록)을 만든다."""
    distribution = []

    if layout["strategy"] == "tabular_embedded":
        # 스플릿마다 표 파일 FileObject 하나 — 라벨은 그 파일의 컬럼에서 추출한다.
        single = layout["single_file"]
        for f in layout["tabular_files"]:
            fid = "data-file" if single else f"data-{f['name']}"
            desc = "데이터 파일(라벨 컬럼 포함)" if single else f"{f['name']} 데이터 파일(라벨 컬럼 포함)"
            distribution.append({
                "@type": "cr:FileObject", "@id": fid, "name": fid,
                "description": desc,
                "contentUrl": f["url"], "encodingFormat": f["mime"],
            })
        return distribution

    if layout["strategy"] == "tabular_join":
        # 스플릿마다 데이터 파일 + 라벨 파일 2개의 FileObject. 조인은 recordSet 에서.
        single = layout["single_file"]
        for f in layout["join_files"]:
            did = "data-file" if single else f"data-{f['name']}"
            lid = "labels-file" if single else f"labels-{f['name']}"
            distribution.append({
                "@type": "cr:FileObject", "@id": did, "name": did,
                "description": "데이터 파일(피처)" if single else f"{f['name']} 데이터 파일(피처)",
                "contentUrl": f["data_url"], "encodingFormat": f["data_mime"],
            })
            if f["label_url"]:
                distribution.append({
                    "@type": "cr:FileObject", "@id": lid, "name": lid,
                    "description": "라벨 파일" if single else f"{f['name']} 라벨 파일",
                    "contentUrl": f["label_url"], "encodingFormat": f["label_mime"],
                })
        return distribution

    # ---- file_sidecar ----
    # 스플릿마다 "전용" 디렉토리 FileObject(컨테이너)를 만들고 그 안에서 glob 한다.
    # (하나의 컨테이너를 여러 FileSet 이 공유하면 mlcroissant 연산 그래프에 사이클이
    #  생겨 두 번째 스플릿 로딩이 실패하므로, 스플릿별로 컨테이너를 분리한다.)
    data_root, label_root = layout["data_root"], layout.get("label_root")

    for s in layout["eff_splits"]:
        sub = s["sub"]
        # 데이터: 스플릿 폴더 컨테이너 + 그 안의 FileSet(패턴은 폴더 기준 상대 glob)
        d_dir = f"{data_root}/{sub}" if sub else data_root
        d_root_id = f"data-root-{s['name']}"
        distribution.append(_dir_fileobject(d_root_id, d_dir, f"{s['name']} 데이터 폴더"))
        distribution.append(_fileset(
            f"data-{s['name']}", d_root_id, layout["data_patterns"], layout["data_mimes"],
            f"{s['name']} 데이터 파일"))
        # 라벨: 별도 라벨 파일이 있을 때만
        if layout["has_label_files"]:
            l_dir = f"{label_root}/{sub}" if sub else label_root
            l_root_id = f"labels-root-{s['name']}"
            distribution.append(_dir_fileobject(l_root_id, l_dir, f"{s['name']} 라벨 폴더"))
            distribution.append(_fileset(
                f"labels-{s['name']}", l_root_id, layout["label_patterns"], layout["label_mimes"],
                f"{s['name']} 라벨 파일 ({layout.get('label_fmt')})"))
    return distribution


def _enum_recordsets(classes, splits, classes_dtype):
    """클래스/스플릿 enumeration RecordSet(공통)."""
    record_sets = []
    if classes:
        record_sets.append({
            "@type": "cr:RecordSet", "@id": "classes", "name": "classes",
            "description": "클래스 정의", "dataType": "sc:Enumeration",
            "key": {"@id": "classes/id"},
            "field": [
                {"@type": "cr:Field", "@id": "classes/id", "name": "id",
                 "description": "클래스 인덱스", "dataType": classes_dtype},
                {"@type": "cr:Field", "@id": "classes/name", "name": "name",
                 "description": "클래스 이름", "dataType": "sc:Text"},
            ],
            "data": [{"classes/id": i, "classes/name": n} for i, n in classes],
        })
    if splits:
        record_sets.append({
            "@type": "cr:RecordSet", "@id": "splits", "name": "splits",
            "dataType": "cr:Split", "key": {"@id": "splits/name"},
            "field": [
                {"@type": "cr:Field", "@id": "splits/name", "name": "name", "dataType": "sc:Text"},
                {"@type": "cr:Field", "@id": "splits/ratio", "name": "ratio", "dataType": "sc:Float"},
            ],
            "data": [{"splits/name": s["name"], "splits/ratio": s["ratio"]} for s in splits],
        })
    return record_sets


def _infer_cr_dtype(values):
    """샘플 값에서 Croissant dataType 을 추론한다(범용 규칙).

      - 값이 전부 비어있으면(전결측 컬럼) 숫자 파이프라인 기본값 Float.
      - 결측이 하나도 없고 모두 정수면 Integer.
      - 그 외 모두 숫자면 Float. (결측이 있는 정수열은 Float 로 둔다:
        mlcroissant 가 결측을 None 으로 돌려주므로 Float 가 일관적이다.)
      - 하나라도 숫자가 아니면 Text.
    """
    seen_missing = False
    vals = []
    for v in values:
        if is_empty(v):
            seen_missing = True
        else:
            vals.append(str(v).strip())
    if not vals:
        return "sc:Float"
    all_int = True
    for s in vals:
        try:
            int(s)
        except ValueError:
            all_int = False
            break
    if all_int and not seen_missing:
        return "sc:Integer"
    all_float = True
    for s in vals:
        try:
            float(s)
        except ValueError:
            all_float = False
            break
    return "sc:Float" if all_float else "sc:Text"


def _sniff_delimiter(lines, ext):
    """샘플 라인에서 구분자를 추정. 실패 시 확장자 기본값(csv=',', tsv='\\t')."""
    sample = "\n".join(lines[:10])
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except Exception:
        return "\t" if ext == "tsv" else ","


def _read_columns_and_dtypes(path, sample_rows=300):
    """임의의 표 파일에서 (컬럼명, dataType, 진단정보) 를 읽는다.

    범용 처리:
      - 인코딩: utf-8(BOM 제거) 우선, 실패 시 latin-1 로 대체.
      - 구분자: csv.Sniffer 로 , ; \\t | 자동 추정(확장자 기본값 폴백).
      - csv/tsv 는 표준 csv 모듈, parquet 는 pandas(있을 때).
    반환 diag: {delimiter, encoding_note, non_comma, duplicates, blank_count}
    """
    ext = str(path).rsplit(".", 1)[-1].lower()
    diag = {"delimiter": ",", "encoding_note": None, "non_comma": False,
            "duplicates": [], "blank_count": 0}

    if ext == "parquet":
        try:
            import pandas as pd
            df = pd.read_parquet(path)
            header = [str(c) for c in df.columns]
            dtypes = []
            for c in df.columns:
                k = df[c].dtype.kind
                dtypes.append("sc:Integer" if k in "iu"
                              else "sc:Float" if k == "f" else "sc:Text")
            return header, dtypes, diag
        except Exception:
            return None, None, diag

    # ----- csv / tsv (그 외 확장자도 텍스트 표로 시도) -----
    try:
        with open(path, "rb") as fh:
            raw = fh.read()
    except OSError:
        return None, None, diag
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
        diag["encoding_note"] = "utf-8 디코딩 실패로 latin-1 로 읽었습니다(원본 인코딩 확인 권장)."

    lines = text.splitlines()
    if not lines:
        return None, None, diag
    delim = _sniff_delimiter(lines, ext)
    diag["delimiter"] = delim
    diag["non_comma"] = (delim != ",")

    reader = csv.reader(lines, delimiter=delim)
    header = next(reader, None)
    if not header:
        return None, None, diag

    samples = [[] for _ in header]
    for i, row in enumerate(reader):
        if i >= sample_rows:
            break
        for j in range(len(header)):
            samples[j].append(row[j] if j < len(row) else "")
    dtypes = [_infer_cr_dtype(col) for col in samples]

    # 진단: 빈 컬럼명 / 중복 컬럼명
    diag["blank_count"] = sum(1 for h in header if is_empty(h))
    seen, dups = set(), []
    for h in header:
        if is_empty(h):
            continue
        if h in seen and h not in dups:
            dups.append(h)
        seen.add(h)
    diag["duplicates"] = dups
    return header, dtypes, diag


def _scan_table_facts(path, delimiter, header, label_col, feature_cols):
    """실제 표 파일을 1회 전체 스캔해 검증용 '사실'을 모은다.

    반환 dict:
      n_rows        : 데이터 행 수(헤더 제외)
      ragged_rows   : 헤더와 컬럼 수가 다른 행 수(깨진 CSV 신호)
      label_dist    : {라벨값(str): 개수}  (라벨 컬럼이 있을 때만)
      label_missing : 라벨이 빈 행 수
      feat_allnan   : 값이 전부 빈 피처 컬럼명 리스트
      feat_const    : 값이 하나뿐(상수)인 피처 컬럼명 리스트
    라벨/피처 컬럼명이 헤더에 없으면 그 부분 통계는 건너뛴다.
    """
    facts = {"n_rows": 0, "ragged_rows": 0, "label_dist": {}, "label_missing": 0,
             "feat_allnan": [], "feat_const": []}
    ncol = len(header)
    label_idx = header.index(label_col) if label_col in header else None
    feat_idx = [(c, header.index(c)) for c in (feature_cols or []) if c in header]
    # 피처별 상태: nonempty(비어있지 않은 값을 본 적), first(첫 비어있지 않은 값), const(계속 같은가)
    fstate = {c: {"nonempty": False, "first": None, "const": True} for c, _ in feat_idx}
    dist = {}
    try:
        with open(path, "rb") as fh:
            text = fh.read().decode("utf-8-sig", errors="replace")
    except OSError:
        return facts
    reader = csv.reader(text.splitlines(), delimiter=delimiter)
    next(reader, None)  # 헤더 스킵
    for row in reader:
        if not row:
            continue
        facts["n_rows"] += 1
        if len(row) != ncol:
            facts["ragged_rows"] += 1
        if label_idx is not None:
            v = row[label_idx].strip() if label_idx < len(row) else ""
            if v == "":
                facts["label_missing"] += 1
            else:
                dist[v] = dist.get(v, 0) + 1
        for c, j in feat_idx:
            val = row[j].strip() if j < len(row) else ""
            if val == "":
                continue
            st = fstate[c]
            if not st["nonempty"]:
                st["nonempty"] = True
                st["first"] = val
            elif st["const"] and val != st["first"]:
                st["const"] = False
    facts["label_dist"] = dist
    facts["feat_allnan"] = [c for c, _ in feat_idx if not fstate[c]["nonempty"]]
    facts["feat_const"] = [c for c, _ in feat_idx
                           if fstate[c]["nonempty"] and st_const(fstate[c])]
    return facts


def st_const(state):
    return state["const"]


def _detect_join_key(data_header, label_header):
    """데이터/라벨 헤더의 공통 컬럼으로 조인 키를 자동 탐지한다.

    반환: (key 또는 None, error 또는 None, common 리스트)
      - 공통이 정확히 1개 → 그 컬럼이 키.
      - 0개 → 조인할 키가 없음(error).
      - 2개 이상 → 어느 걸 키로 쓸지 모호(error, 후보 나열).
    """
    common = [c for c in data_header if c in set(label_header) and not is_empty(c)]
    # 중복 제거(순서 유지)
    seen, uniq = set(), []
    for c in common:
        if c not in seen:
            seen.add(c)
            uniq.append(c)
    if len(uniq) == 1:
        return uniq[0], None, uniq
    if not uniq:
        return None, ("데이터 파일과 라벨 파일에 공통 컬럼이 없어 조인 키를 찾지 못했습니다. "
                      "두 파일이 공유하는 식별자 컬럼(같은 이름)을 두세요.", []), uniq
    return None, (f"공통 컬럼이 여러 개라 조인 키가 모호합니다: {', '.join(uniq)}. "
                  "공유 식별자를 하나만 남기거나 키를 지정하세요.", uniq), uniq


def _read_key_column(path, delimiter, colname):
    """지정 컬럼(키)의 값 리스트를 전체 스캔으로 읽는다. 실패 시 []."""
    try:
        with open(path, "rb") as fh:
            text = fh.read().decode("utf-8-sig", errors="replace")
    except OSError:
        return []
    reader = csv.reader(text.splitlines(), delimiter=delimiter)
    header = next(reader, None)
    if not header or colname not in header:
        return []
    j = header.index(colname)
    return [(row[j].strip() if j < len(row) else "") for row in reader if row]


def _safe_field_name(col, used):
    """컬럼명을 Field @id/name 으로 쓸 수 있게 정규화하고 중복을 피한다."""
    base = re.sub(r"[^0-9A-Za-z_]+", "_", str(col)).strip("_") or "col"
    if base[0].isdigit():
        base = "f_" + base          # '0' -> 'f_0' (숫자로 시작 방지)
    name, k = base, 2
    while name in used:
        name = f"{base}_{k}"
        k += 1
    used.add(name)
    return name


def _feature_fields(rid, fid, columns, dtypes, label_col, include_cols=None):
    """피처 컬럼을 추출 Field 로 만든다. 범용성 처리 포함.

      - include_cols 가 주어지면 그 컬럼(그 순서)만 피처로 명세한다
        (엑셀 '입력(Feature) 컬럼' 셀 참조로 지정된 컬럼).
      - include_cols 가 None 이면 라벨을 제외한 모든 컬럼을 피처로 명세한다(하위호환).
      - 공통: 무명 컬럼·중복 컬럼·라벨 컬럼은 제외. source.extract.column 에는
        원본 컬럼명 그대로, @id/name 만 정규화.
    """
    dtype_map = {c: d for c, d in zip(columns, dtypes)}
    if include_cols is not None:
        seq = list(include_cols)                       # 명세된 컬럼만(순서 유지)
    else:
        seq = [c for c in columns if str(c) != str(label_col)]

    fields, used, used_cols = [], set(), set()
    for col in seq:
        if is_empty(col) or str(col) == str(label_col):
            continue
        if col in used_cols:                           # 중복 컬럼명 → 첫 등장만
            continue
        if col not in dtype_map:                       # 헤더에 없는 이름은 건너뜀(방어)
            continue
        used_cols.add(col)
        name = _safe_field_name(col, used)
        fields.append({
            "@type": "cr:Field", "@id": f"{rid}/{name}", "name": name,
            "description": f"피처 컬럼 '{col}'",
            "dataType": dtype_map[col],
            "source": {"fileObject": {"@id": fid}, "extract": {"column": str(col)}},
        })
    return fields


def build_record_sets(classes, splits, layout):
    """전략에 맞는 RecordSet 목록 생성.

    - tabular_embedded: 데이터 파일에서 '라벨 컬럼'을 추출하는 records RecordSet.
    - file_sidecar    : 스플릿별로 (데이터 파일명[+이미지 내용]) / (라벨 파일명+내용) RecordSet.
    """
    classes_dtype = "sc:Integer" if (classes and all(isinstance(c, int) for c, _ in classes)) else "sc:Text"
    record_sets = _enum_recordsets(classes, splits, classes_dtype)

    if layout["strategy"] == "tabular_embedded":
        label_col = layout["label_column"] or "label"
        # 라벨 dtype: 클래스 id 가 모두 정수면 Integer, 아니면 Text.
        label_dtype = "sc:Integer" if (classes and all(isinstance(c, int) for c, _ in classes)) else "sc:Text"
        single = layout["single_file"]
        for f in layout["tabular_files"]:
            fid = "data-file" if single else f"data-{f['name']}"
            rid = "records" if single else f"records-{f['name']}"
            field = {
                "@type": "cr:Field", "@id": f"{rid}/label", "name": "label",
                "description": f"라벨(데이터 컬럼 '{label_col}')",
                "dataType": label_dtype,
                "source": {"fileObject": {"@id": fid}, "extract": {"column": label_col}},
            }
            # 클래스 enumeration 이 있으면 라벨이 그 클래스를 참조함을 명시(선택).
            if classes:
                field["references"] = {"field": {"@id": "classes/id"}}
            head = "데이터 레코드" if single else f"{f['name']} 데이터 레코드"

            fields = [field]
            # --- 피처 컬럼 명세: 실제 표 파일의 헤더를 읽어 라벨 외 모든 컬럼을 Field 로 추가 ---
            base_dir = layout.get("base_dir")
            path = os.path.join(base_dir, f["url"]) if base_dir else f["url"]
            resolved = layout.setdefault("resolved_columns", {})
            if os.path.isfile(path):
                columns, dtypes, diag = _read_columns_and_dtypes(path)
                resolved[fid] = {"columns": columns,
                                 "label_present": bool(columns) and (label_col in columns),
                                 "url": f["url"]}
                if columns:
                    # 엑셀 '입력(Feature) 컬럼' 셀 참조가 있으면 그 컬럼만, 없으면 라벨 외 전체.
                    # 잘못된 참조는 여기서 '보정'하지 않고 오류로 모아 검증 단계에서 반려한다.
                    spec = layout.get("feature_spec")
                    if spec:
                        include_cols, rerr = resolve_feature_columns(
                            layout.get("feature_idxs") or [], columns, label_col)
                        layout["feature_resolve_errors"].extend(f"[{fid}] {e}" for e in rerr)
                    else:
                        include_cols = None
                    feats = _feature_fields(rid, fid, columns, dtypes, label_col,
                                            include_cols=include_cols)
                    fields.extend(feats)
                    # --- 실제 데이터 기반 검증용 사실 수집(1회 전체 스캔) ---
                    feat_names = (list(include_cols) if include_cols is not None
                                  else [c for c in columns if str(c) != str(label_col)])
                    facts = _scan_table_facts(path, diag.get("delimiter", ","),
                                              columns, label_col, feat_names)
                    resolved[fid].update({
                        "delimiter": diag.get("delimiter", ","),
                        "ext": ext_of(f["url"]), "feature_cols": feat_names,
                        "duplicates": diag.get("duplicates", []),
                        "blank_count": diag.get("blank_count", 0),
                        **facts,
                    })
                    src = "엑셀 셀 참조로 지정된" if include_cols is not None else "라벨 외 전체"
                    desc = (f"{head}. 라벨('{label_col}')과 피처 {len(feats)}개 컬럼"
                            f"({src})을 같은 행에서 추출한다.")
                    # 범용성 진단 경고
                    if diag.get("encoding_note"):
                        layout["notes"].append(f"[{fid}] " + diag["encoding_note"])
                    if diag.get("non_comma"):
                        layout["notes"].append(
                            f"[{fid}] 구분자가 콤마가 아닙니다(추정: '{diag['delimiter']}'). "
                            "로더/도구가 같은 구분자로 읽는지 확인하세요.")
                    if diag.get("blank_count"):
                        layout["notes"].append(
                            f"[{fid}] 이름이 빈 컬럼 {diag['blank_count']}개는 이름으로 추출할 수 "
                            "없어 제외했습니다(예: 인덱스 컬럼).")
                    if diag.get("duplicates"):
                        layout["notes"].append(
                            f"[{fid}] 중복 컬럼명 {diag['duplicates']} 은 첫 등장만 명세했습니다"
                            "(이름 기반 추출은 중복을 구분할 수 없음). 헤더를 유일하게 만드는 것을 권장.")
                else:
                    desc = head + ". (표 파일을 읽었으나 컬럼을 인식하지 못해 라벨만 명세함)"
                    layout["notes"].append(
                        f"[{fid}] '{f['url']}' 의 컬럼을 읽지 못해 피처를 명세하지 못했습니다"
                        "(라벨만 생성). 파일 형식을 확인하세요.")
            else:
                resolved[fid] = {"columns": None, "label_present": None, "url": f["url"]}
                desc = head + ". (실제 데이터 파일이 없어 라벨만 명세함)"
                layout["notes"].append(
                    f"[{fid}] 실제 표 파일('{f['url']}')이 없어 피처 컬럼을 명세하지 못했습니다"
                    "(라벨만 생성). 데이터가 준비된 위치에서 다시 생성하세요.")

            record_sets.append({
                "@type": "cr:RecordSet", "@id": rid, "name": rid,
                "description": desc,
                "field": fields,
            })
        return record_sets

    if layout["strategy"] == "tabular_join":
        label_col = layout["label_column"] or "label"
        label_dtype = "sc:Integer" if (classes and all(isinstance(c, int) for c, _ in classes)) else "sc:Text"
        single = layout["single_file"]
        base_dir = layout.get("base_dir")
        jinfo = layout.setdefault("join_info", {})
        for f in layout["join_files"]:
            did = "data-file" if single else f"data-{f['name']}"
            lid = "labels-file" if single else f"labels-{f['name']}"
            rid = "records" if single else f"records-{f['name']}"
            dpath = os.path.join(base_dir, f["data_url"]) if base_dir and f["data_url"] else f["data_url"]
            lpath = os.path.join(base_dir, f["label_url"]) if base_dir and f["label_url"] else f["label_url"]

            # 두 파일이 모두 있어야 조인 명세가 가능하다.
            if not (f["data_url"] and f["label_url"] and dpath and lpath
                    and os.path.isfile(dpath) and os.path.isfile(lpath)):
                jinfo[rid] = {"ok": False, "reason": "missing_file",
                              "data_url": f["data_url"], "label_url": f["label_url"]}
                record_sets.append({
                    "@type": "cr:RecordSet", "@id": rid, "name": rid,
                    "description": "데이터/라벨 파일이 없어 조인을 명세하지 못했습니다.",
                    "field": [],
                })
                continue

            dcols, ddtypes, ddiag = _read_columns_and_dtypes(dpath)
            lcols, _, ldiag = _read_columns_and_dtypes(lpath)
            key, kerr, common = _detect_join_key(dcols or [], lcols or [])
            jinfo[rid] = {"ok": key is not None, "key": key, "kerr": kerr,
                          "data_url": f["data_url"], "label_url": f["label_url"],
                          "dpath": dpath, "lpath": lpath, "ddelim": ddiag.get("delimiter", ","),
                          "ldelim": ldiag.get("delimiter", ","),
                          "dcols": dcols, "lcols": lcols, "label_present": bool(lcols) and (label_col in lcols)}

            fields = []
            if key:
                # (1) 키 필드: 데이터 키가 라벨 키를 참조 → 조인 선언(검증된 문법).
                kname = _safe_field_name(key, set())
                fields.append({
                    "@type": "cr:Field", "@id": f"{rid}/{kname}", "name": kname,
                    "description": f"조인 키 '{key}'",
                    "dataType": "sc:Text",
                    "source": {"fileObject": {"@id": did}, "extract": {"column": str(key)}},
                    "references": {"fileObject": {"@id": lid}, "extract": {"column": str(key)}},
                })
                # (2) 피처 필드: 데이터 파일에서. 키/라벨 컬럼은 피처에서 제외.
                spec = layout.get("feature_spec")
                if spec:
                    include_cols, rerr = resolve_feature_columns(
                        layout.get("feature_idxs") or [], dcols, label_col)
                    layout["feature_resolve_errors"].extend(f"[{did}] {e}" for e in rerr)
                else:
                    include_cols = None
                feats = _feature_fields(rid, did, dcols, ddtypes, label_col, include_cols=include_cols)
                feats = [ff for ff in feats if ff["source"]["extract"]["column"] != str(key)]
                fields.extend(feats)
                # (3) 라벨 필드: 라벨 파일에서.
                fields.append({
                    "@type": "cr:Field", "@id": f"{rid}/label", "name": "label",
                    "description": f"라벨(라벨 파일 컬럼 '{label_col}')",
                    "dataType": label_dtype,
                    "source": {"fileObject": {"@id": lid}, "extract": {"column": label_col}},
                })
                desc = (f"데이터 레코드. 데이터 파일의 피처 {len(feats)}개와 "
                        f"라벨('{label_col}')을 공통 키 '{key}'로 조인해 추출한다.")
            else:
                desc = "조인 키를 찾지 못해 필드를 명세하지 못했습니다(검증 단계에서 반려)."

            record_sets.append({
                "@type": "cr:RecordSet", "@id": rid, "name": rid,
                "description": desc, "field": fields,
            })
        return record_sets

    # ---- file_sidecar ----
    data_is_image = any(m in IMAGE_MIMES for m in layout["data_mimes"])
    for s in layout["eff_splits"]:
        rid = f"data-records-{s['name']}"
        d_fs = f"data-{s['name']}"
        fields = []
        if data_is_image:
            fields.append({
                "@type": "cr:Field", "@id": f"{rid}/image", "name": "image",
                "description": "이미지", "dataType": "sc:ImageObject",
                "source": {"fileSet": {"@id": d_fs}, "extract": {"fileProperty": "content"}}})
        fields.append({
            "@type": "cr:Field", "@id": f"{rid}/filename", "name": "filename",
            "description": "데이터 파일명(라벨 매핑 키)", "dataType": "sc:Text",
            "source": {"fileSet": {"@id": d_fs}, "extract": {"fileProperty": "filename"}}})
        record_sets.append({
            "@type": "cr:RecordSet", "@id": rid, "name": rid,
            "description": f"{s['name']} 데이터 예시",
            "field": fields,
        })
        # 라벨 파일이 있으면 파일명+내용을 그대로 내주는 RecordSet(로더가 파일명으로 매핑).
        if layout["has_label_files"]:
            lrid = f"labels-records-{s['name']}"
            l_fs = f"labels-{s['name']}"
            record_sets.append({
                "@type": "cr:RecordSet", "@id": lrid, "name": lrid,
                "description": f"{s['name']} 라벨 파일(파일명+내용)",
                "field": [
                    {"@type": "cr:Field", "@id": f"{lrid}/filename", "name": "filename",
                     "description": "라벨 파일명(데이터 매핑 키)", "dataType": "sc:Text",
                     "source": {"fileSet": {"@id": l_fs}, "extract": {"fileProperty": "filename"}}},
                    {"@type": "cr:Field", "@id": f"{lrid}/content", "name": "content",
                     "description": "라벨 파일 내용(원본 텍스트)", "dataType": "sc:Text",
                     "source": {"fileSet": {"@id": l_fs}, "extract": {"fileProperty": "content"}}},
                ],
            })
    return record_sets


def build_citation(base):
    """엑셀엔 없는 citeAs 를 기존 값으로 BibTeX 형태 자동 생성."""
    pub = iso(base.get("생성일"))
    year = pub[:4] if pub else ""
    return (
        f"@misc{{{base.get('데이터셋 ID')},\n"
        f"  title  = {{{base.get('데이터셋 명')}}},\n"
        f"  author = {{{base.get('제출 기업')}}},\n"
        f"  year   = {{{year}}},\n"
        f"  note   = {{{base.get('데이터셋 ID')}, version {base.get('버전')}}}\n"
        f"}}"
    )


def build_metadata(sheets, distribution, record_sets):
    """최상위 Croissant 메타데이터 dict 생성(None/빈값 제거)."""
    base = sheets["base"]
    keywords = [k for k in [base.get("대상 업종"), base.get("대상 공정"),
                            base.get("데이터 유형"), base.get("AI Task 유형")] if k]
    ds_id = base.get("데이터셋 ID")
    metadata = {
        "@context": CANONICAL_CONTEXT,
        "@type": "sc:Dataset",
        "conformsTo": "http://mlcommons.org/croissant/1.1",
        "name": ds_id,
        "alternateName": base.get("데이터셋 명"),
        "description": base.get("설명"),
        "citeAs": build_citation(base),
        "version": str(base.get("버전")) if not is_empty(base.get("버전")) else None,
        "datePublished": iso(base.get("생성일")),
        "keywords": keywords,
        "license": LICENSE_URL,
        "creator": {"@type": "sc:Organization", "name": base.get("제출 기업")}
        if not is_empty(base.get("제출 기업")) else None,
        # url: 실제 배포 URL(URL_TEMPLATE)이 설정된 경우에만 넣는다. 없으면 생성하지 않음.
        "url": URL_TEMPLATE.format(ds_id=ds_id) if (URL_TEMPLATE and not is_empty(ds_id)) else None,
        # 체크섬(md5/sha256)을 명세에 담지 않으므로, 해시 없는 FileObject 도 통과하도록 라이브로 표시.
        "isLiveDataset": True,
        "distribution": distribution,
        "recordSet": record_sets,
    }
    return {k: v for k, v in metadata.items() if v not in (None, "", [], {})}


# ---------------------------------------------------------------------------
# 5) 검증
# ---------------------------------------------------------------------------
def _fileobject_urls(distribution):
    """distribution 에서 (@id -> contentUrl) 매핑(디렉토리/파일 FileObject)."""
    return {d["@id"]: d.get("contentUrl") for d in distribution
            if d["@type"] == "cr:FileObject"}


def _iter_glob_specs(distribution):
    """검증용: 각 항목이 실제로 어떤 파일들을 가리키는지 (id, [glob들], is_dir) 로.

    - FileObject(파일/디렉토리): contentUrl 자체.
    - FileSet: containedIn 디렉토리의 contentUrl + 각 include.
    """
    urls = _fileobject_urls(distribution)
    specs = []
    for d in distribution:
        if d["@type"] == "cr:FileObject":
            is_dir = d.get("encodingFormat") == DIRECTORY_MIME
            specs.append((d["@id"], [d.get("contentUrl")], is_dir))
        elif d["@type"] == "cr:FileSet":
            root_id = (d.get("containedIn") or {}).get("@id")
            base = urls.get(root_id, "")
            inc = d.get("includes")
            inc = inc if isinstance(inc, list) else [inc]
            globs = [f"{base}/{i}" if base else i for i in inc]
            specs.append((d["@id"], globs, False))
    return specs


def _stem(path):
    return os.path.splitext(os.path.basename(path))[0]


def _glob_split_files(base_dir, root, sub, patterns):
    """base_dir/root[/sub] 아래에서 패턴에 맞는 실제 파일 목록(상대경로)."""
    d = os.path.join(base_dir, root, sub) if sub else os.path.join(base_dir, root)
    out = []
    if os.path.isdir(d):
        for pat in patterns:
            out += sorted(glob.glob(os.path.join(d, pat)))
    return out


def validate_sidecar_mapping(layout, base_dir):
    """file_sidecar(파일 기반) 매핑이 실제로 짝을 짓는지 검증한다(1층 정적 매핑).

      - filename: 데이터 파일 ↔ 라벨 파일을 파일명(stem)으로 매칭. 고아(짝 없음)는 반려.
      - index   : 데이터/라벨 파일 개수가 어긋나면 순서 매칭이 전부 밀리므로 반려.
      - id/key  : 조인 키 컬럼 정보가 템플릿에 없어 정적 검증 불가 → 경고.
    반환: (errors, warns)
    """
    errors, warns, infos = [], [], []
    if layout["strategy"] != "file_sidecar" or not base_dir:
        return errors, warns, infos
    mapping = layout["mapping"] or "filename"
    if not layout.get("has_label_files"):
        warns.append("[02_라벨 정보] 별도 라벨 파일 경로가 없어 파일 기반 매핑을 검증할 수 없습니다.")
        return errors, warns, infos

    for s in layout["eff_splits"]:
        sub = s["sub"]
        dfiles = _glob_split_files(base_dir, layout["data_root"], sub, layout["data_patterns"])
        lfiles = _glob_split_files(base_dir, layout["label_root"], sub, layout["label_patterns"])
        tag = f"[{s['name']}] " if s["name"] != "all" else ""

        if mapping in ("filename", ""):
            dmap, lmap = {_stem(p): p for p in dfiles}, {_stem(p): p for p in lfiles}
            data_wo_label = sorted(set(dmap) - set(lmap))
            label_wo_data = sorted(set(lmap) - set(dmap))
            if data_wo_label:
                shown = ", ".join(data_wo_label[:5]) + (f" 외 {len(data_wo_label)-5}개" if len(data_wo_label) > 5 else "")
                errors.append(f"{tag}데이터에 대응하는 라벨 파일이 없습니다: {shown}. "
                              "→ 파일명(stem)이 일치하는 라벨 파일을 두거나 매핑을 확인하세요.")
            if label_wo_data:
                shown = ", ".join(label_wo_data[:5]) + (f" 외 {len(label_wo_data)-5}개" if len(label_wo_data) > 5 else "")
                errors.append(f"{tag}대응 데이터가 없는 라벨 파일이 있습니다: {shown}. "
                              "→ 라벨이 유실되지 않도록 데이터/매핑을 확인하세요.")
            if dfiles and lfiles and not data_wo_label and not label_wo_data:
                infos.append(f"{tag}데이터-라벨 파일명 매칭 OK ({len(dfiles)}쌍).")
        elif mapping == "index":
            if len(dfiles) != len(lfiles):
                errors.append(f"{tag}index(순서) 매핑인데 데이터 파일 {len(dfiles)}개, "
                              f"라벨 파일 {len(lfiles)}개로 개수가 다릅니다. → 순서 매칭이 어긋납니다.")
            else:
                infos.append(f"{tag}index 매핑: 데이터/라벨 개수 일치({len(dfiles)}).")
                warns.append(f"{tag}index 매핑은 순서 의존이므로 로더 정렬 기준을 고정하세요(재현성 주의).")
    return errors, warns, infos


def validate_join_mapping(layout, label, classes, skip_verify):
    """tabular_join(id/key) 매핑의 조인 정합성을 검증한다(1층 정적 매핑).

      - 키 자동 탐지 실패(공통 0개/모호 2개+) → 반려.
      - 라벨 컬럼이 라벨 파일에 없음 → 반려.
      - (verify) 라벨 없는 데이터 키(조인 시 None) → 반려.
      - (verify) 라벨 파일 키 중복(조인 팬아웃으로 행 복제) → 반려.
      - (verify) 데이터 없는 라벨 키/데이터 키 중복 → 경고.
      - (verify) 라벨 값 ↔ 선언 클래스 불일치 → 반려.
    """
    import collections
    errors, warns, infos = [], [], []
    if layout["strategy"] != "tabular_join":
        return errors, warns, infos
    label_col = layout["label_column"]
    declared = {str(c) for c, _ in classes} if classes else set()

    for rid, ji in (layout.get("join_info") or {}).items():
        if ji.get("reason") == "missing_file":
            errors.append(f"[{rid}] 데이터/라벨 파일을 찾지 못해 조인을 만들 수 없습니다 "
                          f"(data='{ji.get('data_url')}', labels='{ji.get('label_url')}').")
            continue
        if ji.get("kerr"):
            errors.append(f"[{rid}] 조인 키 자동 탐지 실패: {ji['kerr'][0]}")
            continue
        key = ji["key"]
        if not ji.get("label_present"):
            errors.append(f"[{rid}] 라벨 컬럼 '{label_col}' 이(가) 라벨 파일 헤더에 없습니다.")
        infos.append(f"[{rid}] 조인 키 자동 탐지: '{key}' (데이터/라벨 공통 컬럼).")
        if skip_verify:
            continue

        dk = _read_key_column(ji["dpath"], ji["ddelim"], key)
        lk = _read_key_column(ji["lpath"], ji["ldelim"], key)
        dset, lset = set(dk), set(lk)
        orphan_data = sorted(dset - lset)
        orphan_label = sorted(lset - dset)
        dup_label = [k for k, c in collections.Counter(lk).items() if c > 1]
        dup_data = [k for k, c in collections.Counter(dk).items() if c > 1]

        def _shown(xs):
            return ", ".join(map(str, xs[:5])) + (f" 외 {len(xs)-5}개" if len(xs) > 5 else "")

        if orphan_data:
            errors.append(f"[{rid}] 라벨이 없는 데이터 키 {len(orphan_data)}개 — 조인 시 라벨이 "
                          f"비어(None) 유실됩니다. 예: {_shown(orphan_data)}")
        if dup_label:
            errors.append(f"[{rid}] 라벨 파일에서 키가 중복됩니다 {len(dup_label)}개 — 조인 팬아웃으로 "
                          f"데이터 행이 복제됩니다. 예: {_shown(dup_label)}")
        if orphan_label:
            warns.append(f"[{rid}] 데이터가 없는 라벨 키 {len(orphan_label)}개(사용되지 않는 라벨).")
        if dup_data:
            warns.append(f"[{rid}] 데이터 파일에서 키가 중복됩니다 {len(dup_data)}개.")

        if declared and ji.get("label_present"):
            lvals = {v for v in _read_key_column(ji["lpath"], ji["ldelim"], label_col) if v != ""}
            if lvals and lvals != declared:
                only_d = ", ".join(sorted(lvals - declared)) or "-"
                only_c = ", ".join(sorted(declared - lvals)) or "-"
                errors.append(f"[{rid}] 라벨 파일의 실제 값이 선언 클래스와 다릅니다. "
                              f"데이터에만: {{{only_d}}} / 선언에만: {{{only_c}}}.")
    return errors, warns, infos


def validate(sheets, splits, classes, distribution, layout, base_dir, tried, skip_verify):
    """(errors, warns, infos) 반환. errors 가 있으면 파일을 만들면 안 된다.

    - errors : 반려 사유(파일 미생성).
    - warns  : 생성은 되지만 확인 권장(상수/전결측 피처, 비율 합 등).
    - infos  : 참고 정보(라벨 분포, 매칭 OK, 매핑 방식 등) — 문제 아님.
    """
    errors, warns, infos = [], [], []
    base, data_i, label = sheets["base"], sheets["data"], sheets["label"]

    # layout 계획 단계에서 모인 안내는 경고로 올린다.
    warns.extend(layout.get("notes", []))

    # (1) '필수 여부=Y' 인데 '입력' 이 비어 있는 항목
    for sheet_name, vals, reqs in sheets["list"]:
        for key, required in reqs.items():
            if required and is_empty(vals.get(key)):
                errors.append(f"[{sheet_name}] '{key}' 은(는) 필수 항목입니다. → '입력' 열에 값을 채워주세요.")

    # (2) 데이터셋 ID(name) 형식
    ds_id = base.get("데이터셋 ID")
    if not is_empty(ds_id):
        ds_id = str(ds_id).strip()
        if not re.fullmatch(r"[A-Za-z0-9\-_.]+", ds_id):
            errors.append(f"[기본정보] 데이터셋 ID '{ds_id}' 에 사용할 수 없는 문자가 있습니다. "
                          "→ 공백·특수문자 없이 영문/숫자/-/_/. 만 사용하세요 (예: MAX_TR_DS01).")
        if len(ds_id) > 255:
            errors.append("[기본정보] 데이터셋 ID 가 너무 깁니다(255자 초과). → 짧게 줄여주세요.")

    # (3) 파일 패턴 형식(*.확장자, 여러 개 허용)
    for sheet_name, vals in [("01_데이터 정보", data_i), ("02_라벨 정보", label)]:
        raw = vals.get("파일 패턴")
        if is_empty(raw):
            continue
        pats = parse_patterns(raw)
        bad = [p for p in pats if not re.fullmatch(r"\*\.[A-Za-z0-9]+", p)]
        if bad:
            errors.append(f"[{sheet_name}] 파일 패턴 '{raw}' 형식이 올바르지 않습니다"
                          f"(문제: {', '.join(bad)}). "
                          "→ '*.jpg' 처럼 '*.확장자' 형태로, 여러 개면 콤마로 구분하세요 (예: *.jpg, *.png).")

    # (4) 알 수 없는 파일 형식 경고
    for sheet_name, vals in [("01_데이터 정보", data_i), ("02_라벨 정보", label)]:
        for p in parse_patterns(vals.get("파일 패턴")):
            if mime_from_pattern(p) == "application/octet-stream":
                warns.append(f"[{sheet_name}] 파일 패턴 '{p}' 의 확장자를 인식하지 못했습니다. "
                             "→ jpg/png/csv/tsv/txt 등 알려진 확장자인지 확인하세요.")

    # (5) 매핑 기준 값 & 라벨 컬럼(조건부 필수)
    mapping = layout["mapping"]
    if is_empty(mapping):
        warns.append("[02_라벨 정보] '데이터-라벨 매핑 기준'이 비어 있습니다. "
                     "→ column / filename / id/key / index 중 하나를 지정하세요.")
    elif mapping not in ALLOWED_MAPPINGS:
        errors.append(f"[02_라벨 정보] 데이터-라벨 매핑 기준 '{mapping}' 은(는) 허용되지 않는 값입니다. "
                      "→ column / filename / id/key / index 중에서 고르세요.")
    else:
        label_col = layout["label_column"]
        if mapping in ("column", "id/key") and is_empty(label_col):
            errors.append(f"[02_라벨 정보] 매핑이 '{mapping}' 이면 '라벨 컬럼' 이 필수입니다. "
                          "→ 라벨로 쓰는 컬럼명(예: Pass/Fail)을 입력하세요.")
        if mapping in ("filename", "index") and not is_empty(label_col):
            warns.append(f"[02_라벨 정보] 매핑이 '{mapping}'(파일 기반)인데 '라벨 컬럼'({label_col})이 채워져 "
                         "있습니다. → 파일 기반 매핑이면 라벨 컬럼은 보통 공란입니다.")

    # (6) 스플릿 비율(0~1 숫자, 합≈1)
    ratio_sum = 0.0
    for s in splits:
        r = s["ratio"]
        if r is None:
            continue
        try:
            rf = float(r)
        except (TypeError, ValueError):
            errors.append(f"[03_데이터 스플릿 정보] {s['name']} 비율 '{r}' 이(가) 숫자가 아닙니다. "
                          "→ 0~1 사이 숫자로 입력하세요 (예: 0.83).")
            continue
        if not (0 < rf <= 1):
            errors.append(f"[03_데이터 스플릿 정보] {s['name']} 비율 '{r}' 은(는) 0~1 범위를 벗어났습니다. "
                          "→ 0.83 처럼 0과 1 사이 값으로 입력하세요.")
        else:
            ratio_sum += rf
    if splits and abs(ratio_sum - 1.0) > 0.01:
        warns.append(f"[03_데이터 스플릿 정보] 스플릿 비율 합이 {ratio_sum:.2f} 입니다(1.0 권장). "
                     "→ 학습/검증/평가 비율 합이 1이 되도록 확인하세요.")

    # (7) 생성일 형식
    raw_date = base.get("생성일")
    if not is_empty(raw_date) and not isinstance(raw_date, datetime.datetime):
        try:
            datetime.date.fromisoformat(str(raw_date).strip()[:10])
        except ValueError:
            warns.append(f"[기본정보] 생성일 '{raw_date}' 이(가) 날짜 형식이 아닙니다. → YYYY-MM-DD 형태로 입력하세요.")

    # (8) 클래스 수 vs 실제 클래스 개수
    cls_count = label.get("클래스 수")
    if not is_empty(cls_count) and classes:
        try:
            if int(cls_count) != len(classes):
                warns.append(f"[02_라벨 정보] 클래스 수({cls_count})와 실제 클래스 분류 개수({len(classes)})가 다릅니다. "
                             "→ 둘을 일치시켜 주세요.")
        except (TypeError, ValueError):
            warns.append(f"[02_라벨 정보] 클래스 수 '{cls_count}' 이(가) 정수가 아닙니다.")

    # (9) 데이터 파일 실제 존재 여부 — 엑셀의 '데이터 경로 + 스플릿'으로 만든 경로를 검사한다.
    #     기준 폴더(base_dir)는 자동으로 정해지며(-d 없이도 동작), --no-verify 로 건너뛸 수 있다.
    data_root = layout["data_root"]
    if skip_verify:
        warns.append("실제 파일 존재 여부 검사를 건너뛰었습니다(--no-verify). "
                     "→ 데이터가 준비되면 검사를 켜고 다시 확인하세요.")
    elif not (base_dir and data_root and os.path.exists(os.path.join(base_dir, data_root))):
        looked = tried or ([base_dir] if base_dir else [])
        tried_str = ", ".join(f"'{os.path.join(t, data_root)}'" for t in looked) or f"'{data_root}'"
        errors.append(f"엑셀의 데이터 경로 '{data_root}' 을(를) 찾지 못했습니다(확인한 위치: {tried_str}). "
                      "→ 데이터 폴더가 있는 위치에서 실행하거나, -d/--base-dir 로 "
                      "데이터가 들어 있는 상위 폴더를 지정하세요. (데이터 없이 메타데이터만 만들려면 --no-verify)")
    else:
        for fid, globs, is_dir in _iter_glob_specs(distribution):
            found = False
            for g in globs:
                if not g:
                    continue
                full = os.path.join(base_dir, g)
                if is_dir:
                    if os.path.isdir(full):
                        found = True; break
                else:
                    if glob.glob(full, recursive=True):
                        found = True; break
            if not found:
                target = "폴더" if is_dir else "파일"
                errors.append(f"[{fid}] '{', '.join(g for g in globs if g)}' 에 해당하는 {target}이(가) 없습니다. "
                              f"→ '{base_dir}' 아래에 해당 {target}이 있는지, "
                              "데이터 경로/스플릿/패턴이 맞는지 확인하세요.")

    # (10) 라벨 컬럼이 실제 표 파일에 존재하는지(컬럼 기반 매핑 전용).
    #      build_record_sets 가 실제 헤더를 읽어 layout['resolved_columns'] 에 남긴다.
    for fid, info in (layout.get("resolved_columns") or {}).items():
        if info.get("columns") and info.get("label_present") is False:
            errors.append(
                f"[{fid}] 라벨 컬럼 '{layout.get('label_column')}' 이(가) 데이터 파일"
                f"('{info.get('url')}')의 헤더에 없습니다. "
                "→ 02_라벨 정보의 '라벨 컬럼'을 실제 컬럼명과 일치시키세요.")

    # (11) 입력(Feature) 컬럼 셀 참조 — Tabular 은 필수. 비었거나 잘못 기입 시 반려.
    if layout["strategy"] in ("tabular_embedded", "tabular_join"):
        spec = layout.get("feature_spec")
        if not spec:
            errors.append("[01_데이터 정보] 입력(Feature) 컬럼이 비어 있습니다. Tabular 데이터는 "
                          "어떤 컬럼이 입력 피처인지 셀 참조(예: B1:VS1)로 반드시 지정해야 합니다. "
                          "(라벨/ID/시간 등 비피처 컬럼은 범위에서 제외)")
        for e in layout.get("feature_parse_errors", []):
            errors.append(f"[01_데이터 정보] 입력(Feature) 컬럼 '{spec}': {e}")
        for e in layout.get("feature_resolve_errors", []):
            errors.append(f"입력(Feature) 컬럼: {e}")

    # ======================================================================
    # 실제 데이터 기반 검증 (데이터 경로로 파일을 직접 읽어 확인) — verify 시에만.
    # ======================================================================
    if not skip_verify:
        declared_ids = [cid for cid, _ in classes] if classes else []
        for fid, info in (layout.get("resolved_columns") or {}).items():
            if not info.get("columns"):
                continue
            url = info.get("url")

            # (12) 파일/형식
            ext = info.get("ext")
            pat_exts = {ext_of(p) for p in layout.get("data_patterns", [])}
            if ext and pat_exts and ext not in pat_exts:
                errors.append(f"[{fid}] 실제 파일 확장자 '.{ext}' 가 선언한 파일 패턴"
                              f"({', '.join(sorted(pat_exts))})과 다릅니다('{url}').")
            fmt = (layout.get("data_format") or "").lower()
            if fmt and ext and fmt not in (ext, "csv" if ext == "csv" else fmt):
                warns.append(f"[{fid}] 선언한 파일 형식 '{layout.get('data_format')}' 과 "
                             f"실제 확장자 '.{ext}' 가 달라 보입니다. 확인하세요.")
            if info.get("n_rows", 0) == 0:
                errors.append(f"[{fid}] 데이터 행이 없습니다(헤더만 있거나 빈 파일: '{url}').")
            if info.get("ragged_rows", 0):
                errors.append(f"[{fid}] 컬럼 수가 헤더와 다른 행이 {info['ragged_rows']}개 있습니다"
                              f"(깨진 표: '{url}'). → 행마다 열 수가 같아야 합니다.")

            # (13) 라벨 값 ↔ 선언 클래스
            dist = info.get("label_dist", {})
            if info.get("label_present") and declared_ids:
                actual = set(dist.keys())
                declared = {str(c) for c in declared_ids}
                if actual != declared:
                    only_data = ", ".join(sorted(actual - declared)) or "-"
                    only_decl = ", ".join(sorted(declared - actual)) or "-"
                    errors.append(f"[{fid}] 라벨 컬럼의 실제 값이 선언한 클래스와 다릅니다. "
                                  f"데이터에만 있음: {{{only_data}}} / 선언에만 있음: {{{only_decl}}}. "
                                  "→ 클래스 분류와 실제 라벨을 일치시키세요.")
                cls_count = label.get("클래스 수")
                if not is_empty(cls_count):
                    try:
                        if int(cls_count) != len(actual):
                            errors.append(f"[{fid}] 클래스 수(선언 {cls_count}) 와 실제 라벨 고유값 "
                                          f"개수({len(actual)}) 가 다릅니다.")
                    except (TypeError, ValueError):
                        pass
            if info.get("label_missing", 0):
                warns.append(f"[{fid}] 라벨이 비어 있는 행이 {info['label_missing']}개 있습니다.")
            if dist:
                pretty = ", ".join(f"{k}:{v}" for k, v in sorted(dist.items()))
                infos.append(f"[{fid}] 실제 라벨 분포 — {pretty} (총 {info.get('n_rows')}행).")

            # (14) 피처 값
            feat_cols = info.get("feature_cols", [])
            hdr = info.get("columns", [])
            bad_blank = [c for c in feat_cols if is_empty(c)]
            bad_dup = sorted({c for c in feat_cols if hdr.count(c) > 1})
            if bad_blank:
                errors.append(f"[{fid}] 지정된 피처 중 이름이 빈 컬럼이 있습니다"
                              "(이름으로 추출 불가). → 헤더에 컬럼명을 넣으세요.")
            if bad_dup:
                errors.append(f"[{fid}] 지정된 피처 중 중복된 컬럼명이 있습니다: {', '.join(bad_dup)} "
                              "(이름 기반 추출이 모호). → 헤더를 유일하게 만드세요.")
            if info.get("feat_allnan"):
                shown = ", ".join(info["feat_allnan"][:5]) + (f" 외 {len(info['feat_allnan'])-5}개" if len(info["feat_allnan"]) > 5 else "")
                warns.append(f"[{fid}] 전부 비어 있는(all-NaN) 피처: {shown}. 학습에 무의미할 수 있습니다.")
            if info.get("feat_const"):
                shown = ", ".join(info["feat_const"][:5]) + (f" 외 {len(info['feat_const'])-5}개" if len(info["feat_const"]) > 5 else "")
                warns.append(f"[{fid}] 값이 하나뿐인(상수) 피처: {shown}. 분산이 없어 학습에 무의미할 수 있습니다.")

            # (15) 스플릿 대비 데이터 크기(단일 파일 + 비율 분할일 때)
            if layout.get("single_file") and splits:
                mn = min(dist.values()) if dist else info.get("n_rows", 0)
                if info.get("n_rows", 0) < 10:
                    warns.append(f"[{fid}] 데이터가 {info.get('n_rows')}행뿐이라 비율 분할이 불안정할 수 있습니다.")
                elif dist and mn < 2:
                    warns.append(f"[{fid}] 가장 적은 클래스 표본이 {mn}개라 층화 분할이 어려울 수 있습니다.")

        # (16) 매핑 정합성 — 1층 정적 매핑 검증
        if layout["strategy"] == "tabular_embedded":
            if layout.get("mapping") == "column":
                infos.append("[매핑] column 방식: 라벨과 피처가 같은 행에서 추출되어 행 정렬이 보장됩니다.")
        elif layout["strategy"] == "tabular_join":
            j_err, j_warn, j_info = validate_join_mapping(layout, label, classes, skip_verify)
            errors.extend(j_err)
            warns.extend(j_warn)
            infos.extend(j_info)
        else:
            m_err, m_warn, m_info = validate_sidecar_mapping(layout, base_dir)
            errors.extend(m_err)
            warns.extend(m_warn)
            infos.extend(m_info)

    return errors, warns, infos


# ---------------------------------------------------------------------------
# 6) 오케스트레이션
# ---------------------------------------------------------------------------
def convert(input_path, output_path=None, base_override=None, skip_verify=False):
    """엑셀 -> 검증 -> (통과 시) JSON-LD 작성.

    데이터 경로는 엑셀의 '데이터 경로 + 스플릿'으로 만들어 실제 파일 존재를 검사한다.
    기준 폴더는 자동으로 정해지며(-d 없이 동작), base_override 로 상위 폴더를 지정할 수 있다.
    output_path 가 None 이면 데이터셋 ID 로 파일명을 만든다(예: MAX_TR_DS01.jsonld).
    반환값: 종료 코드(0=성공, 1=검증 실패로 파일 미생성, 2=입력 파일 없음).
    """
    if not os.path.isfile(input_path):
        print(f"❌ 입력 파일을 찾을 수 없습니다: {input_path}")
        return 2

    sheets = load_sheets(input_path)
    splits = build_splits(sheets["split"])
    classes = parse_classes(sheets["label"])
    # 엑셀의 데이터 경로를 기준으로 실제 파일이 있는 폴더를 자동으로 찾는다.
    data_root = clean_path(sheets["data"].get("데이터 경로"))
    base_dir, tried = resolve_base_dir(input_path, data_root, base_override)
    layout = plan_layout(sheets["base"], sheets["data"], sheets["label"], splits, base_dir)
    distribution = build_distribution(layout)
    record_sets = build_record_sets(classes, splits, layout)
    metadata = build_metadata(sheets, distribution, record_sets)

    errors, warns, infos = validate(sheets, splits, classes, distribution, layout, base_dir, tried, skip_verify)

    if not skip_verify and data_root and os.path.isdir(os.path.join(base_dir, data_root)):
        print(f"🔎 데이터 위치 확인 기준: {os.path.join(base_dir, data_root)}\n")

    # 오류가 있으면 먼저 반려 사유부터 보여준다(파일 미생성).
    if errors:
        out_label = output_path or f"{sheets['base'].get('데이터셋 ID')}.jsonld"
        print(f"❌ 오류 — {out_label} 를 생성하지 않았습니다. 수정 후 다시 실행해 주세요:\n")
        for e in errors:
            print("   -", e)
        print(f"\n총 {len(errors)}개 항목을 수정해야 합니다.")
        return 1

    # 파일은 생성된다. 확인 권장(경고) → 참고(정보) 순으로 보여준다.
    if warns:
        print("⚠️  경고 (파일은 생성되나 확인 권장):")
        for w in warns:
            print("   -", w)
        print()
    if infos:
        print("ℹ️  정보 (참고):")
        for i in infos:
            print("   -", i)
        print()

    # -o 를 생략했으면 데이터셋 ID 로 파일명을 만든다(검증을 통과했으므로 ID 는 유효함).
    if not output_path:
        output_path = f"{metadata['name']}.jsonld"

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)
    print("✅ 검증 통과 — 생성 완료:", output_path)
    return 0


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="엑셀 데이터셋 명세를 Croissant JSON-LD 로 변환합니다.")
    parser.add_argument("-i", "--input", default=DEFAULT_INPUT,
                        help=f"입력 엑셀 파일 (기본값: {DEFAULT_INPUT})")
    parser.add_argument("-o", "--output", default=None,
                        help="출력 JSON-LD 파일 (기본값: 데이터셋 ID, 예: MAX_TR_DS01.jsonld)")
    parser.add_argument("-d", "--base-dir", default=None,
                        help="데이터가 들어 있는 상위 폴더(선택). 미지정 시 현재 폴더 또는 엑셀 파일이 "
                             "있는 폴더를 기준으로 엑셀의 '데이터 경로'를 찾아 자동으로 검증합니다.")
    parser.add_argument("--no-verify", action="store_true",
                        help="실제 파일 존재 여부 검사를 건너뜁니다(데이터 없이 메타데이터만 생성).")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    return convert(args.input, args.output, args.base_dir, args.no_verify)


if __name__ == "__main__":
    sys.exit(main())
