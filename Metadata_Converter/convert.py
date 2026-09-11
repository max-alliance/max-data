#!/usr/bin/env python3
"""template.xlsx -> Croissant metadata.jsonld 변환기.

사용법:
    python convert.py                          # 기본값: -i template.xlsx -o metadata.jsonld
    python convert.py -i template.xlsx -o metadata.jsonld
    python convert.py -i template.xlsx -o out.jsonld -d ./MAX_TR_DS01_root

변환 전에 필수값/형식/데이터 경로를 검사한다. 문제(ERROR)가 하나라도 있으면
출력 파일을 만들지 않고 "이렇게 고치세요" 안내만 출력하고 종료 코드 1을 반환한다.
"""
import argparse
import datetime
import glob
import json
import os
import re
import sys

import openpyxl

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

EXT2MIME = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
            "txt": "text/plain", "csv": "text/csv"}

CANONICAL_CONTEXT = {
    "@language": "en", "@vocab": "https://schema.org/", "citeAs": "cr:citeAs",
    "column": "cr:column", "conformsTo": "dct:conformsTo",
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


def mime_from_pattern(pattern):
    ext = str(pattern).split(".")[-1].lower().strip()
    return EXT2MIME.get(ext, "application/octet-stream")


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
# 2) 파싱: 스플릿 / 클래스
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
    """02 시트의 '클래스 분류'("0: 정상 / 1: 비정상")를 [(0,'정상'), ...] 로 파싱."""
    cls_raw = label.get("클래스 분류")
    classes = []
    if cls_raw:
        for line in str(cls_raw).splitlines():
            line = line.strip()
            if ":" in line:
                k, v = line.split(":", 1)
                if k.strip().isdigit():
                    classes.append((int(k.strip()), v.strip()))
    return classes


# ---------------------------------------------------------------------------
# 3) Croissant 구성요소 빌드
# ---------------------------------------------------------------------------
def build_distribution(data, label, splits):
    """스플릿별 이미지/라벨 FileSet(해시 불필요) 생성."""
    data_root = clean_path(data.get("데이터 경로"))
    label_root = clean_path(label.get("데이터 경로"))
    img_pat = str(data.get("파일 패턴", "*.jpg")).strip()
    lbl_pat = str(label.get("파일 패턴", "*.txt")).strip()
    img_mime = mime_from_pattern(img_pat)
    lbl_mime = mime_from_pattern(lbl_pat)
    lbl_fmt = label.get("라벨 파일 형식")

    distribution = []
    for s in splits:
        distribution.append({
            "@type": "cr:FileSet", "@id": f"images-{s['name']}",
            "name": f"images-{s['name']}",
            "description": f"{s['name']} 이미지",
            "includes": f"{data_root}/{s['sub']}/{img_pat}",
            "encodingFormat": img_mime,
        })
        distribution.append({
            "@type": "cr:FileSet", "@id": f"labels-{s['name']}",
            "name": f"labels-{s['name']}",
            "description": f"{s['name']} 라벨 ({lbl_fmt})",
            "includes": f"{label_root}/{s['sub']}/{lbl_pat}",
            "encodingFormat": lbl_mime,
        })
    return distribution


def build_record_sets(classes, splits, data_type=None, distribution=None):
    """클래스/스플릿 enumeration + (이미지 데이터일 때) 스플릿별 예시 RecordSet 생성.

    data-records-<split> RecordSet 은 이미지 FileSet 하나만 읽어(join 불필요) 안정적으로
    image(PIL) + filename 을 emit 한다. 이게 있어야 tf.data 로 로드할 수 있다.
    라벨(YOLO bbox 등)은 이미지↔라벨 FileSet join 이 필요한데 현재 mlcroissant 에서
    까다로워, 여기서는 넣지 않고 학습 코드( TF 제너레이터 )에서 파일명으로 매핑하는 것을 권장한다.
    """
    record_sets = []
    if classes:
        record_sets.append({
            "@type": "cr:RecordSet", "@id": "classes", "name": "classes",
            "description": "클래스 정의", "dataType": "sc:Enumeration",
            "key": {"@id": "classes/id"},
            "field": [
                {"@type": "cr:Field", "@id": "classes/id", "name": "id",
                 "description": "클래스 인덱스", "dataType": "sc:Integer"},
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

    # 이미지 데이터셋이면 스플릿별로 학습용 예시 RecordSet 을 만든다.
    if str(data_type).strip().lower() == "image":
        label_ids = {d["@id"] for d in (distribution or [])}
        for s in splits:
            rid = f"data-records-{s['name']}"
            img_fs = f"images-{s['name']}"
            record_sets.append({
                "@type": "cr:RecordSet", "@id": rid, "name": rid,
                "description": f"{s['name']} 예시(이미지)",
                "field": [
                    {"@type": "cr:Field", "@id": f"{rid}/image", "name": "image",
                     "description": "이미지", "dataType": "sc:ImageObject",
                     "source": {"fileSet": {"@id": img_fs}, "extract": {"fileProperty": "content"}}},
                    {"@type": "cr:Field", "@id": f"{rid}/filename", "name": "filename",
                     "description": "이미지 파일명(라벨 매핑 키)", "dataType": "sc:Text",
                     "source": {"fileSet": {"@id": img_fs}, "extract": {"fileProperty": "filename"}}},
                ],
            })
            # 라벨 FileSet 이 있으면, 라벨 파일명+내용을 그대로 내주는 RecordSet 도 만든다.
            # (로더는 파일 경로를 조립하지 않고 여기서 filename/content 를 바로 받는다.)
            lbl_fs = f"labels-{s['name']}"
            if lbl_fs in label_ids:
                lrid = f"labels-records-{s['name']}"
                record_sets.append({
                    "@type": "cr:RecordSet", "@id": lrid, "name": lrid,
                    "description": f"{s['name']} 라벨 파일(파일명+내용)",
                    "field": [
                        {"@type": "cr:Field", "@id": f"{lrid}/filename", "name": "filename",
                         "description": "라벨 파일명(이미지 매핑 키)", "dataType": "sc:Text",
                         "source": {"fileSet": {"@id": lbl_fs}, "extract": {"fileProperty": "filename"}}},
                        {"@type": "cr:Field", "@id": f"{lrid}/content", "name": "content",
                         "description": "라벨 파일 내용(원본 텍스트)", "dataType": "sc:Text",
                         "source": {"fileSet": {"@id": lbl_fs}, "extract": {"fileProperty": "content"}}},
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
        "distribution": distribution,
        "recordSet": record_sets,
    }
    return {k: v for k, v in metadata.items() if v not in (None, "", [], {})}


# ---------------------------------------------------------------------------
# 4) 검증
# ---------------------------------------------------------------------------
def validate(sheets, splits, classes, distribution, data_base_dir):
    """(errors, warns) 반환. errors 가 있으면 파일을 만들면 안 된다."""
    errors, warns = [], []
    base, data_i, label = sheets["base"], sheets["data"], sheets["label"]

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

    # (3) 파일 패턴 형식(*.확장자)
    for sheet_name, vals in [("01_데이터 정보", data_i), ("02_라벨 정보", label)]:
        pat = vals.get("파일 패턴")
        if not is_empty(pat) and not re.fullmatch(r"\*\.[A-Za-z0-9]+", str(pat).strip()):
            errors.append(f"[{sheet_name}] 파일 패턴 '{pat}' 형식이 올바르지 않습니다. "
                          "→ '*.jpg' 처럼 '*.확장자' 형태로 입력하세요.")

    # (4) 알 수 없는 파일 형식 경고
    img_pat = str(data_i.get("파일 패턴", "")).strip()
    lbl_pat = str(label.get("파일 패턴", "")).strip()
    if not is_empty(img_pat) and mime_from_pattern(img_pat) == "application/octet-stream":
        warns.append(f"[01_데이터 정보] 파일 패턴 '{img_pat}' 의 확장자를 인식하지 못했습니다. "
                     "→ jpg/png/csv/txt 등 알려진 확장자인지 확인하세요.")
    if not is_empty(lbl_pat) and mime_from_pattern(lbl_pat) == "application/octet-stream":
        warns.append(f"[02_라벨 정보] 파일 패턴 '{lbl_pat}' 의 확장자를 인식하지 못했습니다.")

    # (5) 스플릿 비율(0~1 숫자, 합≈1)
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

    # (6) 생성일 형식
    raw_date = base.get("생성일")
    if not is_empty(raw_date) and not isinstance(raw_date, datetime.datetime):
        try:
            datetime.date.fromisoformat(str(raw_date).strip()[:10])
        except ValueError:
            warns.append(f"[기본정보] 생성일 '{raw_date}' 이(가) 날짜 형식이 아닙니다. → YYYY-MM-DD 형태로 입력하세요.")

    # (7) 클래스 수 vs 실제 클래스 개수
    cls_count = label.get("클래스 수")
    if not is_empty(cls_count) and classes:
        try:
            if int(cls_count) != len(classes):
                warns.append(f"[02_라벨 정보] 클래스 수({cls_count})와 실제 클래스 분류 개수({len(classes)})가 다릅니다. "
                             "→ 둘을 일치시켜 주세요.")
        except (TypeError, ValueError):
            warns.append(f"[02_라벨 정보] 클래스 수 '{cls_count}' 이(가) 정수가 아닙니다.")

    # (8) 데이터 경로(파일) 실제 존재 여부
    if data_base_dir is None:
        warns.append("데이터 경로 존재 여부는 검사하지 않았습니다. "
                     "→ 실제 데이터 폴더가 준비되면 -d/--data-dir 로 그 폴더를 지정하세요.")
    elif not os.path.isdir(data_base_dir):
        errors.append(f"데이터 루트 폴더 '{data_base_dir}' 가 존재하지 않습니다. "
                      "→ -d/--data-dir 를 실제 데이터셋 폴더 경로로 지정하세요.")
    else:
        for d in distribution:
            if not glob.glob(os.path.join(data_base_dir, d["includes"])):
                errors.append(f"[{d['@id']}] 경로 패턴 '{d['includes']}' 에 해당하는 파일이 없습니다. "
                              f"→ '{data_base_dir}' 아래에 해당 파일이 있는지, 경로/패턴이 맞는지 확인하세요.")

    return errors, warns


# ---------------------------------------------------------------------------
# 5) 오케스트레이션
# ---------------------------------------------------------------------------
def convert(input_path, output_path=None, data_base_dir=None):
    """엑셀 -> 검증 -> (통과 시) JSON-LD 작성.

    output_path 가 None 이면 데이터셋 ID 로 파일명을 만든다(예: MAX_TR_DS01.jsonld).
    반환값: 종료 코드(0=성공, 1=검증 실패로 파일 미생성, 2=입력 파일 없음).
    """
    if not os.path.isfile(input_path):
        print(f"❌ 입력 파일을 찾을 수 없습니다: {input_path}")
        return 2

    sheets = load_sheets(input_path)
    splits = build_splits(sheets["split"])
    classes = parse_classes(sheets["label"])
    distribution = build_distribution(sheets["data"], sheets["label"], splits)
    record_sets = build_record_sets(classes, splits, sheets["base"].get("데이터 유형"), distribution)
    metadata = build_metadata(sheets, distribution, record_sets)

    errors, warns = validate(sheets, splits, classes, distribution, data_base_dir)

    if warns:
        print("⚠️  경고 (파일은 생성됩니다):")
        for w in warns:
            print("   -", w)
        print()

    if errors:
        out_label = output_path or f"{sheets['base'].get('데이터셋 ID')}.jsonld"
        print(f"❌ 다음 문제 때문에 {out_label} 를 생성하지 않았습니다. 수정 후 다시 실행해 주세요:\n")
        for e in errors:
            print("   -", e)
        print(f"\n총 {len(errors)}개 항목을 수정해야 합니다.")
        return 1

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
    parser.add_argument("-d", "--data-dir", default=None,
                        help="데이터 실제 루트 폴더(지정 시 파일 존재 여부까지 검사, 기본값: 검사 안 함)")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    return convert(args.input, args.output, args.data_dir)


if __name__ == "__main__":
    sys.exit(main())
