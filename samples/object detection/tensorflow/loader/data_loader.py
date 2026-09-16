import os
import random

import mlcroissant as mlc
import numpy as np
import tensorflow as tf


def _as_str(value):
    return value.decode() if isinstance(value, bytes) else value


def _stem(name):
    return os.path.splitext(name)[0]


def load_labels_by_stem(ds, split):
    """Load labels-records-<split> as {filename_stem: label_text}."""
    record_set = f"labels-records-{split}"
    available = {record.uuid for record in ds.metadata.record_sets}
    if record_set not in available:
        return {}

    mapping = {}
    for record in ds.records(record_set=record_set):
        filename = _as_str(record[f"{record_set}/filename"])
        content = _as_str(record[f"{record_set}/content"])
        mapping[_stem(filename)] = content
    return mapping


def _maybe_shuffle(gen, shuffle):
    """tf.data 의 .shuffle(buffer) 와 동일한 버퍼 셔플. shuffle=0 이면 통과."""
    if not shuffle:
        yield from gen
        return
    buf = []
    for item in gen:
        buf.append(item)
        if len(buf) >= shuffle:
            yield buf.pop(random.randrange(len(buf)))
    random.shuffle(buf)
    yield from buf


def _numpy_batches(gen, batch_size):
    """제너레이터를 batch_size 씩 묶어 numpy 배치로. yield 원소 개수와 무관하게 동작."""
    batch = []
    for item in gen:
        batch.append(item)
        if len(batch) == batch_size:
            yield tuple(np.stack(x) for x in zip(*batch))
            batch = []
    if batch:  # 남은 자투리 배치
        yield tuple(np.stack(x) for x in zip(*batch))


def build_detection_dataset(jsonld_path, split="train",
                            img_size=(224, 224), max_boxes=10, batch_size=8, shuffle=0,
                            parser=None, backend="numpy"):
    """(image, boxes, classes, num_boxes) 배치를 내놓는 데이터셋.

    backend : "numpy"(기본) | "tensorflow" | "torch"
    parser  : BaseParser 인스턴스 (기본 YOLOParser). boxes/classes 분리 반환.

    image     : (H, W, 3) float32, [0,1]  (세 백엔드 모두 채널-라스트로 통일)
    boxes     : (max_boxes, 4) float32  — [cx, cy, w, h], 뒤쪽 0-패딩
    classes   : (max_boxes,)   float32  — 패딩 자리는 -1
    num_boxes : () int32                — 유효 박스 수
    """
    if parser is None:
        raise ValueError("parser를 지정하세요. 예: YOLOParser(max_boxes=10)")

    ds = mlc.Dataset(jsonld_path)
    record_set = f"data-records-{split}"
    img_key = f"{record_set}/image"
    name_key = f"{record_set}/filename"

    labels_by_stem = load_labels_by_stem(ds, split)  # 파일명 stem -> 라벨 텍스트

    def generator():
        for rec in ds.records(record_set=record_set):
            img = rec[img_key]  # PIL.Image
            arr = np.asarray(img.convert("RGB").resize(img_size), dtype=np.float32) / 255.0
            fname = _as_str(rec[name_key])
            annotation = labels_by_stem.get(_stem(fname), "")     # 파일명으로 라벨 매칭
            boxes, classes, n = parser.parse(annotation)          # 4개로 분리
            yield arr, boxes, classes, np.int32(n)

    backend = backend.lower()

    # ---- numpy (기본) ----
    if backend == "numpy":
        return _numpy_batches(_maybe_shuffle(generator(), shuffle), batch_size)

    # ---- tensorflow ----
    if backend == "tensorflow":
        import tensorflow as tf
        output_signature = (
            tf.TensorSpec(shape=(*img_size, 3), dtype=tf.float32),
            tf.TensorSpec(shape=(max_boxes, 4), dtype=tf.float32),   # boxes
            tf.TensorSpec(shape=(max_boxes,), dtype=tf.float32),     # classes
            tf.TensorSpec(shape=(), dtype=tf.int32),                 # num_boxes
        )
        tf_ds = tf.data.Dataset.from_generator(generator, output_signature=output_signature)
        if shuffle:
            tf_ds = tf_ds.shuffle(shuffle)
        return tf_ds.batch(batch_size).prefetch(tf.data.AUTOTUNE)

    # ---- torch ----
    if backend == "torch":
        import torch

        class _IterableDS(torch.utils.data.IterableDataset):
            def __iter__(self):
                return _maybe_shuffle(generator(), shuffle)

        return torch.utils.data.DataLoader(
            _IterableDS(),
            batch_size=batch_size,
            pin_memory=torch.cuda.is_available(),
        )

    raise ValueError(f"backend must be 'numpy', 'tensorflow', or 'torch' (got {backend!r})")


def build_tabular_dataset(jsonld_path, split="train", numeric_only=True):
    """Croissant Dataset -> (X, y, feature_names).
 
    record_set 이름은 관례상 'records-{split}' 또는 단일 파일이면 'records'.
    """
    _NUMERIC = {"Float", "Integer"}
    
    ds = mlc.Dataset(jsonld_path)                     # 경로가 맞으면 mapping 불필요
 
    ids = {r.uuid for r in ds.metadata.record_sets}
    record_set = f"records-{split}" if f"records-{split}" in ids else "records"
    rs = next(r for r in ds.metadata.record_sets if r.uuid == record_set)
 
    label_key = f"{record_set}/label"
    feature_fields = [f for f in rs.fields if f.uuid != label_key]
    if numeric_only:
        def _is_numeric(field):
            return any(str(t).split("/")[-1] in _NUMERIC for t in field.data_types)

        feature_fields = [f for f in feature_fields if _is_numeric(f)]
    feature_keys = [f.uuid for f in feature_fields]
    feature_names = [k.split("/")[-1] for k in feature_keys]
 
    X, y = [], []
    for record in ds.records(record_set=record_set):
        X.append([float("nan") if record[k] is None else record[k] for k in feature_keys])
        y.append(record.get(label_key))
    return np.asarray(X, dtype=np.float32), np.asarray(y), feature_names