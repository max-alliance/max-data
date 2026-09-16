import os

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


def build_detection_dataset(
    jsonld_path,
    split="train",
    img_size=(224, 224),
    max_boxes=10,
    batch_size=8,
    shuffle=0,
    parser=None,
):
    """Croissant Dataset -> detection TensorFlow Dataset.

    Annotation 해석은 parser에 위임합니다.
    """
    if parser is None:
        raise ValueError("parser를 지정하세요. 예: YOLOParser(max_boxes=10)")

    ds = mlc.Dataset(jsonld_path)
    record_set = f"data-records-{split}"
    img_key = f"{record_set}/image"
    name_key = f"{record_set}/filename"
    labels_by_stem = load_labels_by_stem(ds, split)

    def generator():
        for record in ds.records(record_set=record_set):
            image = record[img_key]
            array = np.asarray(
                image.convert("RGB").resize(img_size),
                dtype=np.float32,
            ) / 255.0

            filename = _as_str(record[name_key])
            annotation = labels_by_stem.get(_stem(filename), "")
            boxes, classes, num_boxes = parser.parse(annotation)

            yield (
                array,
                {
                    "boxes": np.asarray(boxes, dtype=np.float32),
                    "classes": np.asarray(classes, dtype=np.float32),
                    "num_boxes": np.int32(num_boxes),
                },
            )

    output_signature = (
        tf.TensorSpec(shape=(*img_size, 3), dtype=tf.float32),
        {
            "boxes": tf.TensorSpec(shape=(max_boxes, 4), dtype=tf.float32),
            "classes": tf.TensorSpec(shape=(max_boxes,), dtype=tf.float32),
            "num_boxes": tf.TensorSpec(shape=(), dtype=tf.int32),
        },
    )

    tf_dataset = tf.data.Dataset.from_generator(generator, output_signature=output_signature)
    if shuffle:
        tf_dataset = tf_dataset.shuffle(shuffle)
    return tf_dataset.batch(batch_size).prefetch(tf.data.AUTOTUNE)


def _get_record_set(ds, split):
    ids = {record.uuid for record in ds.metadata.record_sets}
    preferred = f"records-{split}"
    if preferred in ids:
        return preferred
    if "records" in ids:
        return "records"
    raise ValueError(
        f"Croissant metadata에서 '{preferred}' 또는 'records' RecordSet을 찾을 수 없습니다. "
        f"사용 가능한 RecordSet: {sorted(ids)}"
    )


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