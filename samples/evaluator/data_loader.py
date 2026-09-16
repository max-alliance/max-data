import os

import mlcroissant as mlc
import numpy as np
import tensorflow as tf

from .parser.tabular_parser import CroissantTabularParser


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


def build_tabular_dataset(
    jsonld_path,
    split="train",
    numeric_only=True,
    exclude_fields=None,
    parser=None,
):
    """Croissant Dataset -> (X, y, feature_names).

    Parser는 RecordSet/Field 구조 해석만 담당합니다. 기관/모델별 전처리는
    이 함수 밖에서 수행해야 합니다.
    """
    ds = mlc.Dataset(jsonld_path)
    record_set = _get_record_set(ds, split)
    rs = next(r for r in ds.metadata.record_sets if r.uuid == record_set)

    label_key = f"{record_set}/label"
    if not any(f.uuid == label_key for f in rs.fields):
        raise ValueError(f"RecordSet '{record_set}'에 label field가 없습니다: {label_key}")

    if parser is None:
        parser = CroissantTabularParser(
            numeric_only=numeric_only,
            exclude_fields=exclude_fields,
        )

    return parser.parse(ds, record_set, label_key, rs.fields)


def load_csv_classification_dataset(
    csv_path,
    label_column="Pass/Fail",
    feature_columns=None,
    exclude_columns=("Time",),
    dtype=np.float32,
):
    """Load a raw CSV classification dataset as (X, y, feature_names).

    This is intentionally a thin I/O adapter. It does not perform imputation,
    scaling, feature selection, or feature engineering; those remain external
    to the common module and should be implemented in the institution Notebook
    when required by the model.
    """
    import pandas as pd

    df = pd.read_csv(csv_path)
    if label_column not in df.columns:
        raise ValueError(f"label_column '{label_column}'이 CSV 컬럼에 없습니다.")

    if feature_columns is None:
        excluded = set(exclude_columns or ()) | {label_column}
        feature_columns = [column for column in df.columns if column not in excluded]

    missing = [column for column in feature_columns if column not in df.columns]
    if missing:
        raise ValueError(f"CSV에 없는 feature column: {missing[:10]}")

    X = df.loc[:, feature_columns].to_numpy(dtype=dtype)
    y = df.loc[:, label_column].to_numpy()
    return X, y, list(feature_columns)
