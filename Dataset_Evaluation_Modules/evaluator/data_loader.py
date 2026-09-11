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
    """Croissant Dataset -> (image, boxes, classes, num_boxes) Dataset.

    Annotation 해석은 parser에 위임합니다. 따라서 기관은 Notebook에서
    자체 Parser를 구현해 주입할 수 있습니다.
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
                np.asarray(boxes, dtype=np.float32),
                np.asarray(classes, dtype=np.float32),
                np.int32(num_boxes),
            )

    output_signature = (
        tf.TensorSpec(shape=(*img_size, 3), dtype=tf.float32),
        tf.TensorSpec(shape=(max_boxes, 4), dtype=tf.float32),
        tf.TensorSpec(shape=(max_boxes,), dtype=tf.float32),
        tf.TensorSpec(shape=(), dtype=tf.int32),
    )

    tf_dataset = tf.data.Dataset.from_generator(
        generator,
        output_signature=output_signature,
    )

    if shuffle:
        tf_dataset = tf_dataset.shuffle(shuffle)

    return tf_dataset.batch(batch_size).prefetch(tf.data.AUTOTUNE)
