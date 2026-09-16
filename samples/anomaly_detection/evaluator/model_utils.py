import keras
import keras_cv


def load_detection_model(
    model_path,
    confidence_threshold=0.01,
    iou_threshold=0.5,
):
    model = keras.models.load_model(model_path)
    model.compile(
        box_loss="ciou",
        classification_loss="binary_crossentropy",
        jit_compile=False,
    )
    model.prediction_decoder = keras_cv.layers.NonMaxSuppression(
        bounding_box_format="xyxy",
        from_logits=True,
        confidence_threshold=confidence_threshold,
        iou_threshold=iou_threshold,
    )
    return model
