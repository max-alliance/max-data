from keras_cv.metrics import BoxCOCOMetrics


def evaluate_detection_model(model, dataset):
    metric = BoxCOCOMetrics(
        bounding_box_format="xyxy",
        evaluate_freq=1e9,
    )
    metric.reset_state()

    for images, y_true in dataset:
        y_pred = model.predict(images, verbose=0)
        metric.update_state(y_true, y_pred)

    return metric.result(force=True)


def print_evaluation_result(result, split="test"):
    print(f"=== 평가 결과 ({split}셋) ===")
    print(f"  mAP50      (IoU=0.5)   : {float(result['MaP@[IoU=50]']):.4f}")
    print(f"  mAP50:95   (COCO 기본) : {float(result['MaP']):.4f}")
    print(f"  Recall@100             : {float(result['Recall@[max_detections=100]']):.4f}")
