import numpy as np

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


def _classification_predictions(model, X):
    """Return hard labels and optional positive-class probabilities."""
    y_pred = np.asarray(model.predict(X))
    y_prob = None

    if y_pred.ndim == 2 and y_pred.shape[1] >= 2:
        y_prob = y_pred[:, 1]
        y_pred = np.argmax(y_pred, axis=1)
    else:
        y_pred = y_pred.reshape(-1)
        # Binary classifiers commonly expose probability for the positive class.
        if hasattr(model, "predict_proba"):
            proba = np.asarray(model.predict_proba(X))
            if proba.ndim == 2 and proba.shape[1] >= 2:
                y_prob = proba[:, 1]
                y_pred = (y_prob >= 0.5).astype(int)
        elif np.issubdtype(y_pred.dtype, np.floating) and np.all((y_pred >= 0) & (y_pred <= 1)):
            y_prob = y_pred.copy()
            y_pred = (y_pred >= 0.5).astype(int)

    return y_pred, y_prob


def evaluate_classification_model(model, X, y_true, labels=None):
    """Evaluate a classification model from tabular data.

    Parameters
    ----------
    model : sklearn/xgboost-compatible estimator
        Must expose ``predict``; ``predict_proba`` is used when available.
    X : array-like
        Model-ready feature matrix. Dataset-specific preprocessing belongs outside
        this evaluator.
    y_true : array-like
        Ground-truth labels.
    labels : optional sequence
        Label order for confusion matrix/report. If omitted, inferred from y_true.

    Returns
    -------
    dict
        accuracy, precision, recall, f1 and (when probabilities/classes permit)
        roc_auc, plus confusion_matrix and per-class metrics.
    """
    try:
        from sklearn.metrics import (
            accuracy_score,
            classification_report,
            confusion_matrix,
            f1_score,
            precision_score,
            recall_score,
            roc_auc_score,
        )
    except ImportError as exc:
        raise ImportError("Classification evaluation에는 scikit-learn이 필요합니다.") from exc

    y_true = np.asarray(y_true).reshape(-1)
    y_pred, y_prob = _classification_predictions(model, X)
    if len(y_true) != len(y_pred):
        raise ValueError(f"y_true({len(y_true)})와 y_pred({len(y_pred)})의 길이가 다릅니다.")

    unique_labels = np.unique(y_true) if labels is None else np.asarray(labels)
    average = "binary" if len(unique_labels) == 2 else "weighted"

    result = {
        "accuracy": float(accuracy_score(y_true, y_pred)),
        "precision": float(precision_score(y_true, y_pred, average=average, zero_division=0)),
        "recall": float(recall_score(y_true, y_pred, average=average, zero_division=0)),
        "f1": float(f1_score(y_true, y_pred, average=average, zero_division=0)),
        "confusion_matrix": confusion_matrix(y_true, y_pred, labels=unique_labels),
        "classification_report": classification_report(
            y_true, y_pred, labels=unique_labels, zero_division=0, output_dict=True
        ),
        "labels": unique_labels.tolist(),
    }

    if y_prob is not None and len(unique_labels) == 2:
        try:
            # For labels such as {-1, 1}, the positive class is the larger label.
            positive_label = unique_labels[-1]
            y_binary = (y_true == positive_label).astype(int)
            result["roc_auc"] = float(roc_auc_score(y_binary, y_prob))
        except ValueError:
            # A single-class test split cannot define AUROC.
            result["roc_auc"] = None
    else:
        result["roc_auc"] = None

    return result


def print_classification_result(result, split="test"):
    print(f"=== Classification 평가 결과 ({split}셋) ===")
    print(f"  Accuracy  : {result['accuracy']:.4f}")
    print(f"  Precision : {result['precision']:.4f}")
    print(f"  Recall    : {result['recall']:.4f}")
    print(f"  F1        : {result['f1']:.4f}")
    if result.get("roc_auc") is not None:
        print(f"  ROC-AUC   : {result['roc_auc']:.4f}")
    print("  Confusion Matrix:")
    print(result["confusion_matrix"])


def print_evaluation_result(result, split="test"):
    print(f"=== 평가 결과 ({split}셋) ===")
    print(f"  mAP50      (IoU=0.5)   : {float(result['MaP@[IoU=50]']):.4f}")
    print(f"  mAP50:95   (COCO 기본) : {float(result['MaP']):.4f}")
    print(f"  Recall@100             : {float(result['Recall@[max_detections=100]']):.4f}")
