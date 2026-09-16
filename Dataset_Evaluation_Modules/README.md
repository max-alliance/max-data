# Dataset Evaluation Modules

Croissant 기반 Dataset Evaluation 모듈입니다. 현재 Object Detection과 Tabular Classification 평가를 지원합니다.

## Classification / CSV 구조

```text
Croissant metadata 또는 CSV
        ↓
Common Loader / Parser
        ↓
(X, y, feature_names)
        ↓
기관 Notebook
  ├─ Custom preprocessing
  ├─ Model input conversion
  └─ Model
        ↓
evaluate_classification_model
        ↓
Accuracy / Precision / Recall / F1 / ROC-AUC
```

### Croissant

```python
from evaluator.data_loader import build_tabular_dataset
from evaluator.evaluator import evaluate_classification_model, print_classification_result

X, y, feature_names = build_tabular_dataset(
    "MAX_SECOM_DS01.jsonld",
    split="test",
    numeric_only=True,
    exclude_fields={"records-test/Time"},
)

# 기관/모델별 전처리가 필요하면 이 지점에서 직접 수행
# X_model = custom_preprocess(X)
X_model = X

result = evaluate_classification_model(model, X_model, y)
print_classification_result(result)
```

### Raw CSV

```python
from evaluator.data_loader import load_csv_classification_dataset

X, y, feature_names = load_csv_classification_dataset(
    "uci-secom.csv",
    label_column="Pass/Fail",
    exclude_columns=("Time",),
)
```

공통 CSV loader는 CSV 읽기와 feature/label 분리까지만 담당합니다. 결측치 처리, scaling, feature selection/engineering 등 모델별 전처리는 Notebook에서 수행합니다.
