# Dataset Loader

Croissant 기반 Dataset Loader 모듈입니다.  
현재 Object Detection과 Tabular Dataset 로딩을 지원합니다.

## Object Detection 구조

```text
Croissant metadata (JSON-LD)
        ↓
Dataset Loader / Parser
        ↓
Image / Bounding Box / Class
        ↓
NumPy / TensorFlow / PyTorch
        ↓
기관 Notebook
  ├─ Custom preprocessing
  ├─ Model input conversion
  └─ Model
```

### Croissant

```python
from dataset_loader.data_loader import build_detection_dataset
from dataset_loader.parser import YOLOParser

dataset = build_detection_dataset(
    "dataset.jsonld",
    split="test",
    parser=YOLOParser(max_boxes=10),
    backend="torch",
    batch_size=8,
)

for images, boxes, classes, num_boxes in dataset:
    # 기관/모델별 전처리 및 평가
    pass
```

### Detection Output

```text
image
    (H, W, 3) / float32

boxes
    (max_boxes, 4)
    [cx, cy, w, h] / normalized

classes
    (max_boxes,)
    padding = -1

num_boxes
    실제 Bounding Box 개수
```

`build_detection_dataset()`은 `numpy`, `tensorflow`, `torch` backend를 지원합니다.

## Tabular Dataset

```text
Croissant metadata (JSON-LD)
        ↓
Dataset Loader
        ↓
(X, y, feature_names)
        ↓
기관 Notebook
  ├─ Custom preprocessing
  ├─ Model input conversion
  └─ Model
```

### Croissant

```python
from dataset_loader.data_loader import build_tabular_dataset

X, y, feature_names = build_tabular_dataset(
    "dataset.jsonld",
    split="test",
    numeric_only=True,
)
```

공통 Loader는 데이터 로딩과 feature/label 분리까지만 담당합니다.  
결측치 처리, scaling, feature selection/engineering 등 모델별 전처리는 Notebook에서 수행합니다.

## Package Structure

```text
dataset_loader/
├─ __init__.py
├─ data_loader.py
└─ parser/
   ├─ __init__.py
   ├─ base_parser.py
   └─ yolo_parser.py
```
