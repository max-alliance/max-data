# Dataset Evaluation Modules

Croissant 기반 데이터셋을 읽고, **기관별 Annotation Parser와 모델 입력 변환 로직을 Notebook에서 연결하여** AI 모델을 평가하기 위한 공통 모듈입니다.

> 원칙: `evaluator/`는 공통 기능을 제공하고, 기관별 Parser·Converter·Model 연결 코드는 각 기관의 Jupyter Notebook에서 작성합니다.

## 📁 Module Structure

```text
Dataset_Evaluation_Modules/
├── evaluator/
│   ├── data_loader.py
│   ├── parser/
│   │   ├── base_parser.py
│   │   └── yolo_parser.py
│   ├── model_utils.py
│   └── evaluator.py
│
├── samples/
└── test.ipynb
```

| Module | Description |
|---|---|
| `data_loader.py` | Croissant Dataset을 읽고 Parser에 Annotation 처리를 위임 |
| `parser/` | Annotation 형식을 내부 표준 Detection 형태로 변환 |
| `base_parser.py` | Parser 기본 인터페이스 |
| `yolo_parser.py` | YOLO Annotation Parser 예시 |
| `model_utils.py` | Detection Model 로드 및 NMS 설정 |
| `evaluator.py` | COCO 기반 Detection Metric 계산 |

## 🔄 Processing Flow

```text
Croissant Dataset
      ↓
data_loader.py
      ↓
Parser
      ↓
Internal Standard Detection Format
      ↓
기관별 Notebook에서 Converter 작성
      ↓
기관/모델별 Input Format
      ↓
Detection Model
      ↓
evaluator.py
      ↓
mAP / Recall
```

## 🚀 기관별 사용 방식

기관은 `src/`에 코드를 추가하지 않고 `test.ipynb`를 복사하여 자신의 환경에 맞게 수정합니다.

### 1. 공통 Parser 사용

```python
from evaluator.parser import YOLOParser

parser = YOLOParser(max_boxes=10)
```

### 2. 기관 자체 Annotation이면 Notebook에서 Parser 작성

```python
class MyParser:
    def parse(self, annotation):
        # 기관 Annotation → 내부 표준 형태
        ...
        return boxes, classes, num_boxes
```

### 3. Dataset Load

```python
from evaluator.data_loader import build_detection_dataset

dataset = build_detection_dataset(
    jsonld_path="metadata.jsonld",
    split="test",
    img_size=(640, 640),
    max_boxes=10,
    batch_size=32,
    parser=parser,
)
```

### 4. 기관별 Converter 작성

```python
import tensorflow as tf

def convert_for_my_model(dataset):
    def _convert(images, boxes, classes, num_boxes):
        # 기관/모델의 입력 규격에 맞게 변환
        ...
        return images, model_inputs

    return dataset.map(_convert)

test_ds = convert_for_my_model(dataset)
```

KerasCV를 사용하는 경우에도 변환 로직은 Notebook에서 직접 작성할 수 있습니다.

```text
Internal Standard Format
          ↓
  Notebook Converter
          ↓
   KerasCV / PyTorch /
   Company Model Input
```

## 📌 역할 분리

- **Croissant**: Dataset metadata 및 구조 표준화
- **`data_loader.py`**: Croissant Dataset 로딩
- **Parser**: Annotation 형식을 내부 표준 형태로 해석
- **기관 Notebook**: 기관별 Parser/Converter/Model 연결
- **`model_utils.py`**: 공통 Model 로딩 기능
- **`evaluator.py`**: 공통 Detection 평가

이 구조에서는 기관별 구현이 공통 저장소의 `src/`에 종속되지 않으며, 동일한 평가 모듈을 각 기관의 데이터와 모델에 맞게 Notebook에서 재사용할 수 있습니다.
