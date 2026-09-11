# Manufacturing AX Dataset Tools

> **제조 AI 데이터의 표준화된 메타데이터 생성부터 객체탐지 모델 평가까지 지원하는 데이터 활용 모듈**

제조기업이 작성한 데이터 정보를 **MLCommons Croissant 1.1 표준**으로 변환하고, 생성된 Croissant Dataset을 기반으로 **AI 모델을 평가**할 수 있도록 구성한 Python 모듈 모음입니다.

---

## 📁 Repository Structure

```text
Manufacturing-AI-Dataset-Tools/
│
├── Dataset_Evaluation_Modules/
│   └── Croissant Dataset → YOLO/KerasCV
│       → Detection Model Evaluation
│
├── Metadata_Convertor/
│   └── Excel Dataset Template → Croissant 1.1 JSON-LD
│
└── samples/
    └── 데이터 변환 및 평가 예시
```

### Modules

| Folder | Description |
|---|---|
| **Metadata_Convertor** | 제조기업 Excel 데이터 제출 양식을 Croissant 1.1 JSON-LD 메타데이터로 변환 |
| **Dataset_Evaluation_Modules** | Croissant Dataset을 TensorFlow/KerasCV 형식으로 로드하고 객체탐지 모델 평가 |
| **samples** | 각 모듈의 실제 사용 예시 및 샘플 데이터 |

---

## 🔄 Overall Workflow

```text
제조기업 데이터
      │
      ▼
┌────────────────────────┐
│ Metadata_Convertor      │
│ Excel → Croissant JSON-LD│
└───────────┬────────────┘
            │
            ▼
      metadata.jsonld
            │
            ▼
┌────────────────────────────┐
│ Dataset_Evaluation_Modules │
│ Croissant → Dataset        │
│ YOLO → KerasCV             │
│ Model → Evaluation         │
└────────────┬───────────────┘
             │
             ▼
        mAP / Recall
```

---

## 🚀 Quick Start

### 1. Metadata Conversion

제조기업에서 작성한 Excel 제출 양식을 Croissant 1.1 JSON-LD로 변환합니다.

```bash
cd Metadata_Convertor

python convert.py \
    -i template.xlsx \
    -o metadata.jsonld
```

생성된 `metadata.jsonld`는 `mlcroissant`를 통해 로드 및 검증할 수 있습니다.

자세한 내용은 [`Metadata_Convertor/README.md`](./Metadata_Convertor/README.md)를 참고하세요.

---

### 2. Dataset Evaluation

생성된 Croissant Dataset을 이용하여 객체탐지 모델을 평가합니다.

```python
from src.data_loader import build_detection_dataset
from src.format_converter import to_kcv_format
from src.model_utils import load_detection_model
from src.evaluator import evaluate_detection_model

dataset = build_detection_dataset(
    "metadata.jsonld",
    split="test",
    img_size=(640, 640),
    max_boxes=10,
    batch_size=32,
)

dataset = dataset.map(
    to_kcv_format(640, 640)
)

model = load_detection_model(
    "yolov8_full.keras"
)

result = evaluate_detection_model(
    model,
    dataset,
)
```

자세한 모듈별 사용법은 [`Dataset_Evaluation_Modules/README.md`](./Dataset_Evaluation_Modules/README.md)를 참고하세요.

---

## 🧪 Samples

`samples/` 폴더에는 각 기능을 이해하고 바로 실행해볼 수 있는 **예시 코드 및 데이터**를 제공합니다.

```text
samples/
├── metadata/
│   └── Excel → Croissant 변환 예시
│
└── evaluation/
    └── Dataset 평가 예시
```

> 실제 samples 구성은 추가되는 예제에 따라 확장할 수 있습니다.

---

## 📌 Purpose

본 Repository는 제조 AI 데이터의 **표준화 → 활용 → 검증**을 하나의 흐름으로 연결하는 것을 목적으로 합니다.

- **Standardization** — 제조 데이터 메타데이터를 Croissant 표준으로 통일
- **Reusability** — 데이터 로딩 및 평가 기능을 모듈화하여 재사용
- **Validation** — 데이터 및 모델 평가 과정의 일관성 확보
- **Examples** — `samples/`를 통한 실제 활용 방법 제공

---

## 📚 References

- [MLCommons Croissant](https://mlcommons.org/working-groups/data/croissant/)
- [Croissant 1.1 Specification](https://mlcommons.org/croissant/1.1)

---

## 📄 License

M.AX 얼라이언스 데이터셋 라이선스 정책을 따릅니다.