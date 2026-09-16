import numpy as np

from .base_parser import BaseParser


class YOLOParser(BaseParser):
    """YOLO Annotation -> 내부 표준 Detection 형태.

    반환값:
      boxes:   (max_boxes, 4), normalized [cx, cy, w, h]
      classes: (max_boxes,)
      num_boxes: 실제 Bounding Box 개수
    """

    def __init__(self, max_boxes=10):
        self.max_boxes = max_boxes

    def parse(self, annotation):
        boxes = []
        classes = []

        for line in (annotation or "").splitlines():
            parts = line.split()
            if len(parts) != 5:
                continue

            class_id, cx, cy, w, h = [float(x) for x in parts]
            classes.append(class_id)
            boxes.append([cx, cy, w, h])

        boxes = np.asarray(boxes, dtype=np.float32).reshape(-1, 4)[: self.max_boxes]
        classes = np.asarray(classes, dtype=np.float32).reshape(-1)[: self.max_boxes]

        padded_boxes = np.zeros((self.max_boxes, 4), dtype=np.float32)
        padded_classes = np.full((self.max_boxes,), -1.0, dtype=np.float32)

        count = len(boxes)
        padded_boxes[:count] = boxes
        padded_classes[:count] = classes

        return padded_boxes, padded_classes, count
