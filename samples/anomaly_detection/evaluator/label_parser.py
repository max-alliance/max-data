"""Backward-compatible import for older notebooks."""

from .parser.yolo_parser import YOLOParser


def parse_yolo(text, max_boxes):
    parser = YOLOParser(max_boxes=max_boxes)
    boxes, classes, num_boxes = parser.parse(text)
    return np.column_stack([classes, boxes]), num_boxes


import numpy as np

__all__ = ["parse_yolo", "YOLOParser"]
