from abc import ABC, abstractmethod


class BaseParser(ABC):
    """Annotation을 내부 표준 Detection 형태로 변환하기 위한 Parser 인터페이스."""

    @abstractmethod
    def parse(self, annotation):
        """Annotation을 (boxes, classes, num_boxes) 형태로 반환합니다."""
        raise NotImplementedError
