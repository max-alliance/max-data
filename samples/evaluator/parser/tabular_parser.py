"""Parsers for Croissant/tabular classification datasets."""

from abc import ABC, abstractmethod

import numpy as np


class BaseTabularParser(ABC):
    """Tabular Dataset -> (X, y, feature_names) interface.

    Parser는 Dataset의 구조를 해석하고 내부 표준 형태로 변환하는 역할만 합니다.
    결측치 대체, 스케일링, feature selection/engineering 등 모델별 전처리는
    수행하지 않습니다.
    """

    @abstractmethod
    def parse(self, dataset, record_set, label_key, feature_fields):
        raise NotImplementedError


class CroissantTabularParser(BaseTabularParser):
    """Croissant RecordSet을 (X, y, feature_names)로 변환합니다."""

    def __init__(self, numeric_only=True, exclude_fields=None):
        self.numeric_only = numeric_only
        self.exclude_fields = set(exclude_fields or [])

    @staticmethod
    def _is_numeric(field):
        data_type = str(getattr(field, "data_type", "")).lower()
        return any(token in data_type for token in ("integer", "float", "double", "number"))

    @staticmethod
    def _nan(value):
        if value is None:
            return np.nan
        try:
            if isinstance(value, (float, np.floating)) and np.isnan(value):
                return np.nan
        except TypeError:
            pass
        return value

    def parse(self, dataset, record_set, label_key, feature_fields):
        fields = [
            field for field in feature_fields
            if field.uuid != label_key and field.uuid not in self.exclude_fields
        ]
        if self.numeric_only:
            fields = [field for field in fields if self._is_numeric(field)]

        feature_keys = [field.uuid for field in fields]
        feature_names = [key.split("/")[-1] for key in feature_keys]

        X, y = [], []
        for record in dataset.records(record_set=record_set):
            X.append([self._nan(record.get(key)) for key in feature_keys])
            y.append(record.get(label_key))

        return np.asarray(X, dtype=np.float32), np.asarray(y), feature_names
