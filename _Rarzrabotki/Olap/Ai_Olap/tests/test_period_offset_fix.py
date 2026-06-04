"""TDD tests for period_offset_fix transformer (strip_time + offset).

Скоп: переконатися що transformer
1. Знімає +2000 рік-офсет (4026 → 2026)
2. За замовчуванням обрізає час (strip_time=True): 21:59:59 → 00:00:00
3. Деривує Period_Month як перший день місяця
4. Strip_time можна вимкнути параметром strip_time=False
5. Не падає на None
6. Не падає на вже-фіксованих датах < epoch_threshold
"""
import datetime as dt

import pytest

from ai_olap.transformers import period_offset_fix


def _row(period):
    return {"Period": period}


class TestOffsetStripping:
    def test_datetime_with_2000_offset_year_subtracted(self):
        rows = [_row(dt.datetime(4026, 1, 31, 21, 59, 59))]
        out = period_offset_fix.transform(rows)
        assert out[0]["Period"].year == 2026

    def test_date_with_2000_offset_year_subtracted(self):
        rows = [_row(dt.date(4026, 2, 15))]
        out = period_offset_fix.transform(rows)
        assert out[0]["Period"].year == 2026

    def test_datetime_below_threshold_unchanged(self):
        rows = [_row(dt.datetime(2026, 3, 1, 12, 0, 0))]
        out = period_offset_fix.transform(rows, strip_time=False)
        assert out[0]["Period"] == dt.datetime(2026, 3, 1, 12, 0, 0)

    def test_none_passthrough(self):
        rows = [_row(None)]
        out = period_offset_fix.transform(rows)
        assert out[0]["Period"] is None
        assert out[0]["Period_Month"] is None


class TestStripTime:
    def test_strip_time_default_true_zeroes_time_component(self):
        rows = [_row(dt.datetime(4026, 1, 31, 21, 59, 59, 123456))]
        out = period_offset_fix.transform(rows)
        result = out[0]["Period"]
        assert isinstance(result, dt.datetime)
        assert result.hour == 0
        assert result.minute == 0
        assert result.second == 0
        assert result.microsecond == 0
        assert result.year == 2026
        assert result.month == 1
        assert result.day == 31

    def test_strip_time_explicit_true(self):
        rows = [_row(dt.datetime(4026, 5, 6, 14, 30, 0))]
        out = period_offset_fix.transform(rows, strip_time=True)
        assert out[0]["Period"] == dt.datetime(2026, 5, 6, 0, 0, 0)

    def test_strip_time_false_preserves_time(self):
        rows = [_row(dt.datetime(4026, 5, 6, 14, 30, 0))]
        out = period_offset_fix.transform(rows, strip_time=False)
        assert out[0]["Period"] == dt.datetime(2026, 5, 6, 14, 30, 0)

    def test_strip_time_on_date_object_no_op(self):
        rows = [_row(dt.date(4026, 5, 6))]
        out = period_offset_fix.transform(rows)
        assert out[0]["Period"] == dt.date(2026, 5, 6)

    def test_strip_time_works_for_below_threshold_datetime(self):
        rows = [_row(dt.datetime(2026, 5, 6, 14, 30, 0))]
        out = period_offset_fix.transform(rows, strip_time=True)
        assert out[0]["Period"] == dt.datetime(2026, 5, 6, 0, 0, 0)


class TestPeriodMonthDerivation:
    def test_period_month_first_day_of_month(self):
        rows = [_row(dt.datetime(4026, 5, 6, 14, 30, 0))]
        out = period_offset_fix.transform(rows)
        assert out[0]["Period_Month"] == dt.date(2026, 5, 1)

    def test_period_month_january(self):
        rows = [_row(dt.datetime(4026, 1, 31, 23, 59, 59))]
        out = period_offset_fix.transform(rows)
        assert out[0]["Period_Month"] == dt.date(2026, 1, 1)

    def test_period_month_disabled(self):
        rows = [_row(dt.datetime(4026, 5, 6, 14, 30, 0))]
        out = period_offset_fix.transform(rows, derive_period_month=False)
        assert "Period_Month" not in out[0]


class TestEndToEndScenario:
    def test_real_world_baserp_row(self):
        """1С backend: 4026-02-15 21:59:59.999 → ETL → 2026-02-15 00:00:00.000."""
        rows = [
            {"Period": dt.datetime(4026, 2, 15, 21, 59, 59, 999000), "Sum_Excel": 100.0}
        ]
        out = period_offset_fix.transform(rows)
        assert out[0]["Period"] == dt.datetime(2026, 2, 15, 0, 0, 0)
        assert out[0]["Period_Month"] == dt.date(2026, 2, 1)
        assert out[0]["Sum_Excel"] == 100.0

    def test_multiple_rows(self):
        rows = [
            _row(dt.datetime(4026, 1, 1, 12, 0, 0)),
            _row(dt.datetime(4026, 6, 30, 18, 30, 0)),
            _row(dt.datetime(4027, 12, 31, 23, 59, 59)),
        ]
        out = period_offset_fix.transform(rows)
        assert out[0]["Period"] == dt.datetime(2026, 1, 1)
        assert out[1]["Period"] == dt.datetime(2026, 6, 30)
        assert out[2]["Period"] == dt.datetime(2027, 12, 31)
        for r in out:
            assert r["Period"].hour == 0
            assert r["Period"].minute == 0
            assert r["Period"].second == 0
