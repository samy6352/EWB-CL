"""The exported deck and workbook open cleanly, including for sparse selections."""
# cloudlens-test: requires-offline
from __future__ import annotations

import io

import pytest

from app import report

pytestmark = pytest.mark.offline


def _section(aid: str, label: str, cost: float, change: float | None) -> dict:
    labels = [f"2026-11-{d:02d}" for d in range(1, 11)]
    return {
        "id": aid, "label": label,
        "kpis": {"total": cost, "previous": cost / 2, "change_pct": change, "share_pct": 50.0,
                 "resources": 12, "services": 2},
        "services": [{"name": "Virtual Machines", "cost": cost * 0.7}, {"name": "Disks", "cost": cost * 0.3}],
        "regions": [{"name": "East US", "cost": cost}],
        "trend": {"labels": labels, "values": [cost / 10] * 10},
        "resources_top": [{"name": "vm-1", "grp": "rg-a", "service": "Virtual Machines",
                           "region": "East US", "cost": cost * 0.4}],
    }


def _data(**over) -> dict:
    data = {
        "generated": "2026-11-24 10:00 UTC", "days": 30,
        "reporting_period": {"from": "2026-11-01", "to": "2026-11-10", "month": None},
        "as_of": "2026-11-10", "currency": "USD", "total": 300.0, "areas_total": 300.0,
        "mixed_currency": False,
        "areas": [{"id": "a", "label": "Compute", "cost": 200.0}, {"id": "b", "label": "Storage", "cost": 100.0}],
        "sections": [_section("a", "Compute", 200.0, 12.5), _section("b", "Storage", 100.0, -8.0)],
        "live": {}, "partial": False, "scoped": True, "subscription_count": 2,
    }
    return {**data, **over}


def test_deck_and_workbook_build_and_open():
    from openpyxl import load_workbook
    from pptx import Presentation

    data = _data()
    deck = Presentation(io.BytesIO(report.to_pptx(data)))
    assert len(deck.slides) >= 6  # cover, summary, breakdown, two areas, about
    book = load_workbook(io.BytesIO(report.to_xlsx(data)))
    assert book.sheetnames[0] == "Summary" and {"Compute", "Storage"} <= set(book.sheetnames)
    assert book["Summary"]["A1"].value.startswith("CloudLens")


@pytest.mark.parametrize("over", [
    {"areas": [], "sections": [], "partial": True},
    {"sections": [{**_section("a", "Compute", 200.0, None), "trend": {"labels": [], "values": []},
                   "services": [], "regions": [], "resources_top": []}], "partial": True},
    {"as_of": None, "reporting_period": None, "scoped": False, "subscription_count": None},
])
def test_sparse_selections_still_build(over):
    from openpyxl import load_workbook
    from pptx import Presentation

    data = _data(**over)
    assert len(Presentation(io.BytesIO(report.to_pptx(data))).slides) >= 2
    assert load_workbook(io.BytesIO(report.to_xlsx(data))).sheetnames
