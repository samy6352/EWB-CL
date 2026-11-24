"""Exporting the dashboard as a report someone can send to their finance team.

Three formats, chosen because they are what people actually do with cost data:

- **XLSX** — the one finance asks for. A branded dashboard sheet, one sheet per area, real
  numbers with currency formatting so it can be pivoted rather than retyped. Built in
  `report_xlsx`.
- **PPTX** — the one that goes to a steering committee: an executive summary, key takeaways,
  native charts and ranked breakdowns. Native PowerPoint charts, not pasted images, so the deck
  can be restyled without coming back here. Built in `report_pptx`.
- **CSV** — the lowest common denominator; imports anywhere.

PDF is deliberately absent: the browser already prints this page well (there are print styles),
and a headless renderer on a Free-tier App Service would be a large dependency producing a
worse result than Ctrl-P.

Markdown and JSON were removed: neither is a format a finance or steering-committee audience
asks for, and carrying two more builders meant two more things to keep correct for no reader.

Everything is generated in memory and streamed. Writing a temp file would be a needless
failure mode on a 1 GB instance, and reports are small.
"""
from __future__ import annotations

import io
import logging
from datetime import UTC, datetime
from typing import Any

from .currency import symbol as fx_symbol

log = logging.getLogger("cloudlens.report")

FORMATS = {
    "xlsx": ("Excel dashboard workbook",
             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    "pptx": ("PowerPoint executive deck",
             "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
    "csv": ("CSV", "text/csv"),
}

# Which parts of an area section a report can carry. Selected individually because a deck for
# a steering committee wants the trend and nothing else, while a handover to an engineer wants
# the resource list and could do without the chart.
BLOCKS = ("trend", "services", "regions", "resources")

# Live datasets. Separate from the warehouse sections because they call Azure — including one
# turns a sub-second export into a slow one, so it has to be a deliberate choice.
LIVE = {
    "waste": "Orphaned resources",
    "rightsizing": "Rightsizing",
    "esu": "End of support",
    "advisor": "Advisor",
}


def default_selection() -> dict[str, Any]:
    """Everything from the warehouse, nothing live. The safe, fast default."""
    return {"summary": True, "sections": None, "blocks": list(BLOCKS), "live": []}


def collect(scope: list[str] | None, days: int = 30,
            selection: dict[str, Any] | None = None,
            live_data: dict[str, Any] | None = None,
            currency: str | None = None) -> dict[str, Any]:
    """Everything a report needs, filtered to what was asked for.

    Deliberately the same `Dashboard` the screen uses: a report that disagreed with the page
    it was exported from would be worse than no report. `currency` is part of that agreement —
    a report ordered while the screen showed rupees has to contain rupees.

    `selection` chooses content:
      summary  — include the KPIs and the by-area table
      sections — list of area ids, or None for all
      blocks   — which parts of each area (trend / services / regions / resources)
      live     — which live datasets to embed, fetched by the caller and passed in `live_data`
    """
    from . import reporting
    from .dashboard import get_dashboard

    days = reporting.period_days(days)
    sel = {**default_selection(), **(selection or {})}
    blocks = set(sel.get("blocks") or BLOCKS)
    wanted = sel.get("sections")

    d = get_dashboard()
    period = reporting.metadata(d.wh, scope, days)
    summary = d.sections(scope, days=days, currency=currency)
    areas = [s for s in summary.get("sections", []) if s.get("cost")]
    if wanted is not None:
        keep = set(wanted)
        areas = [a for a in areas if a["id"] in keep]

    sections = []
    for area in areas:
        try:
            section = d.section(area["id"], scope, days=days, currency=currency)
        except KeyError:  # a group with no rows in this window
            continue
        # Drop the blocks that were not asked for, rather than writing empty sheets and slides.
        if "trend" not in blocks:
            section["trend"] = {"labels": [], "values": []}
        for name in ("services", "regions"):
            if name not in blocks:
                section[name] = []
        if "resources" not in blocks:
            section["resources_top"] = []
        sections.append(section)

    return {
        "generated": datetime.now(UTC).strftime("%Y-%m-%d %H:%M UTC"),
        "days": days,
        "reporting_period": period,
        "as_of": summary.get("as_of"),
        "currency": summary.get("currency") or "USD",
        # The headline total always describes the whole estate in scope, even when only some
        # areas are exported — otherwise a partial report looks like a smaller bill.
        "total": summary.get("total") or 0.0,
        "areas_total": round(sum(a["cost"] for a in areas), 2),
        "mixed_currency": summary.get("mixed"),
        "areas": areas if sel.get("summary") else [],
        "sections": sections,
        "live": live_data or {},
        "partial": wanted is not None or set(blocks) != set(BLOCKS) or not sel.get("summary"),
        "scoped": bool(scope),
        "subscription_count": len(scope) if scope else None,
    }


def period_label(data: dict[str, Any]) -> str:
    period = data.get("reporting_period") or {}
    if period.get("from") and period.get("to"):
        kind = period.get("month") or f"Rolling {data['days']} days"
        return f"{kind}: {period['from']} to {period['to']} (inclusive)"
    return f"Last {data['days']} days (no cost data)"


def _money_format(currency: str) -> str:
    mark = fx_symbol(currency)
    return f'"{mark}"#,##0.00' if mark else "#,##0.00"


def live_tables(data: dict[str, Any]) -> list[dict[str, Any]]:
    """Flatten whatever live datasets were included into plain titled tables.

    Each live tool returns its own shape; every export format wants the same thing — a title,
    some headline facts, column names and rows. Normalising once here means the four writers
    below stay simple and cannot disagree about what a finding looks like.
    """
    out: list[dict[str, Any]] = []
    live = data.get("live") or {}
    cur = data["currency"]

    waste = live.get("waste")
    if waste:
        for finding in waste.get("findings", []):
            if not finding.get("count"):
                continue
            out.append({
                "key": "waste",
                "title": str(finding.get("category", "finding")).replace("_", " ").title(),
                "facts": [
                    f"{finding['count']} item(s)",
                    f"{finding.get('cost') or 0:,.2f} {cur} over {waste.get('period_days', '?')} days",
                ],
                "columns": ["Resource", "Resource group", "Region", f"Cost ({cur})"],
                "rows": [[i.get("name") or i.get("id") or "—", i.get("resource_group") or "—",
                          i.get("location") or i.get("region") or "—", i.get("cost")]
                         for i in (finding.get("items") or [])],
                "money_col": 4,
            })

    rs = live.get("rightsizing")
    if rs:
        out.append({
            "key": "rightsizing",
            "title": "Rightsizing",
            "facts": [f"{rs.get('count', 0)} VM(s)", f"{rs.get('idle_count', 0)} idle",
                      f"idle cost {rs.get('idle_cost') or 0:,.2f} {cur}"],
            "columns": ["Name", "Size", "Power state", "CPU avg %", "CPU peak %", f"Cost ({cur})"],
            "rows": [[v.get("name"), v.get("size"), v.get("state"),
                      v.get("cpu_avg"), v.get("cpu_peak"), v.get("cost")]
                     for v in (rs.get("vms") or [])],
            "money_col": 6,
        })

    esu = live.get("esu")
    if esu:
        out.append({
            "key": "esu",
            "title": "End of support",
            "facts": [f"{esu.get('out_of_support', 0)} out of support",
                      f"{esu.get('ending_soon', 0)} ending soon",
                      f"{esu.get('exposed', 0)} uncovered",
                      f"estimated {esu.get('estimated_monthly_cost') or 0:,.2f} {cur}/month"],
            "columns": ["Name", "Product", "Where", "Support ended", "ESU position",
                        f"Est. per month ({cur})"],
            "rows": [[m.get("name"), m.get("product"), m.get("kind"), m.get("support_ends"),
                      m.get("coverage"), m.get("monthly_esu_cost")]
                     for m in (esu.get("machines") or [])],
            "money_col": 6,
        })

    advisor = live.get("advisor")
    if advisor:
        out.append({
            "key": "advisor",
            "title": "Advisor recommendations",
            "facts": [f"{advisor.get('count', 0)} recommendation(s)",
                      f"estimated {advisor.get('estimated_annual_savings') or 0:,.2f} "
                      f"{cur}/year"],
            "columns": ["Impact", "Recommendation", "Resource", f"Annual saving ({cur})"],
            "rows": [[r.get("impact"), r.get("problem") or r.get("solution"),
                      r.get("resource"), r.get("annual_savings")]
                     for r in (advisor.get("recommendations") or [])],
            "money_col": 4,
        })

    return [t for t in out if t["rows"]]


def to_xlsx(data: dict[str, Any]) -> bytes:
    from .report_xlsx import to_xlsx as build

    return build(data)


def to_pptx(data: dict[str, Any]) -> bytes:
    from .report_pptx import to_pptx as build

    return build(data)


# ------------------------------------------------------------------------------ csv
def to_csv(data: dict[str, Any]) -> bytes:
    """One spreadsheet-safe flat table.

    A leading provenance block, because this file outlives the conversation that produced it.
    A bare grid of numbers in a downloads folder cannot say which subscriptions it covers, which
    dates, or which currency — and a cost figure whose scope is unknown is not evidence. The
    comment rows are prefixed with `#` so a pandas or Power Query import can skip them.
    """
    from .spreadsheet import SafeCsvWriter

    out = io.StringIO()
    w = SafeCsvWriter(out, lineterminator="\n")
    cur = data["currency"]
    scope = (f"{data['subscription_count']} subscription(s)" if data["scoped"]
             else "All accessible subscriptions")
    for line in (
        "# CloudLens - Azure cost report",
        f"# Period: {period_label(data)}",
        f"# Data through: {data['as_of'] or 'not available'}",
        f"# Scope: {scope}",
        f"# Currency: {cur}",
        f"# Generated: {data['generated']}",
    ):
        w.writerow([line])
    w.writerow([])
    w.writerow(["Area", "Type", "Name", "Resource group", "Service", "Region", f"Cost ({cur})"])
    for a in data["areas"]:
        w.writerow([a["label"], "area total", "", "", "", "", f"{a['cost']:.2f}"])
    for section in data["sections"]:
        for s in section["services"]:
            w.writerow([section["label"], "service", s["name"], "", "", "", f"{s['cost']:.2f}"])
        for r in section["regions"]:
            w.writerow([section["label"], "region", r["name"], "", "", "", f"{r['cost']:.2f}"])
        for r in section["resources_top"]:
            w.writerow([section["label"], "resource", r["name"], r["grp"], r["service"],
                        r["region"], f"{r['cost']:.2f}"])
    for t in live_tables(data):
        for row in t["rows"]:
            cost = row[t["money_col"] - 1]
            w.writerow([t["title"], t["key"], str(row[0] or ""), "", "", "",
                        f"{cost:.2f}" if isinstance(cost, (int, float)) else ""])
    return out.getvalue().encode("utf-8-sig")  # BOM so Excel opens UTF-8 correctly


BUILDERS = {"xlsx": to_xlsx, "pptx": to_pptx, "csv": to_csv}


def build(fmt: str, scope: list[str] | None, days: int = 30,
          selection: dict[str, Any] | None = None,
          live_data: dict[str, Any] | None = None,
          currency: str | None = None) -> tuple[bytes, str, str]:
    """Returns (bytes, filename, media type)."""
    if fmt not in BUILDERS:
        raise KeyError(fmt)
    data = collect(scope, days=days, selection=selection, live_data=live_data,
                   currency=currency)
    payload = BUILDERS[fmt](data)
    stamp = datetime.now(UTC).strftime("%Y%m%d")
    period = data.get("reporting_period") or {}
    if period.get("from") and period.get("to"):
        stamp = f"{period['from']}_to_{period['to']}"
    # The currency goes in the filename when it is not the billed one. Two files a week apart,
    # one in dollars and one in rupees, are otherwise indistinguishable in a downloads folder.
    tag = f"-{currency.lower()}" if currency and currency.upper() != "BILLED" else ""
    return payload, f"cloudlens-cost-report-{stamp}{tag}.{fmt}", FORMATS[fmt][1]
