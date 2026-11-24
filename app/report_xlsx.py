"""The Excel workbook: a branded, dashboard-style first sheet and consistent, filterable tables.

Real numbers with real currency formats (pivot-ready), native charts, data bars on cost columns.
Presentation only: the figures all come from `report.collect`.
"""
from __future__ import annotations

import io
from datetime import datetime
from typing import Any

NAVY, NAVY_2, BLUE, CYAN = "0B1F3A", "15325C", "0078D4", "00B7C3"
INK, MUTED, LINE, BG, ZEBRA = "1B2733", "66758A", "E3E8EF", "F4F7FB", "F8FAFD"
AREA_COLOURS = ["0078D4", "00B7C3", "6B5BD2", "F2A900", "1E9E6A", "E5584D", "8E9BAE", "B146C2"]


def to_xlsx(data: dict[str, Any]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
    from openpyxl.chart.series import DataPoint
    from openpyxl.formatting.rule import DataBarRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins

    from .report import _money_format, live_tables, period_label

    wb = Workbook()
    cur = data["currency"]
    money = _money_format(cur)
    fill = lambda c: PatternFill("solid", fgColor=c)  # noqa: E731
    hair = Border(bottom=Side(style="thin", color=LINE))
    scope_label = (f"{data['subscription_count']} subscription(s)" if data["scoped"]
                   else "All accessible subscriptions")

    def style_axes(chart, label_skip: int = 1) -> None:
        """Visible axes, light gridlines and no legend: a clean single-series chart."""
        from openpyxl.chart.shapes import GraphicalProperties
        from openpyxl.drawing.line import LineProperties

        chart.legend = None
        for axis in (chart.x_axis, chart.y_axis):
            axis.delete = False
        chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E9EDF3", w=6350))
        if label_skip > 1:
            chart.x_axis.tickLblSkip = label_skip

    def finish_sheet(ws, tab_colour: str, print_title_row: int | None = None) -> None:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = tab_colour
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6)
        ws.oddFooter.right.text = "Page &P of &N"
        ws.oddFooter.left.text = "CloudLens · Azure cost report"
        ws.oddFooter.left.size = ws.oddFooter.right.size = 8
        if print_title_row:
            ws.print_title_rows = f"{print_title_row}:{print_title_row}"

    def banner(ws, title: str, subtitle: str, last_col: int) -> None:
        """Navy title band across the sheet, with a cyan rule underneath."""
        for r in (1, 2, 3):
            for c in range(1, last_col + 1):
                ws.cell(row=r, column=c).fill = fill(NAVY)
        ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=last_col)
        cell = ws.cell(row=1, column=1, value=title)
        cell.font = Font(name="Segoe UI", bold=True, size=20, color="FFFFFF")
        cell.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)
        sub = ws.cell(row=3, column=1, value=subtitle)
        sub.font = Font(name="Segoe UI", size=10.5, color="B9C7DC")
        sub.alignment = Alignment(vertical="center", indent=1)
        for c in range(1, last_col + 1):
            ws.cell(row=4, column=c).fill = fill(CYAN)
        ws.row_dimensions[1].height = ws.row_dimensions[2].height = 20
        ws.row_dimensions[3].height = 20
        ws.row_dimensions[4].height = 3

    def tile(ws, row: int, col: int, span: int, label: str, value, fmt: str | None, sub: str,
             accent: str = BLUE, sub_colour: str = MUTED) -> None:
        """A KPI card: label, big number, one-line context, accent rule on top."""
        for r in range(row, row + 4):
            for c in range(col, col + span):
                ws.cell(row=r, column=c).fill = fill("FFFFFF")
                edge = Side(style="thin", color=LINE)
                ws.cell(row=r, column=c).border = Border(
                    left=edge if c == col else None, right=edge if c == col + span - 1 else None,
                    top=Side(style="thick", color=accent) if r == row else None,
                    bottom=edge if r == row + 3 else None)
        for r, content, font in (
            (row, None, None),
            (row + 1, label.upper(), Font(name="Segoe UI", bold=True, size=8.5, color=MUTED)),
            (row + 2, value, Font(name="Segoe UI", bold=True, size=20, color=INK)),
            (row + 3, sub, Font(name="Segoe UI", size=9.5, color=sub_colour)),
        ):
            if content is None:
                continue
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + span - 1)
            cell = ws.cell(row=r, column=col, value=content)
            cell.font = font
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            if r == row + 2 and fmt:
                cell.number_format = fmt
        ws.row_dimensions[row + 2].height = 30

    def section_title(ws, row: int, title: str, col: int = 1) -> None:
        cell = ws.cell(row=row, column=col, value=title)
        cell.font = Font(name="Segoe UI", bold=True, size=12, color=NAVY_2)

    def write_table(ws, rows: list[list], headers: list[str], start: int = 1,
                    money_cols: tuple[int, ...] = (), freeze: bool = False,
                    date_cols: tuple[int, ...] = (), bars: bool = True) -> int:
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=start, column=c, value=h)
            cell.font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=10)
            cell.fill = fill(NAVY_2)
            cell.alignment = Alignment(horizontal="right" if c in money_cols else "left",
                                       vertical="center", indent=0 if c in money_cols else 1)
        ws.row_dimensions[start].height = 22
        for r, row in enumerate(rows, start=start + 1):
            for c, v in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = Font(name="Segoe UI", size=10, color=INK)
                if (r - start) % 2 == 0:
                    cell.fill = fill(ZEBRA)
                cell.border = hair
                cell.alignment = Alignment(horizontal="right" if c in money_cols else "left",
                                           indent=0 if c in money_cols else 1)
                if c in date_cols and isinstance(v, str):
                    try:
                        cell.value = datetime.strptime(v, "%Y-%m-%d").date()
                        cell.number_format = "yyyy-mm-dd"
                    except ValueError:
                        cell.data_type = "s"
                elif isinstance(v, str):
                    cell.data_type = "s"
                if c in money_cols:
                    cell.number_format = money
        for c, h in enumerate(headers, start=1):
            longest = max([len(str(h))] + [len(str(row[c - 1])) for row in rows[:200]] or [8])
            want = min(max(longest + 4, 12), 46)
            letter = get_column_letter(c)
            ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, want)
        if rows and bars and money_cols:
            col = get_column_letter(money_cols[0])
            ws.conditional_formatting.add(
                f"{col}{start + 1}:{col}{start + len(rows)}",
                DataBarRule(start_type="num", start_value=0, end_type="max", color="9CC9F0",
                            showValue=True))
        if rows and freeze:
            ws.auto_filter.ref = f"A{start}:{get_column_letter(len(headers))}{start + len(rows)}"
            ws.freeze_panes = ws.cell(row=start + 1, column=1)
        return start + len(rows) + 1

    areas = data["areas"]
    sections = data["sections"]
    total = data["total"] or 0.0

    # ------------------------------------------------------------------ Summary
    ws = wb.active
    ws.title = "Summary"
    banner(ws, "CloudLens · Azure cost report", f"{period_label(data)}   |   {scope_label}   |   {cur}", 12)
    for letter, width in zip("ABCDEFGHIJKL", (22, 15, 12, 4, 16, 16, 16, 4, 16, 16, 16, 16)):
        ws.column_dimensions[letter].width = width

    daily: dict[str, float] = {}
    for s in sections:
        trend = s.get("trend") or {}
        for label, value in zip(trend.get("labels", []), trend.get("values", [])):
            daily[label] = daily.get(label, 0.0) + (value or 0.0)
    days_sorted = sorted(daily)
    avg_day = (sum(daily.values()) / len(daily)) if daily else (total / max(1, data["days"]))
    lead = max(areas, key=lambda a: a["cost"]) if areas else None
    previous = sum((s["kpis"].get("previous") or 0.0) for s in sections)
    change = ((sum(s["kpis"]["total"] or 0 for s in sections) - previous) / previous * 100
              if previous and not data["partial"] else None)
    change_text = (f"{'▲' if change > 0 else '▼'} {abs(change):.1f}% vs previous period"
                   if change is not None else cur)

    tile(ws, 6, 1, 3, "Total spend", total, money, change_text, BLUE,
         "E5584D" if (change or 0) > 0.05 else "1E9E6A" if (change or 0) < -0.05 else MUTED)
    tile(ws, 6, 5, 3, "Daily average", avg_day, money, f"over {len(daily) or data['days']} days", CYAN)
    tile(ws, 6, 9, 4, "Largest area", lead["label"] if lead else "—", None,
         f"{lead['cost'] / total * 100:.1f}% of spend" if lead and total else "", "6B5BD2")
    ws.row_dimensions[5].height = 10

    section_title(ws, 11, "Spend by area")
    rows = [[a["label"], a["cost"], (a["cost"] / total) if total else 0]
            for a in sorted(areas, key=lambda a: -a["cost"])]
    end = write_table(ws, rows, ["Area", f"Cost ({cur})", "Share"], start=12, money_cols=(2,))
    for r in range(13, 13 + len(rows)):
        ws.cell(row=r, column=3).number_format = "0.0%"
    if rows:
        chart = DoughnutChart()
        chart.title = None
        chart.holeSize = 62
        chart.legend.position = "r"
        chart.height, chart.width = 8.4, 12.5
        chart.add_data(Reference(ws, min_col=2, min_row=12, max_row=12 + len(rows)), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=13, max_row=12 + len(rows)))
        for i in range(len(rows)):
            point = DataPoint(idx=i)
            point.graphicalProperties.solidFill = AREA_COLOURS[i % len(AREA_COLOURS)]
            point.graphicalProperties.line.solidFill = "FFFFFF"
            chart.series[0].dPt.append(point)
        ws.add_chart(chart, "E11")

    next_row = max(end, 11 + 18) + 1
    if len(days_sorted) > 1:
        section_title(ws, next_row, "Daily spend, all areas")
        at = write_table(ws, [[d, round(daily[d], 2)] for d in days_sorted],
                         ["Date", f"Cost ({cur})"], start=next_row + 1, money_cols=(2,),
                         date_cols=(1,), bars=False)
        line = LineChart()
        line.title = None
        line.height, line.width = 8.4, 20
        line.y_axis.numFmt = money
        style_axes(line, max(1, len(days_sorted) // 8))
        line.add_data(Reference(ws, min_col=2, min_row=next_row + 1, max_row=next_row + 1 + len(days_sorted)),
                      titles_from_data=True)
        line.set_categories(Reference(ws, min_col=1, min_row=next_row + 2, max_row=next_row + 1 + len(days_sorted)))
        line.series[0].graphicalProperties.line.solidFill = BLUE
        line.series[0].graphicalProperties.line.width = 28000
        line.series[0].smooth = False
        line.legend = None
        ws.add_chart(line, f"E{next_row + 1}")
        next_row = max(at, next_row + 19) + 1

    section_title(ws, next_row, "About this report")
    meta = [("Generated", data["generated"]), ("Period", period_label(data)),
            ("Data through", data["as_of"] or "—"), ("Currency", cur), ("Scope", scope_label)]
    for i, (k, v) in enumerate(meta, start=next_row + 1):
        ws.cell(row=i, column=1, value=k).font = Font(name="Segoe UI", bold=True, size=10, color=MUTED)
        ws.cell(row=i, column=2, value=v).font = Font(name="Segoe UI", size=10, color=INK)
    note = ws.cell(row=next_row + len(meta) + 2, column=1,
                   value="Costs are billed amounts as recorded in Azure cost exports. They are not an invoice.")
    note.font = Font(name="Segoe UI", italic=True, size=9.5, color=MUTED)
    finish_sheet(ws, BLUE)

    # ------------------------------------------------------------ one sheet per area
    colour_for = {a["id"]: AREA_COLOURS[i % len(AREA_COLOURS)] for i, a in enumerate(areas)}
    for index, section in enumerate(sections):
        colour = colour_for.get(section.get("id"), AREA_COLOURS[index % len(AREA_COLOURS)])
        sheet = wb.create_sheet(section["label"][:31].replace("/", "-").replace("&", "and"))
        banner(sheet, section["label"], period_label(data), 9)
        for letter, width in zip("ABCDEFGHI", (30, 24, 24, 18, 16, 4, 16, 16, 16)):
            sheet.column_dimensions[letter].width = width
        k = section["kpis"]
        cp = k.get("change_pct")
        sub = (f"{'▲' if cp > 0 else '▼'} {abs(cp):.1f}% vs previous period" if cp is not None else "no prior period")
        sub_colour = "E5584D" if (cp or 0) > 0.05 else "1E9E6A" if (cp or 0) < -0.05 else MUTED
        tile(sheet, 6, 1, 1, "Spend", k["total"], money, sub, colour, sub_colour)
        tile(sheet, 6, 2, 1, "Previous period", k.get("previous"), money, "comparison baseline", "8E9BAE")
        tile(sheet, 6, 3, 1, "Share of total", (k["share_pct"] / 100) if k.get("share_pct") is not None else None,
             "0.0%", "of in-scope spend", colour)
        tile(sheet, 6, 4, 2, "Billed resources", k["resources"], "#,##0", f"{k['services']} services", "8E9BAE")
        sheet.row_dimensions[5].height = 10

        section_title(sheet, 11, "Top resources")
        at = write_table(sheet, [[r["name"], r["grp"], r["service"], r["region"], r["cost"]]
                                 for r in section["resources_top"]],
                         ["Resource", "Resource group", "Service", "Region", f"Cost ({cur})"],
                         start=12, money_cols=(5,), freeze=True)
        section_title(sheet, at + 1, "Services")
        at = write_table(sheet, [[s["name"], s["cost"]] for s in section["services"]],
                         ["Service", f"Cost ({cur})"], start=at + 2, money_cols=(2,))
        section_title(sheet, at + 1, "Regions")
        at = write_table(sheet, [[r["name"], r["cost"]] for r in section["regions"]],
                         ["Region", f"Cost ({cur})"], start=at + 2, money_cols=(2,))
        trend = section.get("trend") or {}
        if trend.get("labels"):
            section_title(sheet, at + 1, "Daily spend")
            at = write_table(sheet, [list(pair) for pair in zip(trend["labels"], trend["values"])],
                             ["Date", f"Cost ({cur})"], start=at + 2, money_cols=(2,), date_cols=(1,),
                             bars=False)
            first = at - len(trend["labels"]) - 1
            chart = BarChart()
            chart.type = "col"
            chart.title = None
            chart.height, chart.width = 8.4, 18
            chart.y_axis.numFmt = money
            style_axes(chart, max(1, len(trend["labels"]) // 8))
            chart.gapWidth = 40
            chart.add_data(Reference(sheet, min_col=2, min_row=first, max_row=first + len(trend["labels"])),
                           titles_from_data=True)
            chart.set_categories(Reference(sheet, min_col=1, min_row=first + 1,
                                           max_row=first + len(trend["labels"])))
            chart.series[0].graphicalProperties.solidFill = colour
            sheet.add_chart(chart, f"C{first}")
        finish_sheet(sheet, colour)

    # ------------------------------------------------------------- live findings
    for t in live_tables(data):
        sheet = wb.create_sheet(t["title"][:31].replace("/", "-").replace("&", "and"))
        banner(sheet, t["title"], " · ".join(t["facts"]), max(6, len(t["columns"])))
        section_title(sheet, 6, "Findings")
        write_table(sheet, t["rows"], t["columns"], start=7, money_cols=(t["money_col"],), freeze=True)
        finish_sheet(sheet, "F2A900", print_title_row=7)

    wb.properties.title = "CloudLens — Azure cost report"
    wb.properties.subject = period_label(data)
    wb.properties.creator = "CloudLens"
    wb.properties.lastModifiedBy = "CloudLens"
    wb.properties.category = "Cost report"
    wb.properties.description = (f"Azure cost report. {period_label(data)}. Currency {cur}. {scope_label}.")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
