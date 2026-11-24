"""The PowerPoint deck: an executive-ready design, still built from native, editable objects.

Charts are real PowerPoint charts (restylable, not pictures); every card, bar and tag is a plain
shape. Colours, type and spacing come from one palette below so the deck reads as a single piece.
Presentation only: the numbers all come from `report.collect`.
"""
from __future__ import annotations

import io
from typing import Any

from .currency import symbol as fx_symbol

NAVY, NAVY_2 = "0B1F3A", "15325C"
BLUE, CYAN, VIOLET = "0078D4", "00B7C3", "6B5BD2"
AMBER, GREEN, CORAL, SLATE, ORCHID = "F2A900", "1E9E6A", "E5584D", "8E9BAE", "B146C2"
INK, MUTED, LINE, BG, WHITE = "1B2733", "66758A", "E3E8EF", "F4F7FB", "FFFFFF"
AREA_COLOURS = [BLUE, CYAN, VIOLET, AMBER, GREEN, CORAL, SLATE, ORCHID]
FONT = "Segoe UI"


def to_pptx(data: dict[str, Any]) -> bytes:
    from lxml import etree
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Emu, Inches, Pt

    from .report import live_tables, period_label

    cur = data["currency"]
    symbol = fx_symbol(cur)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    # ------------------------------------------------------------------ helpers
    def rgb(hex_: str) -> RGBColor:
        return RGBColor.from_string(hex_)

    def money(v, cents: bool = True) -> str:
        if v is None:
            return "—"
        return f"{symbol}{v:,.2f}" if cents else f"{symbol}{v:,.0f}"

    def compact(v) -> str:
        if v is None:
            return "—"
        a = abs(v)
        if a >= 1e6:
            return f"{symbol}{v / 1e6:,.2f}M"
        if a >= 1e3:
            return f"{symbol}{v / 1e3:,.2f}K"
        return f"{symbol}{v:,.2f}"

    def set_alpha(shape, percent: int) -> None:
        clr = shape.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        etree.SubElement(clr, qn("a:alpha")).set("val", str(percent * 1000))

    def rect(slide, x, y, w, h, fill=None, line=None, radius=None, kind=None, line_w=0.75):
        shape = slide.shapes.add_shape(
            kind or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE),
            Inches(x), Inches(y), Inches(w), Inches(h))
        if radius and shape.adjustments:
            shape.adjustments[0] = radius
        if fill:
            shape.fill.solid()
            shape.fill.fore_color.rgb = rgb(fill)
        else:
            shape.fill.background()
        if line:
            shape.line.color.rgb = rgb(line)
            shape.line.width = Pt(line_w)
        else:
            shape.line.fill.background()
        shape.shadow.inherit = False
        return shape

    def text(slide, value, x, y, w, h, size=14, bold=False, colour=INK, align="l",
             anchor="t", spacing=None, line_spacing=None):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE,
                              "b": MSO_ANCHOR.BOTTOM}[anchor]
        p = tf.paragraphs[0]
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = str(value)
        run.font.size, run.font.bold, run.font.name = Pt(size), bold, FONT
        run.font.color.rgb = rgb(colour)
        if spacing:
            run._r.get_or_add_rPr().set("spc", str(spacing))
        return box

    def background(slide, colour):
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(colour)

    def card(slide, x, y, w, h, accent=None):
        rect(slide, x, y, w, h, WHITE, LINE, radius=0.035)
        if accent:
            rect(slide, x + 0.02, y, w - 0.04, 0.07, accent)

    def page_header(slide, title, subtitle="", colour=BLUE):
        background(slide, BG)
        rect(slide, 0.6, 0.48, 0.09, 0.62, colour)
        text(slide, title, 0.85, 0.4, 10.5, 0.55, 28, True, INK, anchor="m")
        if subtitle:
            text(slide, subtitle, 0.85, 0.93, 11.5, 0.3, 12.5, False, MUTED)

    def footer(slide):
        index = len(prs.slides)
        rect(slide, 0.6, 7.0, 12.13, 0.012, LINE)
        text(slide, "CloudLens", 0.6, 7.07, 1.4, 0.25, 9, True, BLUE)
        text(slide, period_label(data), 1.55, 7.07, 9.5, 0.25, 9, False, MUTED)
        text(slide, str(index), 12.23, 7.07, 0.5, 0.25, 9, True, MUTED, align="r")

    def kpi(slide, x, y, w, h, label, value, sub="", accent=BLUE, sub_colour=MUTED):
        card(slide, x, y, w, h, accent)
        text(slide, label.upper(), x + 0.22, y + 0.2, w - 0.4, 0.25, 9, True, MUTED, spacing=90)
        value = str(value)
        size = 24 if len(value) <= 9 else 20 if len(value) <= 13 else 15
        text(slide, value, x + 0.22, y + 0.46, w - 0.4, 0.42, size, True, INK, anchor="m")
        if sub:
            text(slide, sub, x + 0.22, y + h - 0.35, w - 0.4, 0.28, 10.5, False, sub_colour)

    def delta(pct):
        """Colour and arrow for a spend change. Spend going up is the thing to notice."""
        if pct is None:
            return "No prior period", MUTED
        arrow = "▲" if pct > 0.05 else "▼" if pct < -0.05 else "■"
        colour = CORAL if pct > 0.05 else GREEN if pct < -0.05 else MUTED
        return f"{arrow} {abs(pct):.1f}% vs previous period", colour

    def bar_list(slide, x, y, w, items, colour, rows=7, row_h=0.46, show_share=False, total=None):
        """Ranked horizontal bars drawn from shapes: clearer than a table for a top-N list."""
        items = [(n, v) for n, v in items if v is not None][:rows]
        if not items:
            return
        top = max(v for _, v in items) or 1
        for i, (name, value) in enumerate(items):
            yy = y + i * row_h
            label = name if len(name) <= 34 else name[:33] + "…"
            text(slide, label, x, yy, w - 1.5, 0.22, 10.5, False, INK)
            share = f"  ·  {value / total * 100:.1f}%" if show_share and total else ""
            text(slide, money(value) + share, x + w - 2.2, yy, 2.2, 0.22, 10.5, True, INK, align="r")
            rect(slide, x, yy + 0.25, w, 0.09, "EEF2F7", radius=0.5)
            rect(slide, x, yy + 0.25, max(0.05, w * (value / top)), 0.09, colour, radius=0.5)

    def thin_labels(labels, target=8):
        """MM-DD labels with all but every Nth blanked, so a month of dates stays legible."""
        step = max(1, len(labels) // target)
        return [(lab[5:] if len(lab) >= 10 and lab[4] == "-" else lab) if i % step == 0 else " "
                for i, lab in enumerate(labels)]

    def clean_axes(chart, money_axis=True, skip=None):
        chart.has_title = False
        chart.font.size, chart.font.name = Pt(10), FONT
        chart.font.color.rgb = rgb(MUTED)
        try:
            va, ca = chart.value_axis, chart.category_axis
            va.major_gridlines.format.line.color.rgb = rgb("E9EDF3")
            va.major_gridlines.format.line.width = Pt(0.5)
            va.format.line.fill.background()
            ca.format.line.color.rgb = rgb(LINE)
            ca.has_major_gridlines = False
            va.tick_labels.font.size = ca.tick_labels.font.size = Pt(9.5)
            if money_axis:
                va.tick_labels.number_format = f'"{symbol}"#,##0' if symbol else "#,##0"
                va.tick_labels.number_format_is_linked = False
            if skip:
                ca.tick_label_skip = skip
        except (AttributeError, NotImplementedError):
            pass

    def style_table(slide, columns, rows, x, y, w, col_w, align, row_h=0.36, size=10.5,
                    accent=BLUE, money_cols=()):
        shape = slide.shapes.add_table(len(rows) + 1, len(columns), Inches(x), Inches(y),
                                       Inches(w), Inches(row_h * (len(rows) + 1)))
        table = shape.table
        tbl_pr = table._tbl.tblPr
        for attr in ("bandRow", "firstRow"):
            tbl_pr.set(attr, "0")
        style = tbl_pr.find(qn("a:tableStyleId"))
        if style is not None:
            style.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
        for c, cw in enumerate(col_w):
            table.columns[c].width = Inches(cw)
        for r in range(len(rows) + 1):
            table.rows[r].height = Inches(row_h)

        def put(r, c, value, head=False):
            cell = table.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(accent if head else (WHITE if r % 2 else "F8FAFD"))
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT if align[c] == "r" else PP_ALIGN.LEFT
            run = p.add_run()
            run.text = str(value)
            run.font.size, run.font.name = Pt(size), FONT
            run.font.bold = head or (c in money_cols)
            run.font.color.rgb = rgb(WHITE if head else INK)
            border = etree.SubElement(cell._tc.get_or_add_tcPr(), qn("a:lnB"))
            border.set("w", "6350")
            fill = etree.SubElement(border, qn("a:solidFill"))
            etree.SubElement(fill, qn("a:srgbClr")).set("val", LINE if not head else accent)
            # a:lnB must precede the fill in tcPr; move it first.
            tcpr = cell._tc.tcPr
            tcpr.remove(border)
            tcpr.insert(0, border)

        for c, h in enumerate(columns):
            put(0, c, h, head=True)
        for r, row in enumerate(rows, start=1):
            for c, v in enumerate(row):
                put(r, c, v)
        return table

    # ------------------------------------------------------------------- numbers
    areas = data["areas"]
    sections = data["sections"]
    total = data["total"] or 0.0
    scope_label = (f"{data['subscription_count']} subscription(s)" if data["scoped"]
                   else "All accessible subscriptions")
    area_colour = {a["id"]: AREA_COLOURS[i % len(AREA_COLOURS)] for i, a in enumerate(areas)}
    for i, s in enumerate(sections):
        area_colour.setdefault(s.get("id"), AREA_COLOURS[i % len(AREA_COLOURS)])

    # One daily series for the whole estate: the sum of every area's own series.
    daily: dict[str, float] = {}
    for s in sections:
        trend = s.get("trend") or {}
        for label, value in zip(trend.get("labels", []), trend.get("values", [])):
            daily[label] = daily.get(label, 0.0) + (value or 0.0)
    days_sorted = sorted(daily)
    day_count = len(days_sorted)
    avg_day = (sum(daily.values()) / day_count) if day_count else (total / max(1, data["days"]))
    peak_day = max(daily, key=daily.get) if daily else None
    partial_last = bool(days_sorted and data.get("as_of") and days_sorted[-1] == data["as_of"])
    partial_note = "Latest day is partial and may be incomplete." if partial_last else ""

    previous = sum((s["kpis"].get("previous") or 0.0) for s in sections)
    overall_change = ((sum(s["kpis"]["total"] or 0 for s in sections) - previous) / previous * 100
                      if previous and not data["partial"] else None)

    services: dict[str, float] = {}
    for s in sections:
        for item in s.get("services") or []:
            services[item["name"]] = services.get(item["name"], 0.0) + (item["cost"] or 0.0)
    top_service = max(services.items(), key=lambda kv: kv[1]) if services else None

    movers = [s for s in sections if s["kpis"].get("change_pct") is not None
              and (s["kpis"]["total"] or 0) >= max(1.0, total * 0.01)]
    risers = sorted(movers, key=lambda s: -s["kpis"]["change_pct"])
    fallers = sorted(movers, key=lambda s: s["kpis"]["change_pct"])

    # ---------------------------------------------------------------------- cover
    s = prs.slides.add_slide(blank)
    background(s, NAVY)
    big = rect(s, 8.3, -1.4, 7.4, 7.4, NAVY_2, kind=MSO_SHAPE.OVAL)
    ring = rect(s, 9.6, 1.5, 4.6, 4.6, None, CYAN, kind=MSO_SHAPE.OVAL, line_w=1.5)
    dot = rect(s, 11.7, 5.2, 1.1, 1.1, BLUE, kind=MSO_SHAPE.OVAL)
    set_alpha(dot, 70)
    rect(s, 0.9, 1.55, 0.09, 1.9, CYAN)
    text(s, "CLOUDLENS", 1.25, 1.5, 6, 0.3, 12, True, CYAN, spacing=300)
    text(s, "Azure cost report", 1.25, 1.85, 7.6, 1.0, 46, True, WHITE)
    text(s, period_label(data), 1.25, 2.85, 7.4, 0.4, 15, False, "B9C7DC")
    text(s, scope_label, 1.25, 3.25, 7.4, 0.4, 13, False, "8FA3C0")
    rect(s, 1.25, 4.45, 4.7, 1.45, NAVY_2, "2A4A7A", radius=0.07)
    text(s, "TOTAL SPEND", 1.55, 4.68, 4, 0.25, 9.5, True, "8FA3C0", spacing=120)
    text(s, f"{money(total)}", 1.55, 4.98, 4.2, 0.6, 32, True, WHITE)
    text(s, cur, 1.55, 5.52, 2, 0.28, 11, False, CYAN)
    text(s, f"Generated {data['generated']}", 1.25, 6.75, 7, 0.3, 10.5, False, "8FA3C0")

    # ---------------------------------------------------------- executive summary
    s = prs.slides.add_slide(blank)
    page_header(s, "Executive summary", "The period at a glance")
    change_text, change_colour = delta(overall_change)
    kpi(s, 0.6, 1.45, 2.9, 1.3, "Total spend", compact(total),
        change_text if overall_change is not None else cur, BLUE, change_colour)
    kpi(s, 3.65, 1.45, 2.9, 1.3, "Daily average", compact(avg_day), f"over {day_count or data['days']} days", CYAN)
    kpi(s, 6.7, 1.45, 2.9, 1.3, "Peak day", compact(daily[peak_day]) if peak_day else "—",
        peak_day or "no daily data", VIOLET)
    lead = max(areas, key=lambda a: a["cost"]) if areas else None
    kpi(s, 9.75, 1.45, 2.98, 1.3, "Largest area", lead["label"] if lead else "—",
        f"{lead['cost'] / total * 100:.1f}% of spend" if lead and total else "", AMBER)

    card(s, 0.6, 2.95, 7.55, 3.85)
    text(s, "Daily spend", 0.85, 3.12, 5, 0.3, 13, True, INK)
    if partial_note:
        text(s, partial_note, 4.4, 3.15, 3.6, 0.3, 9.5, False, MUTED, align="r")
    if len(days_sorted) > 1:
        cd = CategoryChartData()
        cd.categories = thin_labels(days_sorted)
        cd.add_series("Spend", [round(daily[d], 2) for d in days_sorted])
        chart = s.shapes.add_chart(XL_CHART_TYPE.AREA, Inches(0.7), Inches(3.45), Inches(7.35),
                                   Inches(3.25), cd).chart
        chart.has_legend = False
        ser = chart.plots[0].series[0]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(BLUE)
        ser.format.line.color.rgb = rgb(BLUE)
        ser.format.line.width = Pt(2)
        clean_axes(chart)
    else:
        text(s, "No daily trend in this selection.", 0.85, 3.6, 6, 0.3, 11, False, MUTED)

    card(s, 8.3, 2.95, 4.43, 3.85)
    text(s, "Key takeaways", 8.55, 3.12, 4, 0.3, 13, True, INK)
    points: list[tuple[str, str]] = []
    if lead and total:
        top3 = sum(a["cost"] for a in sorted(areas, key=lambda a: -a["cost"])[:3])
        points.append((BLUE, f"{lead['label']} leads at {lead['cost'] / total * 100:.0f}% of spend; "
                             f"the top three areas make up {top3 / total * 100:.0f}%."))
    if top_service:
        points.append((CYAN, f"{top_service[0]} is the largest single service at {money(top_service[1], False)}."))
    if risers and risers[0]["kpis"]["change_pct"] > 0.5:
        r = risers[0]
        points.append((CORAL, f"{r['label']} grew fastest: {r['kpis']['change_pct']:+.1f}% vs the previous period."))
    if fallers and fallers[0]["kpis"]["change_pct"] < -0.5:
        f = fallers[0]
        points.append((GREEN, f"{f['label']} fell {abs(f['kpis']['change_pct']):.1f}% vs the previous period."))
    if peak_day:
        points.append((VIOLET, f"Highest daily spend was {money(daily[peak_day], False)} on {peak_day}."))
    for i, (colour, line) in enumerate(points[:5]):
        yy = 3.62 + i * 0.62
        rect(s, 8.58, yy + 0.07, 0.12, 0.12, colour, kind=MSO_SHAPE.OVAL)
        text(s, line, 8.85, yy, 3.7, 0.58, 10.5, False, INK, line_spacing=1.05)
    footer(s)

    # -------------------------------------------------------- where the money goes
    if areas:
        s = prs.slides.add_slide(blank)
        page_header(s, "Where the money goes", "Spend by area, largest first", CYAN)
        ranked = sorted(areas, key=lambda a: -a["cost"])
        card(s, 0.6, 1.45, 5.6, 5.35)
        cd = CategoryChartData()
        cd.categories = [a["label"] for a in ranked]
        cd.add_series("Spend", [a["cost"] for a in ranked])
        chart = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(0.8), Inches(1.65), Inches(5.2),
                                   Inches(4.95), cd).chart
        chart.has_legend = True
        chart.has_title = False
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(9.5)
        chart.font.name = FONT
        for i, point in enumerate(chart.plots[0].series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = rgb(AREA_COLOURS[i % len(AREA_COLOURS)])
            point.format.line.color.rgb = rgb(WHITE)
        hole = chart.plots[0]._element.find(qn("c:holeSize"))
        if hole is not None:
            hole.set("val", "68")
        text(s, compact(total), 2.05, 3.55, 2.7, 0.5, 24, True, INK, align="c")
        text(s, "total spend", 2.05, 4.05, 2.7, 0.3, 10.5, False, MUTED, align="c")

        card(s, 6.4, 1.45, 6.33, 5.35)
        text(s, "Ranking", 6.65, 1.62, 4, 0.3, 13, True, INK)
        for i, a in enumerate(ranked[:8]):
            yy = 2.1 + i * 0.58
            rect(s, 6.65, yy + 0.06, 0.13, 0.13, AREA_COLOURS[i % len(AREA_COLOURS)], kind=MSO_SHAPE.OVAL)
            text(s, a["label"], 6.92, yy, 3.2, 0.25, 11, False, INK)
            text(s, money(a["cost"]), 10.0, yy, 1.5, 0.25, 11, True, INK, align="r")
            text(s, f"{a['cost'] / total * 100:.1f}%" if total else "", 11.6, yy, 0.9, 0.25, 11, False, MUTED, align="r")
            rect(s, 6.92, yy + 0.3, 5.6, 0.08, "EEF2F7", radius=0.5)
            share = a["cost"] / ranked[0]["cost"] if ranked[0]["cost"] else 0
            rect(s, 6.92, yy + 0.3, max(0.05, 5.6 * share), 0.08, AREA_COLOURS[i % len(AREA_COLOURS)], radius=0.5)
        footer(s)

    # --------------------------------------------------------------- area slides
    for section in sorted(sections, key=lambda x: -(x["kpis"]["total"] or 0))[:8]:
        k = section["kpis"]
        colour = area_colour.get(section.get("id"), BLUE)
        s = prs.slides.add_slide(blank)
        page_header(s, section["label"], "Area detail", colour)
        change_text, change_colour = delta(k.get("change_pct"))
        kpi(s, 0.6, 1.45, 2.9, 1.2, "Spend", compact(k["total"]), change_text, colour, change_colour)
        kpi(s, 3.65, 1.45, 2.9, 1.2, "Previous period", compact(k.get("previous")), "comparison baseline", SLATE)
        kpi(s, 6.7, 1.45, 2.9, 1.2, "Share of total",
            f"{k['share_pct']:.1f}%" if k.get("share_pct") is not None else "—", "of in-scope spend", colour)
        kpi(s, 9.75, 1.45, 2.98, 1.2, "Billed resources", f"{k['resources']:,}",
            f"{k['services']} services", SLATE)

        card(s, 0.6, 2.85, 7.55, 3.95)
        trend = section.get("trend") or {}
        if len(trend.get("labels", [])) > 1:
            text(s, "Daily spend", 0.85, 3.02, 5, 0.3, 13, True, INK)
            if partial_note:
                text(s, partial_note, 4.4, 3.05, 3.6, 0.3, 9.5, False, MUTED, align="r")
            cd = CategoryChartData()
            cd.categories = thin_labels(trend["labels"])
            cd.add_series(section["label"], trend["values"])
            chart = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.7), Inches(3.35),
                                       Inches(7.35), Inches(3.35), cd).chart
            chart.has_legend = False
            plot = chart.plots[0]
            plot.gap_width = 45
            ser = plot.series[0]
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = rgb(colour)
            clean_axes(chart)
        else:
            text(s, "Top regions", 0.85, 3.02, 5, 0.3, 13, True, INK)
            bar_list(s, 0.85, 3.5, 7.0, [(r["name"], r["cost"]) for r in section.get("regions", [])], colour, 6)

        card(s, 8.3, 2.85, 4.43, 3.95)
        text(s, "Top services", 8.55, 3.02, 4, 0.3, 13, True, INK)
        bar_list(s, 8.55, 3.45, 3.95, [(i["name"], i["cost"]) for i in section.get("services", [])],
                 colour, rows=6, row_h=0.54)
        footer(s)

        resources = (section.get("resources_top") or [])[:8]
        if resources and len(prs.slides) < 24:
            s = prs.slides.add_slide(blank)
            page_header(s, f"{section['label']} · top resources", "Highest billed resources in this area", colour)
            style_table(
                s, ["Resource", "Resource group", "Service", "Region", f"Cost ({cur})"],
                [[r["name"], r["grp"], r["service"], r["region"], money(r["cost"])] for r in resources],
                0.6, 1.5, 12.13, [3.6, 2.6, 2.4, 1.7, 1.83], ["l", "l", "l", "l", "r"],
                row_h=0.5, size=11, accent=colour, money_cols=(4,))
            footer(s)

    # ------------------------------------------------------------ live findings
    for t in live_tables(data):
        s = prs.slides.add_slide(blank)
        page_header(s, t["title"], "   ·   ".join(t["facts"]), AMBER)
        shown = t["rows"][:11]
        ncols = len(t["columns"])
        widths = [12.13 * (2.2 if i == 0 else 1.0) / (ncols + 1.2) for i in range(ncols)]
        scale = 12.13 / sum(widths)
        widths = [w * scale for w in widths]
        rows = [[money(v) if (c + 1 == t["money_col"] and isinstance(v, (int, float)))
                 else ("—" if v is None else str(v)) for c, v in enumerate(row)] for row in shown]
        style_table(s, t["columns"], rows, 0.6, 1.5, 12.13, widths,
                    ["r" if c + 1 == t["money_col"] else "l" for c in range(ncols)],
                    row_h=0.42, size=10, accent=NAVY_2, money_cols=(t["money_col"] - 1,))
        if len(t["rows"]) > len(shown):
            text(s, f"Showing {len(shown)} of {len(t['rows'])} — the full list is in the Excel and CSV exports.",
                 0.6, 6.7, 12, 0.25, 10, False, MUTED)
        footer(s)

    # --------------------------------------------------------------- about slide
    s = prs.slides.add_slide(blank)
    background(s, NAVY)
    rect(s, 0.9, 1.0, 0.09, 1.0, CYAN)
    text(s, "ABOUT THIS REPORT", 1.25, 1.0, 8, 0.3, 11, True, CYAN, spacing=250)
    text(s, "How to read these figures", 1.25, 1.35, 10, 0.7, 32, True, WHITE)
    facts = [("Period", period_label(data)), ("Data through", data["as_of"] or "—"),
             ("Scope", scope_label), ("Currency", cur), ("Generated", data["generated"])]
    for i, (label, value) in enumerate(facts):
        yy = 2.6 + i * 0.62
        text(s, label.upper(), 1.25, yy, 2.2, 0.3, 9.5, True, "8FA3C0", spacing=100)
        text(s, value, 3.5, yy - 0.03, 8.5, 0.35, 14, False, WHITE)
    notes = ["Costs are billed amounts as recorded in Azure cost exports. They are not an invoice.",
             "Percent changes compare against the immediately preceding period of equal length."]
    if data["partial"]:
        notes.append("This is a partial export: some areas or detail blocks were not selected.")
    if data.get("mixed_currency"):
        notes.append("Multiple billing currencies are present; figures are shown in the report currency.")
    for i, note in enumerate(notes):
        text(s, "•  " + note, 1.25, 5.85 + i * 0.32, 11, 0.3, 11, False, "B9C7DC")
    text(s, "CloudLens", 11.2, 6.95, 1.6, 0.3, 12, True, CYAN, align="r")

    prs.core_properties.title = "CloudLens — Azure cost report"
    prs.core_properties.subject = period_label(data)
    prs.core_properties.author = "CloudLens"
    prs.core_properties.last_modified_by = "CloudLens"
    prs.core_properties.category = "Cost report"
    prs.core_properties.comments = f"Azure cost report. {period_label(data)}. Currency {cur}."

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
