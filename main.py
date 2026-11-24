#!/usr/bin/env python3
"""xlgit - Excelファイル(.xlsx/.xlsm)をgitのように管理するツール

内部で本物のgitを使い、Excel特有の弱点(バイナリで差分が見えない)を補います。
  * セル値・書式・条件付き書式・グラフ・画像・入力規則などの差分表示 (xlgit diff)
  * `git diff` / `git log -p` でもセル内容がテキスト表示される (textconv)
  * 過去バージョンの復元、ブランチ、マージ等はgitにそのまま委譲

依存: Python 3.8+, git, openpyxl
"""
import argparse
import io
import os
import subprocess
import sys

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("openpyxl が必要です: pip install openpyxl pillow")

EXCEL_EXT = (".xlsx", ".xlsm")
OFFICE_EXT = EXCEL_EXT + (".pptx",)
SELF = os.path.abspath(__file__)


# ------------------------------------------------------------ git helpers
def git(*args, check=True, capture=False, input_bytes=None):
    r = subprocess.run(
        ["git", *args],
        stdout=subprocess.PIPE if capture else None,
        stderr=subprocess.PIPE if capture else None,
        input=input_bytes,
    )
    if check and r.returncode != 0:
        msg = r.stderr.decode(errors="replace") if capture and r.stderr else ""
        sys.exit(f"git {' '.join(args)} が失敗しました\n{msg}")
    return r.stdout if capture else r.returncode


def repo_root():
    r = subprocess.run(["git", "rev-parse", "--show-toplevel"],
                       stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    if r.returncode != 0:
        sys.exit("gitリポジトリではありません。先に `xlgit init` を実行してください。")
    return r.stdout.decode().strip()


def git_show(rev, path):
    """revision上のファイル内容(bytes)を返す。無ければNone"""
    r = subprocess.run(["git", "show", f"{rev}:{path}"],
                       stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    return r.stdout if r.returncode == 0 else None


def rel_to_root(path):
    root = repo_root()
    return os.path.relpath(os.path.abspath(path), root).replace(os.sep, "/")


# ------------------------------------------------------------ excel parsing
import datetime
import hashlib
import re
import warnings
import zipfile
from collections import Counter

from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

WB = "(ブック全体)"
KIND_LABEL = [  # (対象IDの正規表現, ラベル)
    (r"^[A-Z]+\d+$", "書式等"), (r"^col:|^row:", "列・行"), (r"^merge:", "セル結合"),
    (r"^cf:", "条件付き書式"), (r"^dv:", "入力規則"), (r"^chart", "グラフ"),
    (r"^image@|^media:", "画像"),
    (r"^slide$", "スライド"), (r"^shape:", "図形"), (r"^deck$", "スライドサイズ"), (r"^sheet$", "シート設定"), (r"^name:", "名前定義"),
]


def safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def fmt(v):
    if isinstance(v, datetime.datetime):
        return v.date().isoformat() if v.time() == datetime.time(0) else v.isoformat(" ")
    if isinstance(v, (datetime.date, datetime.time)):
        return v.isoformat()
    if isinstance(v, str):
        return v if v.startswith("=") else repr(v)
    return repr(v)


def coord_key(coord):
    letters = "".join(ch for ch in coord if ch.isalpha())
    digits = "".join(ch for ch in coord if ch.isdigit())
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch.upper()) - 64)
    return int(digits or 0), col


def target_key(t):
    if re.match(r"^[A-Z]+\d+$", t):
        return (0, *coord_key(t), "")
    return (1, 0, 0, t)


def color_str(c):
    if c is None:
        return None
    t = getattr(c, "type", None)
    if t == "rgb":
        return str(c.rgb)
    if t == "theme":
        return f"theme{c.theme}" + (f"{c.tint:+.2f}" if c.tint else "")
    if t == "indexed":
        return f"indexed{c.indexed}"
    return None


def anchor_str(a):
    if isinstance(a, str):
        return a
    f = getattr(a, "_from", None)
    if f is None:
        return "?"
    return f"{get_column_letter(f.col + 1)}{f.row + 1}"


def title_text(t):
    if t is None or t.tx is None:
        return None
    if t.tx.rich is not None:
        return "".join(r.t or "" for p in t.tx.rich.p for r in (p.r or []))
    if t.tx.strRef is not None:
        return t.tx.strRef.f
    return None


def ref_of(x):
    if x is None:
        return None
    for attr in ("numRef", "strRef", "multiLvlStrRef"):
        r = getattr(x, attr, None)
        if r is not None:
            return r.f
    return None


def dxf_desc(dxf, put):
    if dxf is None:
        return
    if dxf.font is not None:
        f = dxf.font
        for k in ("b", "i", "strike"):
            if getattr(f, k, None):
                put(f"dxf.font.{k}", True)
        if f.color is not None:
            put("dxf.font.color", color_str(f.color))
    if dxf.fill is not None:
        fl = dxf.fill
        put("dxf.fill", f"{color_str(fl.bgColor)}/{color_str(fl.fgColor)}")
    if dxf.numFmt is not None:
        put("dxf.numfmt", dxf.numFmt.formatCode)
    if dxf.border is not None:
        for side in ("left", "right", "top", "bottom"):
            s = getattr(dxf.border, side, None)
            if s is not None and s.style:
                put(f"dxf.border.{side}", s.style)


def snapshot(data, cached=False):
    if is_pptx(data):
        return snapshot_pptx(data)
    """xlsxバイト列 -> ({(シート, 対象, 属性): 文字列}, シート順)"""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(io.BytesIO(data), data_only=False)
        wbv = load_workbook(io.BytesIO(data), data_only=True) if cached else None
    props, order = {}, []

    def put_to(sheet):
        def put(target, prop, val):
            if val is not None:
                props[(sheet, target, prop)] = str(val)
        return put

    wbput = put_to(WB)
    for n, dn in safe(lambda: list(wb.defined_names.items()), []):
        wbput(f"name:{n}", "ref", safe(lambda: dn.attr_text))

    default_font = wb._fonts[0]
    has_pil = True
    try:
        import PIL  # noqa: F401
    except ImportError:
        has_pil = False

    for ws in wb.worksheets:
        order.append(ws.title)
        put = put_to(ws.title)
        # --- シート設定
        put("sheet", "state", ws.sheet_state if ws.sheet_state != "visible" else None)
        put("sheet", "tab_color", color_str(ws.sheet_properties.tabColor))
        put("sheet", "freeze_panes", ws.freeze_panes)
        put("sheet", "autofilter", ws.auto_filter.ref)
        put("sheet", "print_area", safe(lambda: ws.print_area))
        put("sheet", "protected", True if ws.protection.sheet else None)
        for n, dn in safe(lambda: list(ws.defined_names.items()), []):
            put(f"name:{n}", "ref", safe(lambda: dn.attr_text))
        # --- セル
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c, MergedCell):
                    continue
                t = c.coordinate
                if c.value is not None and c.value != "":
                    put(t, "value", fmt(c.value))
                    if wbv is not None and isinstance(c.value, str) and c.value.startswith("="):
                        rv = wbv[ws.title][t].value
                        if rv is not None:
                            put(t, "result", fmt(rv))
                if c.comment is not None:
                    put(t, "comment", c.comment.text)
                if c.hyperlink is not None:
                    put(t, "hyperlink", c.hyperlink.target or c.hyperlink.location)
                if not c.has_style:
                    continue
                f = c.font
                d0 = default_font
                put(t, "font.name", f.name if f.name != d0.name else None)
                put(t, "font.size", f.sz if f.sz != d0.sz else None)
                for k, v in (("bold", f.b), ("italic", f.i), ("strike", f.strike)):
                    put(t, f"font.{k}", True if v else None)
                put(t, "font.underline", f.u)
                fc = color_str(f.color)
                put(t, "font.color", fc if fc != color_str(d0.color) else None)
                if c.fill is not None and c.fill.fill_type:
                    put(t, "fill", f"{c.fill.fill_type} {color_str(c.fill.fgColor)}")
                for side in ("left", "right", "top", "bottom"):
                    s = getattr(c.border, side)
                    if s is not None and s.style:
                        put(t, f"border.{side}", f"{s.style} {color_str(s.color) or ''}".strip())
                if c.number_format != "General":
                    put(t, "numfmt", c.number_format)
                a = c.alignment
                put(t, "align.h", a.horizontal if a.horizontal != "general" else None)
                put(t, "align.v", a.vertical if a.vertical != "bottom" else None)
                put(t, "align.wrap", True if a.wrap_text else None)
                put(t, "align.indent", a.indent or None)
                put(t, "align.rotation", a.text_rotation or None)
                if c.protection.locked is False:
                    put(t, "protection.locked", False)
                if c.protection.hidden:
                    put(t, "protection.hidden", True)
        # --- 結合セル
        for rng in ws.merged_cells.ranges:
            put(f"merge:{rng}", "exists", "yes")
        # --- 列幅・行高
        for key, d in ws.column_dimensions.items():
            if d.width:
                put(f"col:{key}", "width", round(d.width, 2))
            put(f"col:{key}", "hidden", True if d.hidden else None)
        for key, d in ws.row_dimensions.items():
            put(f"row:{key}", "height", round(d.height, 2) if d.height else None)
            put(f"row:{key}", "hidden", True if d.hidden else None)
        # --- 条件付き書式
        for cf in ws.conditional_formatting:
            for i, r in enumerate(cf.rules, 1):
                t = f"cf:{cf.sqref}#{i}"
                put(t, "type", r.type)
                put(t, "operator", r.operator)
                put(t, "formula", " | ".join(r.formula) if r.formula else None)
                put(t, "text", r.text)
                put(t, "rank", r.rank)
                put(t, "stop_if_true", True if r.stopIfTrue else None)
                dxf_desc(r.dxf, lambda p, v, t=t: put(t, p, v))
                if r.colorScale is not None:
                    cs = r.colorScale
                    put(t, "color_scale", ", ".join(v.type + (f"={v.val}" if v.val is not None else "") for v in cs.cfvo)
                        + " / colors " + ",".join(str(color_str(c)) for c in cs.color))
                if r.dataBar is not None:
                    put(t, "data_bar", color_str(r.dataBar.color))
                if r.iconSet is not None:
                    put(t, "icon_set", r.iconSet.iconSet)
        # --- 入力規則
        seen = Counter()
        for dv in ws.data_validations.dataValidation:
            base = f"dv:{dv.sqref}"
            seen[base] += 1
            t = base if seen[base] == 1 else f"{base}#{seen[base]}"
            put(t, "type", dv.type)
            put(t, "operator", dv.operator)
            put(t, "formula1", dv.formula1)
            put(t, "formula2", dv.formula2)
            put(t, "allow_blank", dv.allow_blank)
        # --- グラフ
        seen = Counter()
        for ch in ws._charts:
            base = f"chart@{anchor_str(ch.anchor)}"
            seen[base] += 1
            t = base if seen[base] == 1 else f"{base}#{seen[base]}"
            put(t, "type", type(ch).__name__)
            put(t, "title", title_text(ch.title))
            put(t, "grouping", safe(lambda: ch.grouping))
            put(t, "bar_dir", safe(lambda: ch.type))
            put(t, "style", safe(lambda: ch.style))
            for ax in ("x_axis", "y_axis"):
                put(t, f"{ax}.title", safe(lambda: title_text(getattr(ch, ax).title)))
            to = getattr(ch.anchor, "to", None)
            if to is not None:
                put(t, "anchor.to", f"{get_column_letter(to.col + 1)}{to.row + 1}")
            for i, s in enumerate(ch.series, 1):
                nm = safe(lambda: s.tx.strRef.f if s.tx.strRef else s.tx.v)
                put(t, f"series{i}.name", nm)
                put(t, f"series{i}.val", ref_of(getattr(s, "val", None)))
                put(t, f"series{i}.cat", ref_of(getattr(s, "cat", None)))
                put(t, f"series{i}.x", ref_of(getattr(s, "xVal", None)))
                put(t, f"series{i}.y", ref_of(getattr(s, "yVal", None)))
        # --- 画像
        if has_pil:
            seen = Counter()
            for im in ws._images:
                base = f"image@{anchor_str(im.anchor)}"
                seen[base] += 1
                t = base if seen[base] == 1 else f"{base}#{seen[base]}"
                put(t, "md5", safe(lambda: hashlib.md5(im._data()).hexdigest()[:10]))
                put(t, "size", f"{im.width}x{im.height}")
    # --- 生XMLのハッシュ(上記で拾えない細かな変更の検出用)
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        chart_h = sorted(hashlib.md5(z.read(n)).hexdigest()[:10]
                         for n in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n))
        wbput("chart-xml", "count", len(chart_h))
        wbput("chart-xml", "digest", hashlib.md5("".join(chart_h).encode()).hexdigest()[:10])
        if not has_pil:
            for n in z.namelist():
                if n.startswith("xl/media/"):
                    h = hashlib.md5(z.read(n)).hexdigest()[:10]
                    wbput(f"media:{h}", "file", f"{n.split('/')[-1]} ({z.getinfo(n).file_size}B)")
    return props, order


def dump_text(data):
    """textconv用: 全体を「対象<TAB>属性<TAB>値」形式のテキストに展開"""
    props, order = snapshot(data)
    out = []
    for name in order + [WB]:
        label = props.get((name, "slide", "_label"))
        out.append(f"=== Slide: {label} ===" if label else f"=== Sheet: {name} ===")
        items = sorted((k for k in props if k[0] == name and not k[2].startswith("_")),
                       key=lambda k: (target_key(k[1]), k[2]))
        for _, target, prop in items:
            out.append(f"{target}\t{prop}\t{props[(name, target, prop)]}")
    return "\n".join(out) + "\n"


# ------------------------------------------------------------ PowerPoint (.pptx)
EMU_PT = 12700.0
THEME_TAGS = {"TEXT_1": "dk1", "BACKGROUND_1": "lt1", "TEXT_2": "dk2", "BACKGROUND_2": "lt2",
              "DARK_1": "dk1", "LIGHT_1": "lt1", "DARK_2": "dk2", "LIGHT_2": "lt2",
              "HYPERLINK": "hlink", "FOLLOWED_HYPERLINK": "folHlink"}
CONTENT_PREFIX = ("text", "img", "cell:", "chart", "series", "kind")


def is_pptx(data):
    if not data:
        return False
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            return "ppt/presentation.xml" in z.namelist()
    except Exception:
        return False


def pptx_theme(z):
    names = sorted(n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n))
    if not names:
        return {}
    x = z.read(names[0]).decode("utf-8", "replace")
    out = {}
    for tag, rgb, sysc in re.findall(
            r'<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"'
            r'|sysClr [^>]*?lastClr="([0-9A-Fa-f]{6})")', x):
        out[tag] = (rgb or sysc).upper()
    return out


def pcolor(c, theme):
    try:
        from pptx.enum.dml import MSO_COLOR_TYPE
        if c is None or c.type is None:
            return None
        if c.type == MSO_COLOR_TYPE.RGB:
            return str(c.rgb)
        nm = getattr(c.theme_color, "name", str(c.theme_color))
        key = THEME_TAGS.get(nm) or (f"accent{nm[-1]}" if nm.startswith("ACCENT_") else None)
        return theme.get(key) or f"theme:{nm}"
    except Exception:
        return None


def walk_shapes(shapes, tf=(1.0, 1.0, 0.0, 0.0)):
    """グループを展開しつつ、スライド上の絶対座標(EMU)を返す"""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    sx, sy, dx, dy = tf
    for sh in shapes:
        l, t, w, h = (getattr(sh, a, None) or 0 for a in ("left", "top", "width", "height"))
        geo = (l * sx + dx, t * sy + dy, w * sx, h * sy)
        yield sh, geo
        if safe(lambda: sh.shape_type) == MSO_SHAPE_TYPE.GROUP:
            x = safe(lambda: sh._element.grpSpPr.xfrm)
            cx = safe(lambda: x.chExt.cx) or w or 1
            cy = safe(lambda: x.chExt.cy) or h or 1
            ox = safe(lambda: x.chOff.x) or 0
            oy = safe(lambda: x.chOff.y) or 0
            nsx, nsy = geo[2] / cx, geo[3] / cy
            yield from walk_shapes(sh.shapes, (nsx, nsy, geo[0] - ox * nsx, geo[1] - oy * nsy))


def read_text(tf, theme):
    from lxml import etree
    from pptx.text.text import _Run
    paras = []
    for p in tf.paragraphs:
        runs = []
        for el in p._p.iterchildren():
            tag = etree.QName(el).localname
            if tag == "br":
                runs.append({"t": "\n"})
            elif tag == "r":
                r = _Run(el, p)
                f = r.font
                col = pcolor(f.color, theme) if safe(lambda: f.color.type) is not None else None
                runs.append({"t": r.text, "b": bool(f.bold), "i": bool(f.italic),
                             "u": bool(f.underline), "sz": f.size.pt if f.size else None, "c": col})
        paras.append({"al": getattr(p.alignment, "name", None), "runs": runs})
    return paras


def read_shape(sh, geo, z, theme):
    from pptx.enum.dml import MSO_FILL
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    st = safe(lambda: sh.shape_type)
    d = {"id": sh.shape_id, "name": sh.name, "x": geo[0], "y": geo[1], "w": geo[2], "h": geo[3],
         "rot": round(safe(lambda: sh.rotation, 0) or 0, 1), "z": z, "kind": "shape"}
    if st == MSO_SHAPE_TYPE.GROUP:
        d["kind"] = "group"
    elif st == MSO_SHAPE_TYPE.PICTURE:
        d["kind"] = "picture"
        blob = safe(lambda: sh.image.blob)
        if blob:
            d["img_md5"] = hashlib.md5(blob).hexdigest()[:10]
            d["_blob"] = blob
            d["_ext"] = safe(lambda: sh.image.ext, "png")
    elif safe(lambda: sh.has_table):
        d["kind"] = "table"
        d["table"] = [[c.text for c in row.cells] for row in sh.table.rows]
    elif safe(lambda: sh.has_chart):
        d["kind"] = "chart"
        ch = sh.chart
        series = []
        for plot in ch.plots:
            cats = [str(c) for c in safe(lambda: list(plot.categories), [])]
            for s in plot.series:
                series.append({"name": s.name, "cats": cats, "vals": [v for v in s.values]})
        d["chart"] = {"type": getattr(ch.chart_type, "name", str(ch.chart_type)),
                      "title": safe(lambda: ch.chart_title.text_frame.text) if ch.has_title else None,
                      "series": series}
    elif safe(lambda: sh.has_text_frame):
        d["kind"] = "text"
    if safe(lambda: sh.has_text_frame) and st != MSO_SHAPE_TYPE.GROUP:
        d["paras"] = read_text(sh.text_frame, theme)
    try:
        if sh.fill.type == MSO_FILL.SOLID:
            d["fill"] = pcolor(sh.fill.fore_color, theme)
    except Exception:
        pass
    try:
        if sh.line.fill.type == MSO_FILL.SOLID:
            d["line"] = pcolor(sh.line.color, theme)
            d["lw"] = round(sh.line.width.pt, 2) if sh.line.width else None
    except Exception:
        pass
    return d


def read_pptx(data):
    from pptx import Presentation
    prs = Presentation(io.BytesIO(data))
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        theme = pptx_theme(z)
    out = {"w": prs.slide_width, "h": prs.slide_height, "slides": []}
    for pos, slide in enumerate(prs.slides, 1):
        title = (safe(lambda: slide.shapes.title.text_frame.text) or "").strip().replace("\n", " ")
        layout = safe(lambda: slide.slide_layout.name) or ""
        bg = None
        try:
            from pptx.enum.dml import MSO_FILL
            if slide.background.fill.type == MSO_FILL.SOLID:
                bg = pcolor(slide.background.fill.fore_color, theme)
        except Exception:
            pass
        notes = safe(lambda: slide.notes_slide.notes_text_frame.text) if slide.has_notes_slide else ""
        out["slides"].append({
            "key": f"slide:{slide.slide_id}", "pos": pos, "layout": layout, "bg": bg,
            "label": f"{pos}. {title or layout or 'スライド'}", "notes": notes or "",
            "hidden": slide._element.get("show") == "0",
            "shapes": [read_shape(sh, g, i, theme)
                       for i, (sh, g) in enumerate(walk_shapes(slide.shapes))]})
    return out


def flat(s):
    return str(s).replace("\r", "").replace("\v", "⏎").replace("\n", "⏎")


def snapshot_pptx(data):
    deck = read_pptx(data)
    props, order = {}, []

    def mk(key):
        def put(target, prop, val):
            if val is not None and val != "":
                props[(key, target, prop)] = str(val)
        return put

    mk(WB)("deck", "size", f"{deck['w'] / EMU_PT:.0f}x{deck['h'] / EMU_PT:.0f}pt")
    for sl in deck["slides"]:
        order.append(sl["key"])
        put = mk(sl["key"])
        put("slide", "_label", sl["label"])
        put("slide", "position", sl["pos"])
        put("slide", "layout", sl["layout"])
        put("slide", "notes", flat(sl["notes"]))
        put("slide", "hidden", True if sl["hidden"] else None)
        put("slide", "background", sl["bg"])
        for s in sl["shapes"]:
            t = f"shape:{s['id']}"
            put(t, "name", s["name"])
            put(t, "kind", s["kind"])
            put(t, "pos", f"{s['x'] / EMU_PT:.1f},{s['y'] / EMU_PT:.1f}")
            put(t, "size", f"{s['w'] / EMU_PT:.1f}x{s['h'] / EMU_PT:.1f}")
            put(t, "rotation", s["rot"] or None)
            put(t, "z", s["z"])
            put(t, "fill", s.get("fill"))
            put(t, "line", s.get("line"))
            put(t, "line_width", s.get("lw"))
            put(t, "img", s.get("img_md5"))
            if "paras" in s:
                text = "\n".join("".join(r["t"] for r in p["runs"]) for p in s["paras"])
                put(t, "text", flat(text))
                fm = set()
                for p in s["paras"]:
                    for r in p["runs"]:
                        if r["t"].strip():
                            fm.add(" ".join(x for x in (
                                f"{r['sz']:g}pt" if r.get("sz") else "", "bold" if r.get("b") else "",
                                "italic" if r.get("i") else "", "underline" if r.get("u") else "",
                                f"#{r['c']}" if r.get("c") else "") if x) or "既定")
                put(t, "fonts", ", ".join(sorted(fm)))
                put(t, "align", ", ".join(sorted({p["al"] for p in s["paras"] if p["al"]})))
            for ri, row in enumerate(s.get("table", []), 1):
                for ci, cell in enumerate(row, 1):
                    put(t, f"cell:{ri},{ci}", flat(cell))
            if "chart" in s:
                put(t, "chart.type", s["chart"]["type"])
                put(t, "chart.title", s["chart"]["title"])
                for i, se in enumerate(s["chart"]["series"], 1):
                    put(t, f"series{i}.name", se["name"])
                    put(t, f"series{i}.categories", ", ".join(se["cats"]))
                    put(t, f"series{i}.values", ", ".join(str(v) for v in se["vals"]))
    return props, order


def kind_of(target):
    for pat, label in KIND_LABEL:
        if re.search(pat, target):
            return label
    return "その他"


def describe(items, sep="; "):
    return sep.join(items)


def diff_workbooks(old, new):
    """2つのxlsxバイト列(None可)の差分(値・書式・グラフ・画像等)を行リストで返す"""
    op, o_order = snapshot(old) if old else ({}, [])
    np_, n_order = snapshot(new) if new else ({}, [])
    lines = []
    noun = "スライド" if any(k[1] == "slide" for k in list(op) + list(np_)) else "シート"
    lab = lambda key: np_.get((key, "slide", "_label")) or op.get((key, "slide", "_label")) or key
    for name in o_order:
        if name not in n_order:
            lines.append(f"- {noun}削除: {lab(name)}")
    for name in n_order:
        if name not in o_order:
            lines.append(f"+ {noun}追加: {lab(name)}")
    if [s for s in o_order if s in n_order] != [s for s in n_order if s in o_order]:
        lines.append(f"~ {noun}の順序が変更されました")

    chart_lines = 0
    for name in [s for s in n_order if s in o_order] + [WB]:
        o = {k[1:]: v for k, v in op.items() if k[0] == name}
        n = {k[1:]: v for k, v in np_.items() if k[0] == name}
        targets = sorted({t for t, _ in o} | {t for t, _ in n}, key=target_key)
        for t in targets:
            if t == "chart-xml":
                continue
            op_t = {p: v for (tt, p), v in o.items() if tt == t and not p.startswith("_")}
            np_t = {p: v for (tt, p), v in n.items() if tt == t and not p.startswith("_")}
            if op_t == np_t:
                continue
            kind = kind_of(t)
            if name == WB:
                loc = t
            else:
                nm = np_t.get("name") or op_t.get("name")
                loc = f"{lab(name)}!{t}" + (f"({nm})" if t.startswith("shape:") and nm else "")
            if kind == "グラフ":
                chart_lines += 1
            groups = [(None, ("value",)), (kind, None)] if kind == "書式等" else [(kind, None)]
            for tag, only in groups:
                keys = sorted(set(op_t) | set(np_t))
                if only is not None:
                    keys = [k for k in keys if k in only]
                elif kind == "書式等":
                    keys = [k for k in keys if k != "value"]
                if not keys:
                    continue
                if tag is None:  # セル値(従来形式)
                    ov, nv = op_t.get("value"), np_t.get("value")
                    if ov == nv:
                        continue
                    if ov is None:
                        lines.append(f"+ {loc}: {nv}")
                    elif nv is None:
                        lines.append(f"- {loc}: {ov}")
                    else:
                        lines.append(f"~ {loc}: {ov}  ->  {nv}")
                    continue
                if not op_t or all(k not in op_t for k in keys):
                    lines.append(f"+ [{tag}] {loc}  " + describe(f"{k}={np_t[k]}" for k in keys if k in np_t))
                elif not np_t or all(k not in np_t for k in keys):
                    lines.append(f"- [{tag}] {loc}  " + describe(f"{k}={op_t[k]}" for k in keys if k in op_t))
                else:
                    ch = [f"{k}: {op_t.get(k, '(既定)')} -> {np_t.get(k, '(既定)')}"
                          for k in keys if op_t.get(k) != np_t.get(k)]
                    if ch:
                        lines.append(f"~ [{tag}] {loc}  " + describe(ch))
    oc = (op.get((WB, "chart-xml", "digest")), op.get((WB, "chart-xml", "count")))
    nc = (np_.get((WB, "chart-xml", "digest")), np_.get((WB, "chart-xml", "count")))
    if oc != nc and chart_lines == 0 and old is not None and new is not None:
        lines.append(f"~ [グラフ] グラフの内部設定(色・軸・書式など)が変更されました"
                     f" (グラフ数 {oc[1] or 0} -> {nc[1] or 0})")
    return lines


def color(line):
    if not sys.stdout.isatty():
        return line
    c = {"+": "32", "-": "31", "~": "33"}.get(line[:1])
    return f"\033[{c}m{line}\033[0m" if c else line


# ------------------------------------------------------------ commands
def cmd_init(args):
    if subprocess.run(["git", "rev-parse", "--git-dir"],
                      stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode != 0:
        git("init")
    root = repo_root()
    ga = os.path.join(root, ".gitattributes")
    existing = open(ga).read() if os.path.exists(ga) else ""
    add = "".join(f"*{e} diff=office -merge -text\n" for e in OFFICE_EXT
                  if f"*{e} diff=" not in existing)
    if add:
        with open(ga, "a") as f:
            f.write(add)
    # Excelの一時ファイルを無視
    gi = os.path.join(root, ".gitignore")
    gi_text = open(gi).read() if os.path.exists(gi) else ""
    if "~$*" not in gi_text:
        with open(gi, "a") as f:
            f.write("~$*\n.~lock.*#\n")
    for drv in ("office", "xlsx"):  # xlsx は旧バージョンのリポジトリ互換
        git("config", f"diff.{drv}.textconv", f'python3 "{SELF}" textconv')
        git("config", f"diff.{drv}.cachetextconv", "false")
    print(f"初期化しました: {root}")
    print("  - .gitattributes / .gitignore を設定")
    print("  - git diff でセル内容が表示されるよう textconv を登録")
    print("次: xlgit add <file.xlsx> && xlgit commit -m 'メッセージ'")


def cmd_textconv(args):
    with open(args.file, "rb") as f:
        data = f.read()
    try:
        sys.stdout.write(dump_text(data))
    except Exception as e:  # 壊れたファイルでもgit diffを止めない
        sys.stdout.write(f"(読み取り不能: {e})\n")


def cmd_add(args):
    git("add", *args.paths)


def cmd_commit(args):
    cmd = ["commit", "-m", args.message]
    if args.all:
        cmd.insert(1, "-a")
    git(*cmd)


def cmd_status(args):
    git("status", "-sb")


def cmd_log(args):
    extra = ["--stat"] if args.stat else []
    git("log", "--oneline", "--decorate", "--graph", f"-n{args.n}", *extra,
        *(["--", *args.paths] if args.paths else []))


def cmd_diff(args):
    """セル単位差分。
    xlgit diff                 作業ツリー vs HEAD
    xlgit diff REV             作業ツリー vs REV
    xlgit diff REV1 REV2       REV1 vs REV2
    """
    root = repo_root()
    revs = args.revs
    if len(revs) > 2:
        sys.exit("リビジョンは最大2つです")
    old_rev = revs[0] if revs else "HEAD"
    new_rev = revs[1] if len(revs) == 2 else None  # Noneは作業ツリー

    if args.path:
        paths = [rel_to_root(args.path)]
    else:
        base = ["diff", "--name-only", "--no-renames", "-z"]
        names = set()
        if new_rev:
            names |= set(git(*base, old_rev, new_rev, capture=True).decode().split("\0"))
        else:
            names |= set(git(*base, old_rev, capture=True).decode().split("\0"))
        paths = sorted(n for n in names if n.lower().endswith(OFFICE_EXT))

    if not paths:
        print("Excelファイルの変更はありません")
        return
    for p in paths:
        old = git_show(old_rev, p)
        if new_rev:
            new = git_show(new_rev, p)
        else:
            full = os.path.join(root, p)
            new = open(full, "rb").read() if os.path.exists(full) else None
        print(f"--- {p}")
        lines = diff_workbooks(old, new)
        if not lines:
            print("  (セル内容に変更なし)")
        for ln in lines:
            print(color("  " + ln))


def cmd_show(args):
    rev = args.rev
    names = git("show", "--name-only", "--pretty=format:%h %an %ad%n    %s",
                "--date=short", rev, capture=True).decode()
    print(names.split("\n\n")[0] if "\n\n" in names else names.split("\n")[0])
    parent = f"{rev}^"
    paths = [n for n in git("diff-tree", "--no-commit-id", "--name-only", "-r", "-z",
                            "--root", rev, capture=True).decode().split("\0")
             if n.lower().endswith(OFFICE_EXT)]
    for p in paths:
        print(f"--- {p}")
        for ln in diff_workbooks(git_show(parent, p), git_show(rev, p)):
            print(color("  " + ln))


def cmd_restore(args):
    """指定リビジョンのファイルを作業ツリーに復元"""
    p = rel_to_root(args.path)
    data = git_show(args.rev, p)
    if data is None:
        sys.exit(f"{args.rev} に {p} は存在しません")
    dest = os.path.join(repo_root(), p)
    if args.output:
        dest = args.output
    elif os.path.exists(dest) and not args.force:
        # 既存の未コミット変更を守るため、別名での復元を促す
        sys.exit(f"{dest} は存在します。上書きは --force、別名保存は -o を指定してください。")
    with open(dest, "wb") as f:
        f.write(data)
    print(f"復元しました: {dest} ({args.rev})")


def cmd_cat(args):
    """過去版の中身をテキスト表示"""
    data = git_show(args.rev, rel_to_root(args.path))
    if data is None:
        sys.exit("見つかりません")
    sys.stdout.write(dump_text(data))


def cmd_git(args):
    sys.exit(subprocess.call(["git", *args.rest]))


# ------------------------------------------------------------ web viewer
import ast
import json
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

CELL_RE = re.compile(r"^[A-Z]+\d+$")
REV_RE = re.compile(r"^[0-9A-Za-z][0-9A-Za-z_./~^@{}-]*$")
OTHER_LINE_RE = re.compile(r"^[+\-~] (?:シート|\[(?!書式等))")


def unrepr(s):
    if s and s[0] in "'\"":
        try:
            v = ast.literal_eval(s)
            if isinstance(v, str):
                return v
        except Exception:
            pass
    return s


def build_web_book(data):
    """xlsxバイト列 -> ({シート名: {cells, merges, cols, rows}}, シート順)"""
    if data is None:
        return {}, []
    props, order = snapshot(data, cached=True)
    sheets = {n: {"cells": {}, "merges": [], "cols": {}, "rows": {}} for n in order}
    for (sheet, target, prop), val in props.items():
        sh = sheets.get(sheet)
        if sh is None:
            continue
        if CELL_RE.match(target):
            c = sh["cells"].setdefault(target, {"st": {}})
            if prop == "value":
                c["v"] = unrepr(val)
            elif prop == "result":
                c["r"] = unrepr(val)
            else:
                c["st"][prop] = val
        elif target.startswith("merge:"):
            sh["merges"].append(target[6:])
        elif target.startswith("col:") and prop == "width":
            sh["cols"][target[4:]] = float(val)
        elif target.startswith("row:") and prop == "height":
            sh["rows"][target[4:]] = float(val)
    return sheets, order


def read_version(rev, path):
    if rev == "EMPTY":
        return None
    if rev == "WORKING":
        full = os.path.join(repo_root(), path)
        return open(full, "rb").read() if os.path.isfile(full) else None
    return git_show(rev, path)


import base64


def web_shape(s):
    d = {k: v for k, v in s.items() if not k.startswith("_")}
    blob = s.get("_blob")
    mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif", "bmp": "bmp",
            "webp": "webp"}.get(str(s.get("_ext", "")).lower())
    if blob and mime and len(blob) <= 2_000_000:
        d["img"] = f"data:image/{mime};base64," + base64.b64encode(blob).decode()
    return d


def web_slide(sl, deck):
    if sl is None:
        return None
    return {"w": deck["w"], "h": deck["h"], "bg": sl["bg"],
            "shapes": [web_shape(s) for s in sl["shapes"]]}


def _short(v, n=60):
    v = str(v)
    return v if len(v) <= n else v[:n - 1] + "…"


def pptx_changes(po, pn, key):
    """スライド内の変更を [{id, s(add/del/mod/fmt), name, desc}] で返す"""
    def pick(props):
        out = {}
        for (k, t, p), v in props.items():
            if k == key and not p.startswith("_"):
                out.setdefault(t, {})[p] = v
        return out
    o, n = pick(po), pick(pn)
    order = lambda t: (t != "slide", int(t.split(":")[1]) if t.startswith("shape:") else 0)
    out = []
    for t in sorted(set(o) | set(n), key=order):
        a, b = o.get(t, {}), n.get(t, {})
        if a == b:
            continue
        sid = "slide" if t == "slide" else t.split(":")[1]
        name = "スライド設定" if t == "slide" else (b.get("name") or a.get("name") or t)
        if not a:
            s, desc = "add", "追加: " + _short(b.get("text") or b.get("kind", ""))
        elif not b:
            s, desc = "del", "削除: " + _short(a.get("text") or a.get("kind", ""))
        else:
            ps = [p for p in sorted(set(a) | set(b)) if a.get(p) != b.get(p)]
            s = "mod" if any(p.startswith(CONTENT_PREFIX) for p in ps) or t == "slide" else "fmt"
            desc = "; ".join(f"{p}: {_short(a.get(p, '(なし)'))} → {_short(b.get(p, '(なし)'))}" for p in ps)
        out.append({"id": sid, "s": s, "name": name, "desc": desc})
    return out


def web_diff_pptx(old, new, old_rev, new_rev, path):
    do = read_pptx(old) if old else {"slides": [], "w": 0, "h": 0}
    dn = read_pptx(new) if new else {"slides": [], "w": 0, "h": 0}
    po = snapshot_pptx(old)[0] if old else {}
    pn = snapshot_pptx(new)[0] if new else {}
    so = {s["key"]: s for s in do["slides"]}
    sn = {s["key"]: s for s in dn["slides"]}
    keys = [s["key"] for s in dn["slides"]] + [k for k in so if k not in sn]
    sheets = []
    for k in keys:
        a, b = so.get(k), sn.get(k)
        sheets.append({"name": k, "label": (b or a)["label"],
                       "status": "added" if a is None else "removed" if b is None else "same",
                       "old": web_slide(a, do), "new": web_slide(b, dn),
                       "changes": pptx_changes(po, pn, k)})
    others = [l for l in diff_workbooks(old, new) if re.match(r"^[+\-~] スライド", l)]
    return {"kind": "pptx", "path": path, "old": old_rev, "new": new_rev, "sheets": sheets,
            "others": others, "old_exists": old is not None, "new_exists": new is not None}


def web_diff(old_rev, new_rev, path):
    old, new = read_version(old_rev, path), read_version(new_rev, path)
    if old is None and new is None:
        raise FileNotFoundError(f"どちらのバージョンにも {path} がありません")
    if is_pptx(old) or is_pptx(new):
        return web_diff_pptx(old, new, old_rev, new_rev, path)
    os_, oo = build_web_book(old)
    ns, no = build_web_book(new)
    names = list(no) + [n for n in oo if n not in no]
    sheets = [{"name": n,
               "status": "added" if n not in os_ else "removed" if n not in ns else "same",
               "old": os_.get(n), "new": ns.get(n)} for n in names]
    others = [l for l in diff_workbooks(old, new) if OTHER_LINE_RE.match(l)]
    return {"path": path, "old": old_rev, "new": new_rev, "sheets": sheets, "others": others,
            "old_exists": old is not None, "new_exists": new is not None}


def web_log():
    r = subprocess.run(
        ["git", "log", "-n300", "--date=format:%Y-%m-%d %H:%M",
         "--pretty=format:%H%x1f%h%x1f%an%x1f%ad%x1f%s%x1f%D%x1e"],
        stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    out = []
    for rec in r.stdout.decode("utf-8", "replace").split("\x1e"):
        p = rec.strip("\n").split("\x1f")
        if len(p) == 6:
            out.append({"hash": p[0], "short": p[1], "author": p[2], "date": p[3],
                        "subject": p[4], "refs": p[5]})
    return out


def web_changed(old_rev, new_rev):
    names = set()
    def run(*a):
        r = subprocess.run(["git", *a], stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
        return r.stdout.decode("utf-8", "replace").split("\0")
    ob = "4b825dc642cb6eb9a060e54bf8d69288fbee4904" if old_rev == "EMPTY" else old_rev  # 空ツリー
    if new_rev == "WORKING":
        names |= set(run("diff", "--name-only", "--no-renames", "-z", ob))
        names |= set(run("ls-files", "-o", "--exclude-standard", "-z"))
    else:
        nb = "4b825dc642cb6eb9a060e54bf8d69288fbee4904" if new_rev == "EMPTY" else new_rev
        names |= set(run("diff", "--name-only", "--no-renames", "-z", ob, nb))
    return sorted(n for n in names if n.lower().endswith(OFFICE_EXT))


def valid_rev(r):
    return r in ("WORKING", "EMPTY") or (bool(REV_RE.match(r or "")) and ".." not in r)


def valid_path(p):
    if not p or p.startswith("/") or "\0" in p:
        return False
    parts = p.replace("\\", "/").split("/")
    return ".." not in parts and p.lower().endswith(OFFICE_EXT)


class WebHandler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def send(self, code, body, ctype="application/json; charset=utf-8"):
        if not isinstance(body, bytes):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path == "/":
                return self.send(200, INDEX_HTML.encode("utf-8"), "text/html; charset=utf-8")
            if u.path == "/api/log":
                return self.send(200, {"commits": web_log()})
            if u.path in ("/api/changed", "/api/diff"):
                o, n = q.get("old", ""), q.get("new", "")
                if not (valid_rev(o) and valid_rev(n)):
                    return self.send(400, {"error": "不正なリビジョン"})
                if u.path == "/api/changed":
                    return self.send(200, {"files": web_changed(o, n)})
                p = q.get("path", "")
                if not valid_path(p):
                    return self.send(400, {"error": "不正なパス"})
                return self.send(200, web_diff(o, n, p))
            self.send(404, {"error": "not found"})
        except Exception as e:
            self.send(500, {"error": f"{type(e).__name__}: {e}"})


def cmd_web(args):
    os.chdir(repo_root())
    srv = ThreadingHTTPServer((args.host, args.port), WebHandler)
    url = f"http://{args.host}:{args.port}/"
    print(f"xlgit web: {url}  (Ctrl+Cで終了)")
    if args.host not in ("127.0.0.1", "localhost"):
        print("警告: 認証はありません。信頼できるネットワークでのみ公開してください。")
    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n終了します")


INDEX_HTML = r"""<!doctype html>
<html lang="ja"><head><meta charset="utf-8"><title>xlgit</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
:root{--bg:#fff;--fg:#1f2328;--mut:#656d76;--bd:#d0d7de;--panel:#f6f8fa;--add:#b7ebc0;--del:#ffc9c9;--mod:#ffe98a;--fmt:#2f81f7;--head:#eceff3}
@media(prefers-color-scheme:dark){:root{--bg:#0d1117;--fg:#e6edf3;--mut:#8b949e;--bd:#30363d;--panel:#161b22;--add:#1d5c2e;--del:#6b2121;--mod:#6b5a00;--fmt:#58a6ff;--head:#1c222b}}
*{box-sizing:border-box}
body{margin:0;font:13px/1.4 system-ui,"Noto Sans JP",sans-serif;background:var(--bg);color:var(--fg);height:100vh;display:grid;grid-template-columns:270px 1fr}
#side{border-right:1px solid var(--bd);overflow:auto;background:var(--panel)}
#side h2{margin:0;padding:10px 12px;font-size:13px;border-bottom:1px solid var(--bd);position:sticky;top:0;background:var(--panel)}
.ci{padding:8px 12px;border-bottom:1px solid var(--bd);cursor:pointer}
.ci:hover{background:var(--head)}.ci.sel{background:var(--head);border-left:3px solid var(--fmt);padding-left:9px}
.ci small{color:var(--mut)}.ref{background:var(--fmt);color:#fff;border-radius:8px;padding:0 6px;font-size:10px;margin-left:4px}
#main{display:flex;flex-direction:column;min-width:0;height:100vh}
#bar{display:flex;flex-wrap:wrap;gap:8px 14px;align-items:center;padding:8px 12px;border-bottom:1px solid var(--bd)}
#bar select,#bar button{font:inherit;padding:3px 6px;background:var(--bg);color:var(--fg);border:1px solid var(--bd);border-radius:4px;max-width:260px}
.lg span{display:inline-block;padding:1px 8px;border-radius:3px;margin-right:4px;font-size:12px}
#tabs{display:flex;gap:2px;padding:6px 12px 0;border-bottom:1px solid var(--bd);overflow-x:auto}
.tab{padding:5px 12px;border:1px solid var(--bd);border-bottom:none;border-radius:5px 5px 0 0;cursor:pointer;background:var(--panel);white-space:nowrap}
.tab.on{background:var(--bg);font-weight:600}.tab .n{background:var(--mod);border-radius:8px;padding:0 6px;font-size:11px;margin-left:4px;color:var(--fg)}
.tab.added{border-top:3px solid #2da44e}.tab.removed{border-top:3px solid #cf222e;text-decoration:line-through}
#body{flex:1;display:flex;min-height:0}
#grids{flex:1;display:flex;gap:6px;padding:8px;min-width:0}
.gw{flex:1;overflow:auto;border:1px solid var(--bd);min-width:0}
.gt{position:sticky;left:0;font-size:11px;color:var(--mut);padding:2px 6px}
table.grid{border-collapse:collapse;table-layout:fixed}
.grid th{background:var(--head);color:var(--mut);font-weight:400;border:1px solid var(--bd);position:sticky;top:0;z-index:2;font-size:11px;padding:1px 4px}
.grid th.rh{left:0;z-index:3;top:auto;min-width:36px;text-align:right}
.grid thead th.rh{top:0;z-index:4}
.grid td{border:1px solid var(--bd);padding:1px 4px;height:20px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;vertical-align:top;background:var(--bg)}
.grid td small{display:block;opacity:.75;text-decoration:line-through;font-size:11px;font-weight:400}
.grid td.s-add{background:var(--add)!important}.grid td.s-del{background:var(--del)!important;text-decoration:line-through}
.grid td.s-mod{background:var(--mod)!important}.grid td.s-fmt{outline:2px solid var(--fmt);outline-offset:-2px}
.flash{animation:fl 1.6s}@keyframes fl{0%,60%{box-shadow:0 0 0 3px #f0883e inset}100%{box-shadow:none}}
#panel{width:330px;border-left:1px solid var(--bd);overflow:auto;background:var(--panel)}
#panel h3{margin:0;padding:8px 10px;font-size:12px;border-bottom:1px solid var(--bd);position:sticky;top:0;background:var(--panel)}
.li{padding:4px 10px;border-bottom:1px solid var(--bd);cursor:pointer;word-break:break-all;font-size:12px}
.li:hover{background:var(--head)}.li b{font-family:ui-monospace,monospace}.li .o{color:var(--mut);text-decoration:line-through}
.oth{cursor:default;font-family:ui-monospace,monospace;font-size:11px}
.add{color:#2da44e}.del{color:#cf222e}.mod{color:#bf8700}
#msg{padding:30px;color:var(--mut)}
.pv{padding:10px;max-width:1000px;margin:auto}.slide{position:relative;container-type:inline-size;width:100%;color:#222;box-shadow:0 1px 6px rgba(0,0,0,.3);overflow:hidden}
.sp{position:absolute;overflow:hidden;box-sizing:border-box}.sp img{width:100%;height:100%;object-fit:contain;display:block}
.sp .tx{padding:.3cqw .6cqw}.sp p{margin:0;line-height:1.2;white-space:pre-wrap}
.sp .tb{border-collapse:collapse;width:100%;font-size:1.6cqw}.sp .tb td{border:1px solid #999;padding:2px 4px;background:#fff}
.sp .ch{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;background:#f0f4f8;border:1px dashed #889;font-size:2cqw;color:#345;text-align:center}
.sp::after{content:'';position:absolute;inset:0;pointer-events:none}
.sp.s-add::after{background:rgba(46,160,67,.28);box-shadow:inset 0 0 0 3px #2da44e}
.sp.s-del::after{background:rgba(207,34,46,.28);box-shadow:inset 0 0 0 3px #cf222e}
.sp.s-mod::after{background:rgba(255,213,0,.35);box-shadow:inset 0 0 0 3px #d4a72c}
.sp.s-fmt::after{box-shadow:inset 0 0 0 3px #2f81f7}.sp.flash::after{box-shadow:inset 0 0 0 6px #f0883e}
.gw .gt+.pv{padding-top:2px}
</style></head><body>
<div id="side"><h2>コミット履歴 <button id="rf" title="再読み込み">↻</button></h2><div id="commits"></div></div>
<div id="main">
 <div id="bar">
  <label>比較元 <select id="so"></select></label>
  <label>比較先 <select id="sn"></select></label>
  <label>ファイル <select id="sf"></select></label>
  <label>表示 <select id="sm"><option value="overlay">重ねて表示</option><option value="side">左右に並べる</option></select></label>
  <label><input type="checkbox" id="oc"> 変更のある行だけ</label>
  <span class="lg"><span style="background:var(--add)">追加</span><span style="background:var(--del)">削除</span><span style="background:var(--mod)">値の変更</span><span style="outline:2px solid var(--fmt);outline-offset:-2px">書式のみ変更</span></span>
 </div>
 <div id="tabs"></div>
 <div id="body"><div id="grids"><div id="msg">読み込み中…</div></div>
  <div id="panel"><h3 id="ph">変更一覧</h3><div id="pl"></div><h3>グラフ・画像・条件付き書式ほか</h3><div id="po"></div></div></div>
</div>
<script>
const $=s=>document.querySelector(s),enc=encodeURIComponent;
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const S={commits:[],old:null,nw:null,files:[],path:null,data:null,sheet:null,mode:'overlay',only:false,first:true};
async function api(u){const r=await fetch(u);const j=await r.json();if(!r.ok)throw new Error(j.error||r.status);return j}
function colName(n){let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26)}return s}
function parse(c){const m=/^([A-Z]+)(\d+)$/.exec(c);let col=0;for(const ch of m[1])col=col*26+ch.charCodeAt(0)-64;return[parseInt(m[2]),col]}
function rgb(s){if(!s)return null;return /^[0-9A-Fa-f]{6,8}$/.test(s)?'#'+s.slice(-6):null}
const hasV=x=>x&&x.v!=null;
function propDiff(o,n){const a=(o&&o.st)||{},b=(n&&n.st)||{},out=[];for(const k of new Set([...Object.keys(a),...Object.keys(b)]))if(a[k]!==b[k])out.push(`${k}: ${a[k]??'(既定)'} → ${b[k]??'(既定)'}`);return out}
function compute(sh){
 const oc=sh.old?sh.old.cells:{},nc=sh.new?sh.new.cells:{},m=new Map();let mr=0,mc=0;
 for(const c of new Set([...Object.keys(oc),...Object.keys(nc)])){
  const[r,col]=parse(c);mr=Math.max(mr,r);mc=Math.max(mc,col);
  const o=oc[c],n=nc[c],vo=hasV(o),vn=hasV(n),pd=propDiff(o,n);let s=null;
  if(!vo&&vn)s='add';else if(vo&&!vn)s='del';else if(vo&&vn&&(o.v!==n.v||o.r!==n.r))s='mod';else if(pd.length)s='fmt';
  if(s)m.set(c,{s,o,n,pd,r,col});
 }
 for(const side of[sh.old,sh.new])if(side)for(const g of side.merges||[]){for(const p of g.split(':')){const[r,col]=parse(p);mr=Math.max(mr,r);mc=Math.max(mc,col)}}
 return{m,mr,mc,count:m.size};
}
function mergeMap(merges){const span={},skip=new Set();for(const g of merges||[]){const[a,b]=g.split(':');if(!b)continue;const[r1,c1]=parse(a),[r2,c2]=parse(b);span[a]=[r2-r1+1,c2-c1+1];for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)if(r!==r1||c!==c1)skip.add(colName(c)+r)}return{span,skip}}
function css(x,fill){const st=(x&&x.st)||{},a=[];
 if(st['font.bold'])a.push('font-weight:700');if(st['font.italic'])a.push('font-style:italic');if(st['font.underline'])a.push('text-decoration:underline');
 const fc=rgb(st['font.color']);if(fc)a.push('color:'+fc);
 if(fill&&st['fill']){const c=rgb(st['fill'].split(' ')[1]);if(c){a.push('background:'+c);if(!fc)a.push('color:#000')}}
 const h=st['align.h'];if(h&&h!=='general')a.push('text-align:'+h);
 for(const sd of['left','right','top','bottom'])if(st['border.'+sd]){const c=rgb((st['border.'+sd].split(' ')[1]))||'#555';a.push(`border-${sd}:1px solid ${c}`)}
 return a.join(';')}
const txt=x=>x?esc(x.r!==undefined?x.r:x.v):'';
const full=x=>x.v+(x.r!==undefined?`  (結果: ${x.r})`:'');
function tip(c,e,x){const L=[c];if(e){if(e.o&&hasV(e.o))L.push('旧: '+full(e.o));if(e.n&&hasV(e.n))L.push('新: '+full(e.n));e.pd.forEach(p=>L.push('書式 '+p))}else if(x&&hasV(x))L.push(full(x));return esc(L.join('\n'))}
function grid(sh,cmp,side,ids){
 const src=side==='old'?sh.old:(sh.new||sh.old),meta=src||{cells:{},merges:[],cols:{},rows:{}};
 const MAXR=1500,MAXC=200,mr=Math.min(cmp.mr,MAXR),mc=Math.min(cmp.mc,MAXC);
 const R=Math.max(mr,12),C=Math.max(mc,8);
 let rows=[];const keep=new Set();
 if(S.only){for(const e of cmp.m.values()){for(let d=-1;d<=1;d++)keep.add(e.r+d)}}
 for(let r=1;r<=R;r++)if(!S.only||keep.has(r))rows.push(r);
 const mm=S.only?{span:{},skip:new Set()}:mergeMap(meta.merges);
 let h='<table class="grid"><colgroup><col style="width:44px">';
 for(let c=1;c<=C;c++){const w=meta.cols[colName(c)];h+=`<col style="width:${w?Math.round(w*7+8):80}px">`}
 h+='</colgroup><thead><tr><th class="rh"></th>';for(let c=1;c<=C;c++)h+=`<th>${colName(c)}</th>`;h+='</tr></thead><tbody>';
 let prev=0;
 for(const r of rows){
  if(S.only&&prev&&r>prev+1)h+=`<tr><th class="rh">⋮</th><td colspan="${C}" style="text-align:center;color:var(--mut)">… ${r-prev-1} 行省略 …</td></tr>`;prev=r;
  const rh=meta.rows[r];h+=`<tr${rh?` style="height:${Math.round(rh*1.33)}px"`:''}><th class="rh">${r}</th>`;
  for(let c=1;c<=C;c++){
   const co=colName(c)+r;if(mm.skip.has(co))continue;
   const e=cmp.m.get(co),o=sh.old&&sh.old.cells[co],n=sh.new&&sh.new.cells[co];
   let cls='',body='',x;
   if(side==='old'){x=o;if(e){cls=e.s==='del'?'s-del':e.s==='mod'?'s-mod':e.s==='fmt'?'s-fmt':'';}body=txt(o)}
   else if(side==='new'){x=n;if(e){cls=e.s==='add'?'s-add':e.s==='mod'?'s-mod':e.s==='fmt'?'s-fmt':''}body=txt(n)}
   else{x=n||o;
    if(e){cls='s-'+e.s;
     if(e.s==='mod')body=`<b>${txt(n)}</b><small>${txt(o)}</small>`;else if(e.s==='del')body=txt(o);else body=txt(x)}
    else body=txt(x)}
   const sp=mm.span[co],spa=sp?` rowspan="${sp[0]}" colspan="${sp[1]}"`:'';
   h+=`<td${ids?` id="c-${co}"`:''} class="${cls}"${spa} style="${css(x,!cls||cls==='s-fmt')}" title="${tip(co,e,x)}">${body}</td>`;
  }
  h+='</tr>';
 }
 return h+'</tbody></table>';
}
function render(){
 const d=S.data;if(!d){return}
 $('#tabs').innerHTML=d.sheets.map((sh,i)=>`<div class="tab ${sh.name===S.sheet?'on':''} ${sh.status}" data-n="${esc(sh.name)}">${esc(sh.label||sh.name)}${sh._c.count?`<span class="n">${sh._c.count}</span>`:''}</div>`).join('');
 document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{S.sheet=t.dataset.n;render()});
 const sh=d.sheets.find(s=>s.name===S.sheet)||d.sheets[0];if(!sh){$('#grids').innerHTML='<div id="msg">シートがありません</div>';return}
 const cmp=sh._c;
 if(d.kind==='pptx')return pptxRender(d,sh);
 if(S.mode==='side'){
  $('#grids').innerHTML=`<div class="gw" id="g1"><div class="gt">比較元 ${esc(lab(d.old))}${sh.old?'':'(なし)'}</div>${grid(sh,cmp,'old',false)}</div><div class="gw" id="g2"><div class="gt">比較先 ${esc(lab(d.new))}${sh.new?'':'(なし)'}</div>${grid(sh,cmp,'new',true)}</div>`;
  const a=$('#g1'),b=$('#g2');let lock=false;
  const sync=(f,t)=>()=>{if(lock)return;lock=true;t.scrollTop=f.scrollTop;t.scrollLeft=f.scrollLeft;requestAnimationFrame(()=>lock=false)};
  a.onscroll=sync(a,b);b.onscroll=sync(b,a);
 }else $('#grids').innerHTML=`<div class="gw">${grid(sh,cmp,'overlay',true)}</div>`;
 const items=[...cmp.m.entries()].sort((a,b)=>a[1].r-b[1].r||a[1].col-b[1].col);
 $('#ph').textContent=`変更一覧 (${esc(sh.name)}: ${items.length}セル)`;
 $('#pl').innerHTML=items.slice(0,1000).map(([c,e])=>{
  const lab=e.s==='add'?'<span class="add">＋</span>':e.s==='del'?'<span class="del">－</span>':e.s==='mod'?'<span class="mod">～</span>':'<span style="color:var(--fmt)">書</span>';
  const v=e.s==='fmt'?esc(e.pd.join(' / ')):e.s==='mod'?`<span class="o">${txt(e.o)}</span> → ${txt(e.n)}`:e.s==='add'?txt(e.n):txt(e.o);
  return`<div class="li" data-c="${c}">${lab} <b>${c}</b> ${v}</div>`}).join('')||'<div class="li oth">このシートにセルの変更はありません</div>';
 document.querySelectorAll('#pl .li[data-c]').forEach(l=>l.onclick=()=>{const el=document.getElementById('c-'+l.dataset.c);if(el){el.scrollIntoView({block:'center',inline:'center'});el.classList.remove('flash');void el.offsetWidth;el.classList.add('flash')}});
 $('#po').innerHTML=d.others.map(l=>`<div class="li oth ${l[0]==='+'?'add':l[0]==='-'?'del':'mod'}">${esc(l)}</div>`).join('')||'<div class="li oth">変更なし</div>';
}
function runCss(r,W){const a=[];const sz=r.sz||18;a.push(`font-size:${(sz*12700/W*100).toFixed(3)}cqw`);
 if(r.b)a.push('font-weight:700');if(r.i)a.push('font-style:italic');if(r.u)a.push('text-decoration:underline');if(r.c&&/^[0-9A-Fa-f]{6}$/.test(r.c))a.push('color:#'+r.c);return a.join(';')}
function spHtml(s,W,H,cls,c){
 let st=`left:${s.x/W*100}%;top:${s.y/H*100}%;width:${s.w/W*100}%;height:${s.h/H*100}%;`;
 if(s.rot)st+=`transform:rotate(${s.rot}deg);`;if(s.fill&&/^[0-9A-Fa-f]{6}$/.test(s.fill))st+=`background:#${s.fill};`;
 if(s.line&&/^[0-9A-Fa-f]{6}$/.test(s.line))st+=`border:${((s.lw||1)*12700/W*100).toFixed(3)}cqw solid #${s.line};`;
 let inner='';
 if(s.img)inner=`<img src="${s.img}">`;
 else if(s.table)inner='<table class="tb">'+s.table.map(r=>'<tr>'+r.map(t=>`<td>${esc(t)}</td>`).join('')+'</tr>').join('')+'</table>';
 else if(s.chart)inner=`<div class="ch"><div>📊 ${esc(s.chart.type)}${s.chart.title?' : '+esc(s.chart.title):''}</div><small>${esc(s.chart.series.map(x=>x.name).join(', '))}</small></div>`;
 else if(s.kind==='picture')inner='<div class="ch">🖼 画像</div>';
 else if(s.paras)inner='<div class="tx">'+s.paras.map(p=>`<p style="text-align:${({CENTER:'center',RIGHT:'right',JUSTIFY:'justify'})[p.al]||'left'}">`+(p.runs.length?p.runs.map(r=>r.t==='\n'?'<br>':`<span style="${runCss(r,W)}">${esc(r.t)}</span>`).join(''):'&nbsp;')+'</p>').join('')+'</div>';
 const tp=esc(s.name+(c?'\n'+c.desc:''));
 return`<div id="sp-${s.id}" class="sp ${cls}" style="${st}" title="${tp}">${inner}</div>`}
function pptxRender(d,sh){
 const ch=new Map(sh.changes.map(c=>[String(c.id),c]));
 const mk=(sl,side)=>{if(!sl)return'<div id="msg">(このバージョンには存在しません)</div>';
  const list=sl.shapes.slice();
  if(side==='overlay'&&sh.new&&sh.old)for(const s of sh.old.shapes){const c=ch.get(String(s.id));if(c&&c.s==='del')list.push(s)}
  const html=list.map(s=>{const c=ch.get(String(s.id));let cls='';
   if(c){if(side==='old')cls=c.s==='add'?'':'s-'+c.s;else if(side==='new')cls=c.s==='del'?'':'s-'+c.s;else cls='s-'+c.s}
   return spHtml(s,sl.w,sl.h,cls,c)}).join('');
  return`<div class="pv"><div class="slide" style="aspect-ratio:${sl.w}/${sl.h};background:${sl.bg&&/^[0-9A-Fa-f]{6}$/.test(sl.bg)?'#'+sl.bg:'#fff'}">${html}</div></div>`};
 if(S.mode==='side')$('#grids').innerHTML=`<div class="gw"><div class="gt">比較元 ${esc(lab(d.old))}</div>${mk(sh.old,'old')}</div><div class="gw"><div class="gt">比較先 ${esc(lab(d.new))}</div>${mk(sh.new,'new')}</div>`;
 else $('#grids').innerHTML=`<div class="gw">${mk(sh.new||sh.old,'overlay')}</div>`;
 $('#ph').textContent=`変更一覧 (${sh.label||sh.name}: ${sh.changes.length}件)`;
 const icon={add:'<span class="add">＋</span>',del:'<span class="del">－</span>',mod:'<span class="mod">～</span>',fmt:'<span style="color:var(--fmt)">書</span>'};
 $('#pl').innerHTML=sh.changes.map(c=>`<div class="li" data-id="${esc(c.id)}">${icon[c.s]} <b>${esc(c.name)}</b><br><span style="color:var(--mut)">${esc(c.desc)}</span></div>`).join('')||'<div class="li oth">このスライドに変更はありません</div>';
 document.querySelectorAll('#pl .li[data-id]').forEach(l=>l.onclick=()=>{const el=document.getElementById('sp-'+l.dataset.id);if(el){el.scrollIntoView({block:'center'});el.classList.remove('flash');void el.offsetWidth;el.classList.add('flash')}});
 $('#po').innerHTML=d.others.map(l=>`<div class="li oth ${l[0]==='+'?'add':l[0]==='-'?'del':'mod'}">${esc(l)}</div>`).join('')||'<div class="li oth">スライドの増減・順序の変更なし</div>';
}
const lab=r=>r==='WORKING'?'作業ツリー':r==='EMPTY'?'(空)':(S.commits.find(c=>c.hash===r)||{short:r}).short;
function fillSelects(){
 const mk=(arr)=>arr.map(([v,t])=>`<option value="${esc(v)}">${esc(t)}</option>`).join('');
 const cs=S.commits.map(c=>[c.hash,`${c.short} ${c.subject}`.slice(0,40)]);
 $('#so').innerHTML=mk([...cs,['EMPTY','(空: 最初の状態)']]);
 $('#sn').innerHTML=mk([['WORKING','作業ツリー (未コミット)'],...cs]);
 $('#so').value=S.old;$('#sn').value=S.nw;
}
function fillCommits(){
 $('#commits').innerHTML=`<div class="ci ${S.nw==='WORKING'?'sel':''}" data-i="-1"><b>作業ツリー</b><br><small>未コミットの変更</small></div>`+S.commits.map((c,i)=>`<div class="ci ${S.nw===c.hash?'sel':''}" data-i="${i}"><b>${esc(c.subject)}</b>${c.refs?c.refs.split(', ').map(r=>`<span class="ref">${esc(r.replace('HEAD -> ',''))}</span>`).join(''):''}<br><small>${c.short} ・ ${esc(c.author)} ・ ${c.date}</small></div>`).join('');
 document.querySelectorAll('.ci').forEach(e=>e.onclick=()=>{const i=+e.dataset.i;
  if(i<0){S.old=S.commits[0]?S.commits[0].hash:'EMPTY';S.nw='WORKING'}else{S.nw=S.commits[i].hash;S.old=S.commits[i+1]?S.commits[i+1].hash:'EMPTY'}
  S.path=null;reload()});
}
async function reload(){
 fillSelects();fillCommits();
 try{
  let f=(await api(`/api/changed?old=${enc(S.old)}&new=${enc(S.nw)}`)).files;
  if(!f.length&&S.first&&S.commits.length&&S.nw==='WORKING'){ // 未コミット変更が無ければ最新コミットを表示
   S.nw=S.commits[0].hash;S.old=S.commits[1]?S.commits[1].hash:'EMPTY';S.first=false;return reload()}
  S.first=false;S.files=f;
  $('#sf').innerHTML=f.map(p=>`<option value="${esc(p)}">${esc(p)}</option>`).join('')||'<option>(Excelの変更なし)</option>';
  if(!f.length){S.data=null;$('#tabs').innerHTML='';$('#grids').innerHTML='<div id="msg">この範囲にExcelファイルの変更はありません</div>';$('#pl').innerHTML='';$('#po').innerHTML='';return}
  if(!f.includes(S.path))S.path=f[0];$('#sf').value=S.path;
  const d=await api(`/api/diff?old=${enc(S.old)}&new=${enc(S.nw)}&path=${enc(S.path)}`);
  d.sheets.forEach(sh=>sh._c=d.kind==='pptx'?{count:sh.changes.length,m:new Map(),mr:0,mc:0}:compute(sh));S.data=d;
  if(!d.sheets.some(s=>s.name===S.sheet))S.sheet=(d.sheets.find(s=>s._c.count)||d.sheets[0]||{}).name;
  render();
 }catch(e){$('#grids').innerHTML=`<div id="msg">エラー: ${esc(e.message)}</div>`}
}
$('#so').onchange=e=>{S.old=e.target.value;S.path=null;reload()};
$('#sn').onchange=e=>{S.nw=e.target.value;S.path=null;reload()};
$('#sf').onchange=e=>{S.path=e.target.value;reload()};
$('#sm').onchange=e=>{S.mode=e.target.value;render()};
$('#oc').onchange=e=>{S.only=e.target.checked;render()};
$('#rf').onclick=async()=>{S.commits=(await api('/api/log')).commits;reload()};
(async()=>{S.commits=(await api('/api/log')).commits;S.old=S.commits[0]?S.commits[0].hash:'EMPTY';S.nw='WORKING';reload()})();
</script></body></html>
"""


# ------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(
        prog="xlgit", description="Excelをgitのように管理する",
        epilog="上記以外のサブコマンドは `xlgit git <コマンド>` でgitに直接渡せます"
               "(例: xlgit git branch, xlgit git checkout -b feature)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("init", help="リポジトリを初期化").set_defaults(f=cmd_init)

    p = sub.add_parser("add", help="ステージに追加"); p.add_argument("paths", nargs="+")
    p.set_defaults(f=cmd_add)

    p = sub.add_parser("commit", help="コミット")
    p.add_argument("-m", "--message", required=True)
    p.add_argument("-a", "--all", action="store_true", help="追跡済みの変更を全て含める")
    p.set_defaults(f=cmd_commit)

    sub.add_parser("status", help="状態表示").set_defaults(f=cmd_status)

    p = sub.add_parser("log", help="履歴")
    p.add_argument("-n", type=int, default=20)
    p.add_argument("--stat", action="store_true")
    p.add_argument("paths", nargs="*")
    p.set_defaults(f=cmd_log)

    p = sub.add_parser("diff", help="セル単位の差分")
    p.add_argument("revs", nargs="*", help="REV [REV2]")
    p.add_argument("-f", "--path", help="対象ファイルを限定")
    p.set_defaults(f=cmd_diff)

    p = sub.add_parser("show", help="コミットの変更内容(セル単位)")
    p.add_argument("rev", nargs="?", default="HEAD"); p.set_defaults(f=cmd_show)

    p = sub.add_parser("restore", help="過去版を復元")
    p.add_argument("rev"); p.add_argument("path")
    p.add_argument("-o", "--output", help="別名で保存")
    p.add_argument("--force", action="store_true", help="上書きを許可")
    p.set_defaults(f=cmd_restore)

    p = sub.add_parser("cat", help="過去版のセル内容を表示")
    p.add_argument("rev"); p.add_argument("path"); p.set_defaults(f=cmd_cat)

    p = sub.add_parser("textconv", help=argparse.SUPPRESS)
    p.add_argument("file"); p.set_defaults(f=cmd_textconv)

    p = sub.add_parser("web", help="ブラウザでセル表示の差分ビューアを起動")
    p.add_argument("--port", type=int, default=8765)
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--no-browser", action="store_true")
    p.set_defaults(f=cmd_web)

    p = sub.add_parser("git", help="gitコマンドをそのまま実行")
    p.add_argument("rest", nargs=argparse.REMAINDER); p.set_defaults(f=cmd_git)

    args = ap.parse_args()
    args.f(args)


if __name__ == "__main__":
    main()