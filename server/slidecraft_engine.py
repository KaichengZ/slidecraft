#!/usr/bin/env python3
"""Lossless PPTX scene graph reader and surgical python-pptx patch writer.

Reads any .pptx into a structured, editable JSON scene graph (exact shape_id,
coordinates, typography, colors, images, tables, charts, speaker notes) and
applies edits directly back to the original .pptx file via atomic write-back
without destroying untouched slides, masters, charts, or embedded media.
"""

from __future__ import annotations

import base64
import colorsys
import html as html_mod
import io
import math
import mimetypes
import os
import posixpath
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Dict, List, Optional, Tuple

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

PRESET_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
    "blue": "0000FF", "yellow": "FFFF00", "cyan": "00FFFF", "magenta": "FF00FF",
    "gray": "808080", "grey": "808080", "dkGray": "A9A9A9", "ltGray": "D3D3D3",
    "orange": "FFA500", "purple": "800080", "navy": "000080", "teal": "008080",
    "silver": "C0C0C0", "maroon": "800000", "olive": "808000", "lime": "00FF00",
    "gold": "FFD700", "coral": "FF7F50", "indigo": "4B0082", "slateGray": "708090",
}

DEFAULT_THEME_COLORS = {
    "dk1": "111827", "lt1": "FFFFFF", "dk2": "1F2937", "lt2": "F3F4F6",
    "accent1": "2563EB", "accent2": "EA580C", "accent3": "64748B",
    "accent4": "D97706", "accent5": "0EA5E9", "accent6": "16A34A",
    "hlink": "2563EB", "folHlink": "7C3AED",
}

DEFAULT_CLR_MAP = {
    "bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hlink": "hlink", "folHlink": "folHlink",
}

CANVAS_W = 960


# ---------------------------------------------------------------------------
# Color & theme helpers
# ---------------------------------------------------------------------------

def _resolve_rel_path(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _load_rels(zf: zipfile.ZipFile, part_path: str) -> Dict[str, Tuple[str, str]]:
    rels_path = posixpath.join(posixpath.dirname(part_path), "_rels", f"{posixpath.basename(part_path)}.rels")
    if rels_path not in zf.namelist():
        return {}
    out: Dict[str, Tuple[str, str]] = {}
    root = ET.fromstring(zf.read(rels_path))
    for rel in root.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id", "")
        rtype = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if rel.attrib.get("TargetMode") == "External":
            out[rid] = (rtype, target)
        else:
            out[rid] = (rtype, _resolve_rel_path(part_path, target))
    return out


def _apply_color_mods(hex_rgb: str, elem: ET.Element) -> Tuple[str, float]:
    r = int(hex_rgb[0:2], 16) / 255.0
    g = int(hex_rgb[2:4], 16) / 255.0
    b = int(hex_rgb[4:6], 16) / 255.0
    alpha = 1.0
    for child in elem:
        tag = child.tag.split("}")[-1]
        try:
            val = int(child.attrib.get("val", "0")) / 100000.0
        except ValueError:
            continue
        if tag == "alpha":
            alpha = max(0.0, min(1.0, val))
        elif tag == "tint":
            r = r * val + (1.0 - val)
            g = g * val + (1.0 - val)
            b = b * val + (1.0 - val)
        elif tag == "shade":
            r *= val
            g *= val
            b *= val
        elif tag in ("lumMod", "lumOff", "satMod", "satOff"):
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            if tag == "lumMod":
                l = max(0.0, min(1.0, l * val))
            elif tag == "lumOff":
                l = max(0.0, min(1.0, l + val))
            elif tag == "satMod":
                s = max(0.0, min(1.0, s * val))
            elif tag == "satOff":
                s = max(0.0, min(1.0, s + val))
            r, g, b = colorsys.hls_to_rgb(h, l, s)
    rr = max(0, min(255, int(round(r * 255))))
    gg = max(0, min(255, int(round(g * 255))))
    bb = max(0, min(255, int(round(b * 255))))
    return f"#{rr:02X}{gg:02X}{bb:02X}", alpha


def _resolve_color_hex_alpha(
    parent: Optional[ET.Element],
    theme_colors: Dict[str, str],
    clr_map: Dict[str, str],
) -> Tuple[Optional[str], float]:
    if parent is None:
        return None, 1.0
    for child in parent:
        tag = child.tag.split("}")[-1]
        base_hex: Optional[str] = None
        if tag == "srgbClr":
            base_hex = child.attrib.get("val")
        elif tag == "schemeClr":
            key = child.attrib.get("val", "")
            mapped = clr_map.get(key, key)
            base_hex = theme_colors.get(mapped) or theme_colors.get(key)
        elif tag == "sysClr":
            base_hex = child.attrib.get("lastClr") or "000000"
        elif tag == "prstClr":
            base_hex = PRESET_COLORS.get(child.attrib.get("val", ""))
        if base_hex and len(base_hex) == 6:
            return _apply_color_mods(base_hex, child)
    return None, 1.0


def _color_css(hex_str: Optional[str], alpha: float) -> Optional[str]:
    if not hex_str:
        return None
    if alpha >= 0.995:
        return hex_str
    r = int(hex_str[1:3], 16)
    g = int(hex_str[3:5], 16)
    b = int(hex_str[5:7], 16)
    return f"rgba({r},{g},{b},{alpha:.3f})"


def _parse_theme(zf: zipfile.ZipFile) -> Tuple[Dict[str, str], str, str]:
    colors = dict(DEFAULT_THEME_COLORS)
    major = "PingFang SC"
    minor = "PingFang SC"
    themes = [n for n in zf.namelist() if re.match(r"^ppt/theme/theme\d+\.xml$", n)]
    if not themes:
        return colors, major, minor
    root = ET.fromstring(zf.read(sorted(themes)[0]))
    cs = root.find(".//a:clrScheme", NS)
    if cs is not None:
        for ch in cs:
            key = ch.tag.split("}")[-1]
            srgb = ch.find("a:srgbClr", NS)
            sysc = ch.find("a:sysClr", NS)
            if srgb is not None and "val" in srgb.attrib:
                colors[key] = srgb.attrib["val"]
            elif sysc is not None and "lastClr" in sysc.attrib:
                colors[key] = sysc.attrib["lastClr"]
    mj = root.find(".//a:fontScheme/a:majorFont/a:latin", NS)
    mn = root.find(".//a:fontScheme/a:minorFont/a:latin", NS)
    if mj is not None and mj.attrib.get("typeface"):
        major = mj.attrib["typeface"]
    if mn is not None and mn.attrib.get("typeface"):
        minor = mn.attrib["typeface"]
    return colors, major, minor


def _encode_image_bytes(data: bytes, mime: str, max_dim: int = 1400) -> str:
    if max_dim > 0 and mime in ("image/png", "image/jpeg", "image/webp", "image/bmp", "image/tiff") and len(data) > 35_000:
        try:
            with Image.open(io.BytesIO(data)) as im:
                w, h = im.size
                if max(w, h) > max_dim or (max_dim <= 360 and len(data) > 20_000):
                    if max(w, h) > max_dim:
                        im.thumbnail((max_dim, max_dim), Image.LANCZOS)
                    buf = io.BytesIO()
                    if max_dim <= 360:
                        if im.mode in ("RGBA", "LA", "P"):
                            bg = Image.new("RGB", im.size, (255, 255, 255))
                            rgba = im.convert("RGBA")
                            bg.paste(rgba, mask=rgba.split()[-1])
                            bg.save(buf, format="JPEG", quality=72, optimize=True)
                        else:
                            im.convert("RGB").save(buf, format="JPEG", quality=72, optimize=True)
                        data = buf.getvalue()
                        mime = "image/jpeg"
                    else:
                        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                            im.save(buf, format="PNG", optimize=True)
                            mime = "image/png"
                        else:
                            im.convert("RGB").save(buf, format="JPEG", quality=88, optimize=True)
                            mime = "image/jpeg"
                        data = buf.getvalue()
        except Exception:
            pass
    return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"


def _read_media_uri(zf: zipfile.ZipFile, media_path: str, cache: Dict[str, str], max_dim: int = 1400) -> Optional[str]:
    key = f"{media_path}@{max_dim}"
    if key in cache:
        return cache[key]
    if media_path not in zf.namelist():
        return None
    mime, _ = mimetypes.guess_type(media_path)
    if not mime:
        mime = "image/svg+xml" if media_path.lower().endswith(".svg") else "image/png"
    uri = _encode_image_bytes(zf.read(media_path), mime, max_dim=max_dim)
    cache[key] = uri
    return uri


def _parse_fill(
    elem: Optional[ET.Element],
    theme_colors: Dict[str, str],
    clr_map: Dict[str, str],
    zf: zipfile.ZipFile,
    rels: Dict[str, Tuple[str, str]],
    media_cache: Dict[str, str],
    max_dim: int = 1400,
) -> Tuple[Optional[str], Optional[str]]:
    """Return (css_fill, solid_hex_or_none)."""
    if elem is None:
        return None, None
    if elem.find("a:noFill", NS) is not None:
        return "transparent", None
    solid = elem.find("a:solidFill", NS)
    if solid is not None:
        hx, alpha = _resolve_color_hex_alpha(solid, theme_colors, clr_map)
        return _color_css(hx, alpha), hx
    grad = elem.find("a:gradFill", NS)
    if grad is not None:
        stops = []
        gs_lst = grad.find("a:gsLst", NS)
        if gs_lst is not None:
            for gs in gs_lst.findall("a:gs", NS):
                pos = int(gs.attrib.get("pos", "0")) / 1000.0
                hx, alpha = _resolve_color_hex_alpha(gs, theme_colors, clr_map)
                stops.append((pos, _color_css(hx, alpha) or "#FFFFFF", hx))
        stops.sort(key=lambda x: x[0])
        if stops:
            lin = grad.find("a:lin", NS)
            ang = 135.0
            if lin is not None and "ang" in lin.attrib:
                ang = (int(lin.attrib["ang"]) / 60000.0 + 90.0) % 360.0
            s_str = ", ".join(f"{c} {p:.1f}%" for p, c, _ in stops)
            return f"linear-gradient({ang:.1f}deg, {s_str})", stops[0][2]
    bf = elem.find("a:blipFill", NS)
    if bf is not None:
        blip = bf.find("a:blip", NS)
        if blip is not None:
            rid = blip.attrib.get(f"{{{NS['r']}}}embed")
            if rid and rid in rels:
                uri = _read_media_uri(zf, rels[rid][1], media_cache, max_dim=max_dim)
                if uri:
                    return f"url('{uri}') center/cover no-repeat", None
    return None, None


def _extract_xfrm(xfrm: Optional[ET.Element]) -> Optional[Dict[str, float]]:
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    try:
        return {
            "x": float(off.attrib.get("x", "0")),
            "y": float(off.attrib.get("y", "0")),
            "w": float(ext.attrib.get("cx", "0")),
            "h": float(ext.attrib.get("cy", "0")),
            "rot": float(xfrm.attrib.get("rot", "0")) / 60000.0,
        }
    except ValueError:
        return None


def _ph_key(sp: ET.Element) -> Optional[Tuple[str, str]]:
    ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is None:
        return None
    return (ph.attrib.get("type", "body"), ph.attrib.get("idx", ""))


def _collect_placeholders(root: Optional[ET.Element]) -> Dict[Tuple[str, str], Dict[str, Any]]:
    out: Dict[Tuple[str, str], Dict[str, Any]] = {}
    if root is None:
        return out
    for sp in root.findall(".//p:spTree/p:sp", NS):
        k = _ph_key(sp)
        if not k:
            continue
        xf = _extract_xfrm(sp.find("p:spPr/a:xfrm", NS))
        if xf:
            out[k] = {"xfrm": xf}
            if k[1] != "" and (k[0], "") not in out:
                out[(k[0], "")] = {"xfrm": xf}
    return out


def _lookup_ph_xfrm(
    k: Optional[Tuple[str, str]],
    l_phs: Dict[Tuple[str, str], Dict[str, Any]],
    m_phs: Dict[Tuple[str, str], Dict[str, Any]],
) -> Optional[Dict[str, float]]:
    if not k:
        return None
    for st in (l_phs, m_phs):
        if k in st:
            return st[k]["xfrm"]
        if (k[0], "") in st:
            return st[(k[0], "")]["xfrm"]
    alias = {"ctrTitle": "title", "subTitle": "body", "obj": "body"}.get(k[0])
    if alias:
        for st in (l_phs, m_phs):
            if (alias, "") in st:
                return st[(alias, "")]["xfrm"]
    return None


def _default_pt(ph_type: Optional[str], lvl: int = 0) -> float:
    if ph_type in ("title", "ctrTitle"):
        return 34.0
    if ph_type == "subTitle":
        return 20.0
    if ph_type in ("ftr", "dt", "sldNum"):
        return 11.0
    return max(12.0, 18.0 - lvl * 2.0)


def _render_chart_svg(
    zf: zipfile.ZipFile,
    chart_path: str,
    w_px: float,
    h_px: float,
    theme_colors: Dict[str, str],
) -> str:
    if chart_path not in zf.namelist():
        return ""
    try:
        root = ET.fromstring(zf.read(chart_path))
    except Exception:
        return ""
    title = "".join(t.text for t in root.findall(".//c:title//a:t", NS) if t.text).strip()
    palette = [f"#{theme_colors.get(f'accent{i}', '2563EB')}" for i in range(1, 7)]
    plot_area = root.find(".//c:plotArea", NS)
    if plot_area is None:
        return ""
    chart_node = None
    kind = "bar"
    for k, tag in [("bar", "barChart"), ("bar", "bar3DChart"), ("line", "lineChart"), ("area", "areaChart"), ("pie", "pieChart"), ("doughnut", "doughnutChart")]:
        found = plot_area.find(f"c:{tag}", NS)
        if found is not None:
            chart_node = found
            kind = k
            break
    if chart_node is None:
        return ""

    series_list = []
    categories: List[str] = []
    for idx, ser in enumerate(chart_node.findall("c:ser", NS)):
        sname_el = ser.find(".//c:tx//c:v", NS)
        sname = sname_el.text.strip() if sname_el is not None and sname_el.text else f"S{idx + 1}"
        if not categories:
            categories = [
                (pt.find("c:v", NS).text.strip() if pt.find("c:v", NS) is not None and pt.find("c:v", NS).text else "")
                for pt in ser.findall(".//c:cat//c:pt", NS)
            ]
        vals = []
        for pt in ser.findall(".//c:val//c:pt", NS):
            v = pt.find("c:v", NS)
            try:
                vals.append(float(v.text))
            except Exception:
                vals.append(0.0)
        if vals:
            series_list.append({"name": sname, "values": vals, "color": palette[idx % len(palette)]})
    if not series_list:
        return ""
    if not categories:
        categories = [str(i + 1) for i in range(len(series_list[0]["values"]))]

    W, H = max(180.0, w_px), max(120.0, h_px)
    top_p, bot_p, left_p, right_p = (32.0 if title else 14.0), 34.0, 42.0, 16.0
    pw, ph = max(60.0, W - left_p - right_p), max(50.0, H - top_p - bot_p)
    svg = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W:.0f} {H:.0f}" width="100%" height="100%" style="display:block;">']
    if title:
        svg.append(f'<text x="{W/2:.1f}" y="21" text-anchor="middle" font-size="13" font-weight="700" fill="#1F2937">{html_mod.escape(title)}</text>')
    if kind in ("pie", "doughnut"):
        vals = series_list[0]["values"]
        tot = sum(max(0.0, v) for v in vals) or 1.0
        cx, cy, r = W * 0.42, top_p + ph / 2.0, min(pw * 0.38, ph * 0.44)
        ang = -math.pi / 2.0
        for i, v in enumerate(vals):
            frac = max(0.0, v) / tot
            end = ang + frac * 2.0 * math.pi
            x1, y1 = cx + r * math.cos(ang), cy + r * math.sin(ang)
            x2, y2 = cx + r * math.cos(end), cy + r * math.sin(end)
            svg.append(f'<path d="M {cx:.1f} {cy:.1f} L {x1:.1f} {y1:.1f} A {r:.1f} {r:.1f} 0 {1 if frac > 0.5 else 0} 1 {x2:.1f} {y2:.1f} Z" fill="{palette[i % len(palette)]}" stroke="#FFF" stroke-width="1.5"/>')
            ang = end
        if kind == "doughnut":
            svg.append(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r*0.5:.1f}" fill="#FFF"/>')
    else:
        max_v = max((max(s["values"]) for s in series_list), default=1.0) or 1.0
        for g in range(5):
            gy = top_p + ph * (1.0 - g / 4.0)
            svg.append(f'<line x1="{left_p:.1f}" y1="{gy:.1f}" x2="{W-right_p:.1f}" y2="{gy:.1f}" stroke="#E5E7EB" stroke-width="1"/>')
            gv = max_v * g / 4.0
            svg.append(f'<text x="{left_p-5:.1f}" y="{gy+3:.1f}" text-anchor="end" font-size="10" fill="#6B7280">{gv:.0f}</text>')
        gw = pw / max(1, len(categories))
        if kind in ("line", "area"):
            for s in series_list:
                pts = [(left_p + (i + 0.5) * gw, top_p + ph * (1.0 - max(0.0, v) / max_v)) for i, v in enumerate(s["values"])]
                d = " ".join(f"{'M' if i == 0 else 'L'} {px:.1f} {py:.1f}" for i, (px, py) in enumerate(pts))
                svg.append(f'<path d="{d}" fill="none" stroke="{s["color"]}" stroke-width="2.4"/>')
                for px, py in pts:
                    svg.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="3.2" fill="{s["color"]}" stroke="#FFF"/>')
        else:
            n_s = len(series_list)
            bw = max(4.0, (gw * 0.7) / max(1, n_s) - 2.0)
            for si, s in enumerate(series_list):
                for i, v in enumerate(s["values"]):
                    bx = left_p + i * gw + gw * 0.15 + si * (bw + 2.0)
                    bh = ph * (max(0.0, v) / max_v)
                    svg.append(f'<rect x="{bx:.1f}" y="{top_p+ph-bh:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="3" fill="{s["color"]}"/>')
        for i, cat in enumerate(categories):
            svg.append(f'<text x="{left_p + (i + 0.5) * gw:.1f}" y="{top_p+ph+14:.1f}" text-anchor="middle" font-size="10" fill="#4B5563">{html_mod.escape(cat[:12])}</text>')
    svg.append("</svg>")
    return "".join(svg)


def _extract_paragraphs(
    txbody: Optional[ET.Element],
    theme_colors: Dict[str, str],
    clr_map: Dict[str, str],
    major_font: str,
    minor_font: str,
    ph_type: Optional[str],
    default_color: str,
) -> Tuple[List[Dict[str, Any]], str, Dict[str, float]]:
    """Return (paragraphs_list, v_align, padding_px_dict)."""
    if txbody is None:
        return [], "top", {"l": 6.0, "r": 6.0, "t": 4.0, "b": 4.0}
    bpr = txbody.find("a:bodyPr", NS)
    v_align = "top"
    pad = {"l": 6.0, "r": 6.0, "t": 4.0, "b": 4.0}
    if bpr is not None:
        v_align = {"t": "top", "ctr": "middle", "b": "bottom"}.get(bpr.attrib.get("anchor", "t"), "top")
        try:
            pad["l"] = round(float(bpr.attrib.get("lIns", "91440")) / 12700.0, 1)
            pad["r"] = round(float(bpr.attrib.get("rIns", "91440")) / 12700.0, 1)
            pad["t"] = round(float(bpr.attrib.get("tIns", "45720")) / 12700.0, 1)
            pad["b"] = round(float(bpr.attrib.get("bIns", "45720")) / 12700.0, 1)
        except ValueError:
            pass

    paras: List[Dict[str, Any]] = []
    for p in txbody.findall("a:p", NS):
        ppr = p.find("a:pPr", NS)
        lvl = int(ppr.attrib.get("lvl", "0")) if ppr is not None else 0
        algn_raw = ppr.attrib.get("algn", "l") if ppr is not None else "l"
        algn = {"l": "left", "ctr": "center", "r": "right", "just": "justify"}.get(algn_raw, "left")
        if ph_type == "ctrTitle" and (ppr is None or "algn" not in ppr.attrib):
            algn = "center"
        bullet = ""
        if ppr is not None and ppr.find("a:buNone", NS) is None:
            bc = ppr.find("a:buChar", NS)
            ba = ppr.find("a:buAutoNum", NS)
            if bc is not None:
                bullet = bc.attrib.get("char", "•")
            elif ba is not None:
                bullet = "•"

        def_rpr = ppr.find("a:defRPr", NS) if ppr is not None else None
        runs: List[Dict[str, Any]] = []
        full_text_parts: List[str] = []
        for node in p:
            tag = node.tag.split("}")[-1]
            if tag == "br":
                runs.append({"text": "\n", "br": True})
                full_text_parts.append("\n")
                continue
            if tag not in ("r", "fld"):
                continue
            t_el = node.find("a:t", NS)
            if t_el is None or t_el.text is None:
                continue
            txt = t_el.text
            full_text_parts.append(txt)
            rpr = node.find("a:rPr", NS)
            sz_pt = _default_pt(ph_type, lvl)
            if def_rpr is not None and "sz" in def_rpr.attrib:
                sz_pt = int(def_rpr.attrib["sz"]) / 100.0
            if rpr is not None and "sz" in rpr.attrib:
                sz_pt = int(rpr.attrib["sz"]) / 100.0

            bold = (rpr is not None and rpr.attrib.get("b") in ("1", "true")) or (
                rpr is None and def_rpr is not None and def_rpr.attrib.get("b") in ("1", "true")
            )
            if ph_type in ("title", "ctrTitle") and (rpr is None or "b" not in rpr.attrib):
                bold = True
            italic = rpr is not None and rpr.attrib.get("i") in ("1", "true")
            underline = rpr is not None and rpr.attrib.get("u") not in (None, "none")

            hx = None
            if rpr is not None:
                hx, _ = _resolve_color_hex_alpha(rpr.find("a:solidFill", NS), theme_colors, clr_map)
            if not hx and def_rpr is not None:
                hx, _ = _resolve_color_hex_alpha(def_rpr.find("a:solidFill", NS), theme_colors, clr_map)
            if not hx:
                hx = default_color

            font_name = major_font if ph_type in ("title", "ctrTitle") else minor_font
            if rpr is not None:
                lat = rpr.find("a:latin", NS)
                ea = rpr.find("a:ea", NS)
                if lat is not None and lat.attrib.get("typeface") and not lat.attrib["typeface"].startswith("+"):
                    font_name = lat.attrib["typeface"]
                elif ea is not None and ea.attrib.get("typeface") and not ea.attrib["typeface"].startswith("+"):
                    font_name = ea.attrib["typeface"]

            runs.append({
                "text": txt,
                "font_size_pt": round(sz_pt, 1),
                "bold": bold,
                "italic": italic,
                "underline": underline,
                "color": hx,
                "font": font_name,
            })

        first_r = next((r for r in runs if not r.get("br")), None)
        paras.append({
            "text": "".join(full_text_parts),
            "align": algn,
            "level": lvl,
            "bullet": bullet,
            "font_size_pt": first_r["font_size_pt"] if first_r else _default_pt(ph_type, lvl),
            "bold": first_r["bold"] if first_r else False,
            "italic": first_r["italic"] if first_r else False,
            "underline": first_r["underline"] if first_r else False,
            "color": first_r["color"] if first_r else default_color,
            "font": first_r["font"] if first_r else minor_font,
            "runs": runs,
        })
    return paras, v_align, pad


def _extract_slide_parts(zf: zipfile.ZipFile) -> Tuple[float, float, List[str]]:
    pres_root = ET.fromstring(zf.read("ppt/presentation.xml"))
    pres_rels = _load_rels(zf, "ppt/presentation.xml")
    sz = pres_root.find("p:sldSz", NS)
    emu_w = float(sz.attrib.get("cx", "12192000")) if sz is not None else 12192000.0
    emu_h = float(sz.attrib.get("cy", "6858000")) if sz is not None else 6858000.0
    slide_parts: List[str] = []
    sld_lst = pres_root.find("p:sldIdLst", NS)
    if sld_lst is not None:
        for sid in sld_lst.findall("p:sldId", NS):
            rid = sid.attrib.get(f"{{{NS['r']}}}id")
            if rid and rid in pres_rels:
                slide_parts.append(pres_rels[rid][1])
    if not slide_parts:
        slide_parts = sorted(
            [n for n in zf.namelist() if re.match(r"^ppt/slides/slide\d+\.xml$", n)],
            key=lambda s: int(re.search(r"(\d+)\.xml$", s).group(1)),
        )
    return emu_w, emu_h, slide_parts


def _extract_notes_and_title(
    zf: zipfile.ZipFile, s_root: ET.Element, s_rels: Dict[str, Tuple[str, str]]
) -> Tuple[str, str]:
    title = ""
    fallback = ""
    for sp in s_root.findall(".//p:sp", NS):
        txt = " ".join(t.text.strip() for t in sp.findall(".//a:t", NS) if t.text and t.text.strip())
        if not txt:
            continue
        if not fallback:
            fallback = txt
        ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is not None and ph.attrib.get("type") in ("title", "ctrTitle") and not title:
            title = txt
    notes = ""
    for _rid, (rtype, target) in s_rels.items():
        if "notesSlide" in rtype and target in zf.namelist():
            nr = ET.fromstring(zf.read(target))
            n_lines = []
            for sp in nr.findall(".//p:sp", NS):
                ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
                if ph is not None and ph.attrib.get("type") == "sldNum":
                    continue
                for p in sp.findall(".//a:p", NS):
                    ln = "".join(t.text for t in p.findall(".//a:t", NS) if t.text).strip()
                    if ln:
                        n_lines.append(ln)
            notes = "\n".join(n_lines)
            break
    return (title or fallback)[:100], notes


def _parse_elements_from_tree(
    sp_tree: ET.Element,
    scale_x: float,
    scale_y: float,
    theme_colors: Dict[str, str],
    clr_map: Dict[str, str],
    major_font: str,
    minor_font: str,
    zf: zipfile.ZipFile,
    rels: Dict[str, Tuple[str, str]],
    media_cache: Dict[str, str],
    l_phs: Dict[Tuple[str, str], Dict[str, Any]],
    m_phs: Dict[Tuple[str, str], Dict[str, Any]],
    skip_placeholders: bool = False,
    max_img_dim: int = 1400,
) -> List[Dict[str, Any]]:
    elements: List[Dict[str, Any]] = []
    for node in sp_tree:
        tag = node.tag.split("}")[-1]
        if tag == "sp":
            ph_k = _ph_key(node)
            if skip_placeholders and ph_k is not None:
                continue
            cnv = node.find("p:nvSpPr/p:cNvPr", NS)
            sid = int(cnv.attrib.get("id", "0")) if cnv is not None else 0
            sname = cnv.attrib.get("name", f"Shape {sid}") if cnv is not None else f"Shape {sid}"
            spPr = node.find("p:spPr", NS)
            xf = _extract_xfrm(spPr.find("a:xfrm", NS) if spPr is not None else None)
            if xf is None and ph_k is not None:
                xf = _lookup_ph_xfrm(ph_k, l_phs, m_phs)
            if xf is None or (xf["w"] <= 0 and xf["h"] <= 0):
                continue

            prst_el = spPr.find("a:prstGeom", NS) if spPr is not None else None
            prst = prst_el.attrib.get("prst", "rect") if prst_el is not None else "rect"
            fill_css, fill_hex = _parse_fill(spPr, theme_colors, clr_map, zf, rels, media_cache, max_dim=max_img_dim)
            if fill_css is None:
                fref = node.find("p:style/a:fillRef", NS)
                if fref is not None and fref.attrib.get("idx", "0") != "0":
                    hx, alpha = _resolve_color_hex_alpha(fref, theme_colors, clr_map)
                    fill_css, fill_hex = _color_css(hx, alpha), hx

            stroke_css, stroke_w = None, 0.0
            ln = spPr.find("a:ln", NS) if spPr is not None else None
            if ln is not None and ln.find("a:noFill", NS) is None:
                shx, sa = _resolve_color_hex_alpha(ln.find("a:solidFill", NS), theme_colors, clr_map)
                stroke_css = _color_css(shx, sa)
                stroke_w = max(1.0, float(ln.attrib.get("w", "12700")) * scale_x)

            def_txt = f"#{theme_colors.get('dk1', '111827')}"
            if fill_hex and len(fill_hex) == 7:
                rr, gg, bb = int(fill_hex[1:3], 16), int(fill_hex[3:5], 16), int(fill_hex[5:7], 16)
                if (0.299 * rr + 0.587 * gg + 0.114 * bb) < 128:
                    def_txt = "#FFFFFF"

            txb = node.find("p:txBody", NS)
            paras, v_align, pad = _extract_paragraphs(
                txb, theme_colors, clr_map, major_font, minor_font, ph_k[0] if ph_k else None, def_txt
            )
            has_text = any(p["text"].strip() for p in paras)
            kind = "text" if (txb is not None and (has_text or not fill_css or fill_css == "transparent")) else "shape"

            elements.append({
                "shape_id": sid,
                "name": sname,
                "kind": kind,
                "editable": True,
                "x": round(xf["x"] * scale_x, 2),
                "y": round(xf["y"] * scale_y, 2),
                "w": round(max(4.0, xf["w"] * scale_x), 2),
                "h": round(max(4.0, xf["h"] * scale_y), 2),
                "rot": round(xf["rot"], 1),
                "prst": prst,
                "fill": fill_css,
                "fill_hex": fill_hex,
                "stroke": stroke_css,
                "stroke_w": round(stroke_w, 1),
                "v_align": v_align,
                "pad": pad,
                "paragraphs": paras,
            })

        elif tag == "pic":
            cnv = node.find("p:nvPicPr/p:cNvPr", NS)
            sid = int(cnv.attrib.get("id", "0")) if cnv is not None else 0
            sname = cnv.attrib.get("name", f"Image {sid}") if cnv is not None else f"Image {sid}"
            spPr = node.find("p:spPr", NS)
            xf = _extract_xfrm(spPr.find("a:xfrm", NS) if spPr is not None else None)
            blip = node.find("p:blipFill/a:blip", NS)
            if xf is None or blip is None:
                continue
            rid = blip.attrib.get(f"{{{NS['r']}}}embed")
            if not rid or rid not in rels:
                continue
            uri = _read_media_uri(zf, rels[rid][1], media_cache, max_dim=max_img_dim)
            if not uri:
                continue
            elements.append({
                "shape_id": sid,
                "name": sname,
                "kind": "image",
                "editable": True,
                "x": round(xf["x"] * scale_x, 2),
                "y": round(xf["y"] * scale_y, 2),
                "w": round(max(4.0, xf["w"] * scale_x), 2),
                "h": round(max(4.0, xf["h"] * scale_y), 2),
                "rot": round(xf["rot"], 1),
                "data_uri": uri,
            })

        elif tag == "cxnSp":
            cnv = node.find("p:nvCxnSpPr/p:cNvPr", NS)
            sid = int(cnv.attrib.get("id", "0")) if cnv is not None else 0
            spPr = node.find("p:spPr", NS)
            xf = _extract_xfrm(spPr.find("a:xfrm", NS) if spPr is not None else None)
            if xf is None:
                continue
            ln = spPr.find("a:ln", NS) if spPr is not None else None
            shx, sa = _resolve_color_hex_alpha(ln.find("a:solidFill", NS), theme_colors, clr_map) if ln is not None else (None, 1.0)
            elements.append({
                "shape_id": sid,
                "name": f"Line {sid}",
                "kind": "line",
                "editable": True,
                "x": round(xf["x"] * scale_x, 2),
                "y": round(xf["y"] * scale_y, 2),
                "w": round(max(2.0, xf["w"] * scale_x), 2),
                "h": round(max(2.0, xf["h"] * scale_y), 2),
                "rot": round(xf["rot"], 1),
                "stroke": _color_css(shx, sa) or f"#{theme_colors.get('accent1', '2563EB')}",
                "stroke_w": round(max(1.5, float(ln.attrib.get("w", "12700")) * scale_x) if ln is not None else 1.5, 1),
            })

        elif tag == "graphicFrame":
            cnv = node.find("p:nvGraphicFramePr/p:cNvPr", NS)
            sid = int(cnv.attrib.get("id", "0")) if cnv is not None else 0
            sname = cnv.attrib.get("name", f"Graphic {sid}") if cnv is not None else f"Graphic {sid}"
            xf = _extract_xfrm(node.find("p:xfrm", NS))
            if xf is None:
                continue
            x, y = round(xf["x"] * scale_x, 2), round(xf["y"] * scale_y, 2)
            w, h = round(max(12.0, xf["w"] * scale_x), 2), round(max(12.0, xf["h"] * scale_y), 2)

            tbl = node.find(".//a:tbl", NS)
            if tbl is not None:
                cols = [float(c.attrib.get("w", "1")) for c in tbl.findall("a:tblGrid/a:gridCol", NS)]
                tot_c = sum(cols) or 1.0
                col_pcts = [round(c / tot_c * 100.0, 2) for c in cols]
                header_bg = f"#{theme_colors.get('accent1', '2563EB')}"
                stripe_bg = f"#{theme_colors.get('lt2', 'F3F4F6')}"
                rows_out = []
                for r_i, tr in enumerate(tbl.findall("a:tr", NS)):
                    row_cells = []
                    for c_i, tc in enumerate(tr.findall("a:tc", NS)):
                        if tc.attrib.get("hMerge") in ("1", "true") or tc.attrib.get("vMerge") in ("1", "true"):
                            continue
                        tcpr = tc.find("a:tcPr", NS)
                        bg_hx, bg_a = _resolve_color_hex_alpha(tcpr.find("a:solidFill", NS), theme_colors, clr_map) if tcpr is not None else (None, 1.0)
                        bg_css = _color_css(bg_hx, bg_a)
                        def_tc = "#111827"
                        if not bg_css:
                            if r_i == 0:
                                bg_css = header_bg
                                def_tc = "#FFFFFF"
                            else:
                                bg_css = "#FFFFFF" if r_i % 2 == 1 else stripe_bg
                        c_paras, _, _ = _extract_paragraphs(
                            tc.find("a:txBody", NS), theme_colors, clr_map, major_font, minor_font, None, def_tc
                        )
                        c_text = "\n".join(p["text"] for p in c_paras)
                        fp = c_paras[0] if c_paras else {}
                        row_cells.append({
                            "r": r_i,
                            "c": c_i,
                            "colspan": int(tc.attrib.get("gridSpan", "1")),
                            "rowspan": int(tc.attrib.get("rowSpan", "1")),
                            "text": c_text,
                            "bg": bg_css,
                            "bg_hex": bg_hx or (bg_css if isinstance(bg_css, str) and bg_css.startswith("#") and len(bg_css) == 7 else None),
                            "color": fp.get("color", def_tc),
                            "bold": bool(fp.get("bold", r_i == 0)),
                            "italic": bool(fp.get("italic", False)),
                            "underline": bool(fp.get("underline", False)),
                            "align": fp.get("align", "left"),
                            "font_size_pt": fp.get("font_size_pt", 12.0),
                        })
                    rows_out.append(row_cells)
                elements.append({
                    "shape_id": sid,
                    "name": sname,
                    "kind": "table",
                    "editable": True,
                    "x": x, "y": y, "w": w, "h": h, "rot": 0,
                    "col_pcts": col_pcts,
                    "rows": rows_out,
                })
                continue

            ch_ref = node.find(".//c:chart", NS)
            if ch_ref is not None:
                rid = ch_ref.attrib.get(f"{{{NS['r']}}}id")
                if rid and rid in rels:
                    svg = _render_chart_svg(zf, rels[rid][1], w, h, theme_colors)
                    if svg:
                        elements.append({
                            "shape_id": sid,
                            "name": sname,
                            "kind": "chart",
                            "editable": True,
                            "x": x, "y": y, "w": w, "h": h, "rot": 0,
                            "svg": svg,
                        })

        elif tag == "grpSp":
            cnv = node.find("p:nvGrpSpPr/p:cNvPr", NS)
            sid = int(cnv.attrib.get("id", "0")) if cnv is not None else 0
            sname = cnv.attrib.get("name", f"Group {sid}") if cnv is not None else f"Group {sid}"
            g_xfrm = node.find("p:grpSpPr/a:xfrm", NS)
            xf = _extract_xfrm(g_xfrm)
            if xf is None:
                continue
            ch_off = g_xfrm.find("a:chOff", NS)
            ch_ext = g_xfrm.find("a:chExt", NS)
            cx = float(ch_off.attrib.get("x", "0")) if ch_off is not None else 0.0
            cy = float(ch_off.attrib.get("y", "0")) if ch_off is not None else 0.0
            cw = float(ch_ext.attrib.get("cx", str(xf["w"]))) if ch_ext is not None else xf["w"]
            ch = float(ch_ext.attrib.get("cy", str(xf["h"]))) if ch_ext is not None else xf["h"]
            sx = (xf["w"] * scale_x) / max(1.0, cw)
            sy = (xf["h"] * scale_y) / max(1.0, ch)
            children = _parse_elements_from_tree(
                node, sx, sy, theme_colors, clr_map, major_font, minor_font,
                zf, rels, media_cache, l_phs, m_phs, skip_placeholders=skip_placeholders, max_img_dim=max_img_dim,
            )
            # Shift children into group-local coordinates
            for c in children:
                c["x"] = round(c["x"] - cx * sx, 2)
                c["y"] = round(c["y"] - cy * sy, 2)
            elements.append({
                "shape_id": sid,
                "name": sname,
                "kind": "group",
                "editable": True,
                "x": round(xf["x"] * scale_x, 2),
                "y": round(xf["y"] * scale_y, 2),
                "w": round(max(4.0, xf["w"] * scale_x), 2),
                "h": round(max(4.0, xf["h"] * scale_y), 2),
                "rot": round(xf["rot"], 1),
                "children": children,
            })

    return elements


def get_deck_overview(file_path: str) -> Dict[str, Any]:
    real = os.path.realpath(os.path.expanduser(file_path))
    if not os.path.isfile(real):
        raise FileNotFoundError(f"PPTX not found: {file_path}")
    st = os.stat(real)
    with zipfile.ZipFile(real, "r") as zf:
        emu_w, emu_h, slide_parts = _extract_slide_parts(zf)
        canvas_h = max(360, int(round(CANVAS_W * emu_h / max(1.0, emu_w))))
        outline = []
        for idx, spath in enumerate(slide_parts, start=1):
            if spath not in zf.namelist():
                continue
            sr = ET.fromstring(zf.read(spath))
            srels = _load_rels(zf, spath)
            title, notes = _extract_notes_and_title(zf, sr, srels)
            outline.append({"index": idx, "title": title or f"Slide {idx}", "has_notes": bool(notes)})
    return {
        "has_deck": True,
        "file_path": real,
        "file_name": os.path.basename(real),
        "file_size": st.st_size,
        "mtime": st.st_mtime,
        "width": CANVAS_W,
        "height": canvas_h,
        "emu_w": int(emu_w),
        "emu_h": int(emu_h),
        "slide_count": len(outline),
        "outline": outline,
    }


def get_slide_scene(file_path: str, slide_index: int, max_img_dim: int = 1400) -> Dict[str, Any]:
    real = os.path.realpath(os.path.expanduser(file_path))
    if not os.path.isfile(real):
        raise FileNotFoundError(f"PPTX not found: {file_path}")
    with zipfile.ZipFile(real, "r") as zf:
        emu_w, emu_h, slide_parts = _extract_slide_parts(zf)
        if not slide_parts:
            raise ValueError("Presentation has no slides")
        idx = max(1, min(len(slide_parts), int(slide_index)))
        spath = slide_parts[idx - 1]
        canvas_h = max(360, int(round(CANVAS_W * emu_h / max(1.0, emu_w))))
        scale_x = CANVAS_W / max(1.0, emu_w)
        scale_y = canvas_h / max(1.0, emu_h)
        pt_to_px = 12700.0 * scale_x

        theme_colors, major_font, minor_font = _parse_theme(zf)
        media_cache: Dict[str, str] = {}

        s_root = ET.fromstring(zf.read(spath))
        s_rels = _load_rels(zf, spath)

        layout_root, layout_rels, layout_phs = None, {}, {}
        master_root, master_rels, master_phs, clr_map = None, {}, {}, dict(DEFAULT_CLR_MAP)
        for _rid, (rtype, target) in s_rels.items():
            if "slideLayout" in rtype and target in zf.namelist():
                layout_root = ET.fromstring(zf.read(target))
                layout_rels = _load_rels(zf, target)
                layout_phs = _collect_placeholders(layout_root)
                break
        if layout_rels:
            for _rid, (rtype, target) in layout_rels.items():
                if "slideMaster" in rtype and target in zf.namelist():
                    master_root = ET.fromstring(zf.read(target))
                    master_rels = _load_rels(zf, target)
                    master_phs = _collect_placeholders(master_root)
                    cm = master_root.find("p:clrMap", NS)
                    if cm is not None:
                        for k, v in cm.attrib.items():
                            clr_map[k] = v
                    break

        bg_css = None
        for rt, rl in [(s_root, s_rels), (layout_root, layout_rels), (master_root, master_rels)]:
            if rt is None:
                continue
            bg_pr = rt.find("p:cSld/p:bg/p:bgPr", NS)
            if bg_pr is not None:
                bg_css, _ = _parse_fill(bg_pr, theme_colors, clr_map, zf, rl, media_cache, max_dim=max_img_dim)
                if bg_css:
                    break
            bg_ref = rt.find("p:cSld/p:bg/p:bgRef", NS)
            if bg_ref is not None:
                hx, a = _resolve_color_hex_alpha(bg_ref, theme_colors, clr_map)
                bg_css = _color_css(hx, a)
                if bg_css:
                    break
        if not bg_css:
            bg_css = f"#{theme_colors.get('lt1', 'FFFFFF')}"

        backdrop_elements: List[Dict[str, Any]] = []
        show_master = s_root.attrib.get("showMasterSp", "1") != "0"
        if show_master and master_root is not None:
            mt = master_root.find("p:cSld/p:spTree", NS)
            if mt is not None:
                backdrop_elements.extend(
                    _parse_elements_from_tree(
                        mt, scale_x, scale_y, theme_colors, clr_map, major_font, minor_font,
                        zf, master_rels, media_cache, {}, {}, skip_placeholders=True, max_img_dim=max_img_dim,
                    )
                )
        if show_master and layout_root is not None:
            lt = layout_root.find("p:cSld/p:spTree", NS)
            if lt is not None:
                backdrop_elements.extend(
                    _parse_elements_from_tree(
                        lt, scale_x, scale_y, theme_colors, clr_map, major_font, minor_font,
                        zf, layout_rels, media_cache, {}, master_phs, skip_placeholders=True, max_img_dim=max_img_dim,
                    )
                )

        s_tree = s_root.find("p:cSld/p:spTree", NS)
        elements = (
            _parse_elements_from_tree(
                s_tree, scale_x, scale_y, theme_colors, clr_map, major_font, minor_font,
                zf, s_rels, media_cache, layout_phs, master_phs, skip_placeholders=False, max_img_dim=max_img_dim,
            )
            if s_tree is not None
            else []
        )

        title, notes = _extract_notes_and_title(zf, s_root, s_rels)
        return {
            "index": idx,
            "title": title or f"Slide {idx}",
            "notes": notes,
            "width": CANVAS_W,
            "height": canvas_h,
            "pt_to_px": round(pt_to_px, 4),
            "background": bg_css,
            "backdrop": backdrop_elements,
            "elements": elements,
        }


def get_thumbnails_batch(file_path: str, start: int = 1, count: int = 4) -> Dict[str, Any]:
    ov = get_deck_overview(file_path)
    total = ov["slide_count"]
    s_idx = max(1, int(start))
    e_idx = min(total, s_idx + max(1, min(8, int(count))) - 1)
    scenes = []
    for i in range(s_idx, e_idx + 1):
        sc = get_slide_scene(file_path, i, max_img_dim=280)
        scenes.append(sc)
    return {
        "width": ov["width"],
        "height": ov["height"],
        "thumbnails": scenes,
    }


# ---------------------------------------------------------------------------
# Lossless surgical write-back via python-pptx
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_str: str) -> RGBColor:
    h = hex_str.strip().lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _align_enum(align_str: str):
    return {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }.get(align_str, PP_ALIGN.LEFT)


def _atomic_save_prs(prs: Presentation, real_path: str) -> None:
    parent = os.path.dirname(real_path)
    fd, tmp_path = tempfile.mkstemp(prefix=".ppt_edit_", suffix=".pptx", dir=parent)
    os.close(fd)
    try:
        prs.save(tmp_path)
        os.replace(tmp_path, real_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise


def save_slide_edits(
    file_path: str,
    slide_index: int,
    ops: List[Dict[str, Any]],
    notes: Optional[str] = None,
) -> Dict[str, Any]:
    """Surgically patch only edited shapes/notes inside the original .pptx file."""
    real = os.path.realpath(os.path.expanduser(file_path))
    if not os.path.isfile(real):
        raise FileNotFoundError(f"PPTX not found: {file_path}")

    prs = Presentation(real)
    total = len(prs.slides)
    if total == 0:
        raise ValueError("Presentation has no slides")
    idx = max(1, min(total, int(slide_index)))
    slide = prs.slides[idx - 1]

    emu_w = float(prs.slide_width)
    emu_h = float(prs.slide_height)
    canvas_h = max(360, int(round(CANVAS_W * emu_h / max(1.0, emu_w))))
    emu_per_x = emu_w / CANVAS_W
    emu_per_y = emu_h / canvas_h

    shape_map = {s.shape_id: s for s in slide.shapes}
    created_map: Dict[str, int] = {}
    applied = 0

    for op in ops or []:
        kind = op.get("op")
        if kind == "add_textbox":
            x = int(round(float(op.get("x", 100)) * emu_per_x))
            y = int(round(float(op.get("y", 100)) * emu_per_y))
            w = int(round(max(40.0, float(op.get("w", 280))) * emu_per_x))
            h = int(round(max(24.0, float(op.get("h", 60))) * emu_per_y))
            tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
            tf = tb.text_frame
            tf.word_wrap = True
            lines = str(op.get("text", "双击输入文字")).split("\n")
            for li, line in enumerate(lines):
                p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
                if op.get("align"):
                    p.alignment = _align_enum(op["align"])
                run = p.add_run()
                run.text = line
                run.font.size = Pt(float(op.get("font_size_pt", 18)))
                run.font.bold = bool(op.get("bold", False))
                run.font.italic = bool(op.get("italic", False))
                run.font.underline = bool(op.get("underline", False))
                run.font.name = op.get("font") or "PingFang SC"
                if op.get("color"):
                    run.font.color.rgb = _hex_to_rgb(op["color"])
            if op.get("fill_hex"):
                tb.fill.solid()
                tb.fill.fore_color.rgb = _hex_to_rgb(op["fill_hex"])
            shape_map[tb.shape_id] = tb
            if op.get("temp_id"):
                created_map[str(op["temp_id"])] = tb.shape_id
            applied += 1
            continue

        raw_sid = op.get("shape_id")
        if isinstance(raw_sid, str) and raw_sid in created_map:
            sid = created_map[raw_sid]
        else:
            try:
                sid = int(raw_sid)
            except Exception:
                continue
        shape = shape_map.get(sid)
        if shape is None:
            continue

        if kind == "delete_shape":
            el = shape._element
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
                shape_map.pop(sid, None)
                applied += 1
            continue

        if kind == "geometry":
            if "x" in op:
                shape.left = Emu(int(round(float(op["x"]) * emu_per_x)))
            if "y" in op:
                shape.top = Emu(int(round(float(op["y"]) * emu_per_y)))
            if "w" in op:
                shape.width = Emu(int(round(max(4.0, float(op["w"])) * emu_per_x)))
            if "h" in op:
                shape.height = Emu(int(round(max(4.0, float(op["h"])) * emu_per_y)))
            applied += 1

        elif kind == "text" and shape.has_text_frame:
            tf = shape.text_frame
            new_lines = str(op.get("text", "")).split("\n")
            existing = list(tf.paragraphs)

            # Capture default run style from first available run
            def _capture_style(p_obj):
                for r in p_obj.runs:
                    return {
                        "size": r.font.size,
                        "bold": r.font.bold,
                        "italic": r.font.italic,
                        "underline": r.font.underline,
                        "name": r.font.name,
                        "rgb": r.font.color.rgb if (r.font.color and r.font.color.type == 1) else None,
                    }
                return {}

            fallback_style = {}
            for ep in existing:
                st = _capture_style(ep)
                if st:
                    fallback_style = st
                    break

            for li, line in enumerate(new_lines):
                if li < len(existing):
                    p = existing[li]
                    st = _capture_style(p) or fallback_style
                    # Clear existing runs cleanly while preserving paragraph XML properties (alignment/bullets)
                    for r in list(p.runs):
                        r._r.getparent().remove(r._r)
                    # Remove any stray line breaks
                    for br in list(p._p.findall(f"{{{NS['a']}}}br")):
                        p._p.remove(br)
                else:
                    p = tf.add_paragraph()
                    st = fallback_style
                run = p.add_run()
                run.text = line
                if st.get("size"):
                    run.font.size = st["size"]
                if st.get("bold") is not None:
                    run.font.bold = st["bold"]
                if st.get("italic") is not None:
                    run.font.italic = st["italic"]
                if st.get("underline") is not None:
                    run.font.underline = st["underline"]
                if st.get("name"):
                    run.font.name = st["name"]
                if st.get("rgb"):
                    run.font.color.rgb = st["rgb"]

            # Remove extra trailing paragraphs if text became shorter
            for extra_p in existing[len(new_lines):]:
                extra_p._p.getparent().remove(extra_p._p)
            applied += 1

        elif kind == "style":
            if "fill_transparent" in op and op["fill_transparent"]:
                try:
                    shape.fill.background()
                except Exception:
                    pass
            elif op.get("fill_hex"):
                try:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = _hex_to_rgb(op["fill_hex"])
                except Exception:
                    pass

            if shape.has_text_frame:
                tf = shape.text_frame
                for p in tf.paragraphs:
                    if op.get("align"):
                        p.alignment = _align_enum(op["align"])
                    runs = list(p.runs)
                    if not runs and p.text:
                        r = p.add_run()
                        r.text = p.text
                        runs = [r]
                    for r in runs:
                        if op.get("font_size_pt") is not None:
                            r.font.size = Pt(float(op["font_size_pt"]))
                        if op.get("bold") is not None:
                            r.font.bold = bool(op["bold"])
                        if op.get("italic") is not None:
                            r.font.italic = bool(op["italic"])
                        if op.get("underline") is not None:
                            r.font.underline = bool(op["underline"])
                        if op.get("color"):
                            r.font.color.rgb = _hex_to_rgb(op["color"])
            applied += 1

        elif kind == "table_cell" and shape.has_table:
            r_i = int(op.get("r", 0))
            c_i = int(op.get("c", 0))
            tbl = shape.table
            if 0 <= r_i < len(tbl.rows) and 0 <= c_i < len(tbl.columns):
                cell = tbl.cell(r_i, c_i)
                tf = cell.text_frame
                p0 = tf.paragraphs[0]
                st = {"align": p0.alignment}
                if p0.runs:
                    r0 = p0.runs[0]
                    st.update({
                        "size": r0.font.size,
                        "bold": r0.font.bold,
                        "italic": r0.font.italic,
                        "underline": r0.font.underline,
                        "name": r0.font.name,
                        "rgb": r0.font.color.rgb if (r0.font.color and r0.font.color.type == 1) else None,
                    })
                cell.text = str(op.get("text", ""))
                for p in cell.text_frame.paragraphs:
                    if st.get("align") is not None:
                        p.alignment = st["align"]
                    for r in p.runs:
                        if st.get("size"):
                            r.font.size = st["size"]
                        if st.get("bold") is not None:
                            r.font.bold = st["bold"]
                        if st.get("italic") is not None:
                            r.font.italic = st["italic"]
                        if st.get("underline") is not None:
                            r.font.underline = st["underline"]
                        if st.get("name"):
                            r.font.name = st["name"]
                        if st.get("rgb"):
                            r.font.color.rgb = st["rgb"]
                applied += 1

        elif kind == "table_cell_style" and shape.has_table:
            tbl = shape.table
            target_cells = []
            if op.get("r") is not None and op.get("c") is not None:
                r_i, c_i = int(op["r"]), int(op["c"])
                if 0 <= r_i < len(tbl.rows) and 0 <= c_i < len(tbl.columns):
                    target_cells.append(tbl.cell(r_i, c_i))
            else:
                for row in tbl.rows:
                    for cell in row.cells:
                        target_cells.append(cell)

            for cell in target_cells:
                if op.get("fill_transparent"):
                    try:
                        cell.fill.background()
                    except Exception:
                        pass
                elif op.get("fill_hex"):
                    try:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = _hex_to_rgb(op["fill_hex"])
                    except Exception:
                        pass
                for p in cell.text_frame.paragraphs:
                    if op.get("align"):
                        p.alignment = _align_enum(op["align"])
                    runs = list(p.runs)
                    if not runs and p.text:
                        r = p.add_run()
                        r.text = p.text
                        runs = [r]
                    for r in runs:
                        if op.get("font_size_pt") is not None:
                            r.font.size = Pt(float(op["font_size_pt"]))
                        if op.get("bold") is not None:
                            r.font.bold = bool(op["bold"])
                        if op.get("italic") is not None:
                            r.font.italic = bool(op["italic"])
                        if op.get("underline") is not None:
                            r.font.underline = bool(op["underline"])
                        if op.get("color"):
                            r.font.color.rgb = _hex_to_rgb(op["color"])
            if target_cells:
                applied += 1

    if notes is not None:
        slide.notes_slide.notes_text_frame.text = str(notes)
        applied += 1

    _atomic_save_prs(prs, real)
    overview = get_deck_overview(real)
    scene = get_slide_scene(real, idx)
    return {
        "ok": True,
        "applied": applied,
        "created_map": created_map,
        "overview": overview,
        "scene": scene,
    }


def apply_deck_slide_op(
    file_path: str,
    action: str,
    slide_index: int,
    target_index: Optional[int] = None,
) -> Dict[str, Any]:
    """Reorder or delete a slide inside the original .pptx file atomically."""
    real = os.path.realpath(os.path.expanduser(file_path))
    prs = Presentation(real)
    sld_id_lst = prs.slides._sldIdLst
    total = len(sld_id_lst)
    if total == 0:
        raise ValueError("Presentation has no slides")
    idx0 = max(0, min(total - 1, int(slide_index) - 1))

    new_active = idx0 + 1
    if action == "move" and target_index is not None:
        dst0 = max(0, min(total - 1, int(target_index) - 1))
        if idx0 != dst0:
            el = sld_id_lst[idx0]
            sld_id_lst.remove(el)
            sld_id_lst.insert(dst0, el)
        new_active = dst0 + 1
    elif action == "delete":
        if total <= 1:
            raise ValueError("Cannot delete the only remaining slide")
        r_id = sld_id_lst[idx0].rId
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id_lst[idx0])
        new_active = max(1, min(total - 1, idx0 + 1))
    else:
        raise ValueError(f"Unsupported slide action: {action}")

    _atomic_save_prs(prs, real)
    overview = get_deck_overview(real)
    overview["active_slide"] = new_active
    return overview
