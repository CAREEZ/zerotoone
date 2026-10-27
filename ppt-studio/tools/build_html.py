#!/usr/bin/env python3
"""spec(JSON) → 단일 파일 HTML 슬라이드 (모션 + 브라우저 내 직접 수정).

사용법:
    python3 ppt-studio/tools/build_html.py <spec.json> <out.html>

- build_deck.py와 같은 spec을 쓴다. 같은 원고로 PPTX(제출·보고용)와 HTML(발표·피칭용)을 함께 만든다.
- 1280×720(16:9) 고정 캔버스를 화면에 맞춰 확대·축소한다. A4 spec은 거부한다(기관 제출은 PPTX/PDF).
- Pretendard(KS X 1001 2,350자 + 라틴 서브셋)를 base64로 내장 → 인터넷 없는 발표장 PC에서도 동일하게 렌더링.
- 단축키: →/Space 다음, ← 이전, F 전체화면, E 편집 모드, Ctrl+S 수정본 저장, N 발표자 노트, P 인쇄(PDF)
- 차트는 SVG(애니메이션 포함)로 그린다. 차트 수치 수정은 spec에서 하고 다시 빌드한다.
"""
import base64
import html
import json
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_deck import CHECK_TAG, LIMITS, checks, lint_sequence, load_brand, split_keyword, warnings  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
FONT_DIR = ROOT / "assets" / "fonts"
FONT_WEIGHTS = {400: "Regular", 600: "SemiBold", 700: "Bold", 800: "ExtraBold"}
W, H = 1280, 720


def esc(s, page=None, where=""):
    """HTML 이스케이프 + [확인필요] 강조 + 줄바꿈."""
    raw = str(s)
    if CHECK_TAG in raw and page is not None:
        checks.append(f"p{page} {where}: {raw}")
    out = html.escape(raw).replace("\n", "<br>")
    return out.replace(html.escape(CHECK_TAG), f'<mark class="chk">{html.escape(CHECK_TAG)}</mark>')


def fmt(v, pct=False):
    if pct:
        return f"{v * 100:.0f}%"
    if isinstance(v, float) and not v.is_integer():
        return f"{v:,.1f}"
    return f"{int(v):,}"


def font_faces():
    css = []
    for wt, name in FONT_WEIGHTS.items():
        data = base64.b64encode((FONT_DIR / f"Pretendard-{name}.subset.woff2").read_bytes()).decode()
        css.append(
            f"@font-face{{font-family:'Pretendard';font-weight:{wt};font-display:block;"
            f"src:url(data:font/woff2;base64,{data}) format('woff2');}}"
        )
    return "\n".join(css)


class HtmlDeck:
    def __init__(self, spec):
        self.spec = spec
        if spec.get("size", "16:9") != "16:9":
            sys.exit("HTML 모드는 16:9 발표용만 지원함. A4·기관 제출은 build_deck.py(PPTX/PDF)를 사용.")
        self.brand = load_brand(spec.get("brand", "default"))
        self.c = self.brand["colors"]
        self.page = 0
        self.n = 0  # 애니메이션 순번

    def a(self):
        """등장 애니메이션 순번 style 속성."""
        self.n += 1
        return f'style="--i:{self.n}"'

    # ---------- 공통 프레임 ----------
    def frame(self, d, body, cls=""):
        title = d.get("title", "")
        if len(title) > LIMITS["title"]:
            warnings.append(f"p{self.page} 제목 {len(title)}자 > {LIMITS['title']}자: {title}")
        foot = "   |   ".join(x for x in [f"출처: {d['source']}" if d.get("source") else "", self.brand["footer"]] if x)
        head = ""
        if d.get("chapter"):
            head += f'<div class="chapter ed a" {self.a()}>{esc(d["chapter"], self.page, "chapter")}</div>'
        head += f'<h2 class="title ed a" {self.a()}>{esc(title, self.page, "title")}</h2>'
        if d.get("subtitle"):
            head += f'<p class="subtitle ed a" {self.a()}>{esc(d["subtitle"], self.page, "subtitle")}</p>'
        return (
            f'<header>{head}</header><main class="body {cls}">{body}</main>'
            f'<footer><span class="ed">{esc(foot, self.page, "footer")}</span><span>{self.page}</span></footer>'
        )

    # ---------- 유형별 ----------
    def cover(self, d):
        meta = " · ".join(x for x in [d.get("date"), d.get("presenter"), self.brand["company"]] if x)
        sub = f'<p class="cover-sub ed a" {self.a()}>{esc(d["subtitle"], self.page, "cover")}</p>' if d.get("subtitle") else ""
        return ("dark", f'<div class="cover"><div class="rule a" {self.a()}></div>'
                f'<h1 class="ed a" {self.a()}>{esc(d["title"], self.page, "cover")}</h1>{sub}'
                f'</div><div class="meta ed a" {self.a()}>{esc(meta)}</div>')

    def closing(self, d):
        contact = f'<p class="cover-sub ed a" {self.a()}>{esc(d["contact"], self.page, "closing")}</p>' if d.get("contact") else ""
        return ("dark", f'<div class="cover closing"><h1 class="ed a" {self.a()}>{esc(d.get("title", "감사합니다"), self.page, "closing")}</h1>{contact}</div>')

    def section(self, d):
        return ("light2", f'<div class="section"><div class="sec-no a" {self.a()}>{esc(d.get("no", ""))}</div>'
                f'<div class="rule a" {self.a()}></div><h2 class="ed a" {self.a()}>{esc(d["title"], self.page, "section")}</h2></div>')

    def agenda(self, d):
        rows = "".join(
            f'<li class="a" {self.a()}><span class="no">{i + 1:02d}</span><span class="ed">{esc(it, self.page, "agenda")}</span></li>'
            for i, it in enumerate(d["items"]))
        return ("", self.frame({"title": d.get("title", "목차"), **{k: v for k, v in d.items() if k != "title"}},
                               f'<ol class="agenda">{rows}</ol>'))

    def bullets(self, d):
        items = d.get("bullets", [])
        if len(items) > LIMITS["bullets"]:
            warnings.append(f"p{self.page} 불릿 {len(items)}개 > {LIMITS['bullets']}개")
        lis = []
        for b in items:
            if len(b) > LIMITS["bullet_len"]:
                warnings.append(f"p{self.page} 불릿 {len(b)}자 > {LIMITS['bullet_len']}자: {b[:30]}…")
            k, v = split_keyword(b)
            inner = (f"<b>{esc(k)}</b> {esc(v, self.page, 'bullets')}" if k else esc(b, self.page, "bullets"))
            lis.append(f'<li class="ed a" {self.a()}>{inner}</li>')
        msg = f'<div class="message ed a" {self.a()}>{esc(d["message"], self.page, "message")}</div>' if d.get("message") else ""
        return ("", self.frame(d, f'{msg}<ul class="bullets">{"".join(lis)}</ul>'))

    def kpi(self, d):
        items = d["items"]
        if len(items) > LIMITS["kpi"]:
            warnings.append(f"p{self.page} KPI {len(items)}개 > {LIMITS['kpi']}개")
        cards = "".join(
            f'<div class="card a" {self.a()}><div class="kv count ed">{esc(it["value"], self.page, "kpi")}</div>'
            f'<div class="kl ed">{esc(it["label"], self.page, "kpi")}</div>'
            + (f'<div class="kd ed">{esc(it["desc"], self.page, "kpi")}</div>' if it.get("desc") else "")
            + "</div>" for it in items)
        return ("", self.frame(d, f'<div class="cards n{len(items)}">{cards}</div>'))

    def table(self, d):
        cols, rows, hl = d["columns"], d["rows"], d.get("highlight_col")
        if len(rows) > LIMITS["table_rows"]:
            warnings.append(f"p{self.page} 표 {len(rows)}행 > {LIMITS['table_rows']}행")
        th = "".join(f'<th class="ed{" hl" if i == hl else ""}">{esc(c)}</th>' for i, c in enumerate(cols))
        trs = "".join(
            f'<tr class="a" {self.a()}>' + "".join(
                f'<td class="ed{" hl" if i == hl else ""}">{esc(v, self.page, "table")}</td>' for i, v in enumerate(r)) + "</tr>"
            for r in rows)
        return ("", self.frame(d, f'<table class="tbl"><thead><tr class="a" {self.a()}>{th}</tr></thead><tbody>{trs}</tbody></table>'))

    def compare(self, d):
        cols = ""
        for key, cls in (("left", "before"), ("right", "after")):
            col = d[key]
            lis = "".join(f'<li class="ed">{esc(b, self.page, "compare")}</li>' for b in col["bullets"])
            cols += f'<div class="ccol {cls} a" {self.a()}><h3 class="ed">{esc(col["title"])}</h3><ul>{lis}</ul></div>'
        return ("", self.frame(d, f'<div class="compare">{cols}<div class="vs a" {self.a()}>→</div></div>'))

    def process(self, d):
        steps = d["steps"]
        if len(steps) > LIMITS["steps"]:
            warnings.append(f"p{self.page} 단계 {len(steps)}개 > {LIMITS['steps']}개")
        out = []
        for i, st in enumerate(steps):
            last = " last" if i == len(steps) - 1 else ""
            out.append(f'<div class="step{last} a" {self.a()}><div class="sn">{i + 1}</div>'
                       f'<div class="st ed">{esc(st["title"], self.page, "process")}</div>'
                       f'<div class="sd ed">{esc(st.get("desc", ""), self.page, "process")}</div></div>')
            if not last:
                out.append(f'<div class="arrow a" {self.a()}></div>')
        return ("", self.frame(d, f'<div class="process">{"".join(out)}</div>'))

    def chart(self, d):
        kind = d.get("chart", "bar")
        svg = {"bar": self._bar, "stacked": self._bar, "hbar": self._hbar, "line": self._line, "pie": self._pie}[kind](d)
        unit = f'<div class="unit">단위: {esc(d["unit"])}</div>' if d.get("unit") else ""
        side = ""
        if d.get("takeaway"):
            side = f'<div class="takeaway ed a" {self.a()}>{esc(d["takeaway"], self.page, "takeaway")}</div>'
        legend = ""
        if len(d["series"]) > 1 or kind == "pie":
            names = d["categories"] if kind == "pie" else [s["name"] for s in d["series"]]
            legend = '<div class="legend">' + "".join(
                f'<span><i style="background:#{self.pal(i)}"></i>{esc(n)}</span>' for i, n in enumerate(names)) + "</div>"
        return ("", self.frame(d, f'<div class="chartwrap{" withside" if side else ""}"><div class="chart">{unit}{svg}{legend}</div>{side}</div>'))

    def pal(self, i):
        p = self.c["chart"]
        return p[i % len(p)]

    def _bar(self, d, cw=760, ch=400):
        cats, series = d["categories"], d["series"]
        stacked = d.get("chart") == "stacked"
        totals = [sum(s["values"][j] for s in series) for j in range(len(cats))]
        vmax = max(totals) if stacked else max(max(s["values"]) for s in series)
        top, bot = 36, 40
        ph = ch - top - bot
        slot = cw / len(cats)
        ns = 1 if stacked else len(series)
        bw = min(slot * 0.62 / ns, 96)
        hl = d.get("highlight")
        out = [f'<line x1="0" y1="{top + ph}" x2="{cw}" y2="{top + ph}" class="axis"/>']
        for j, cat in enumerate(cats):
            x0 = slot * j + (slot - bw * ns) / 2
            base = top + ph
            for i, s in enumerate(series):
                v = s["values"][j]
                h = ph * v / vmax if vmax else 0
                if len(series) == 1 and hl is not None:
                    col = self.c["accent"] if j == hl else self.c["chart_muted"]
                else:
                    col = self.pal(i)
                x = x0 if stacked else x0 + i * bw
                y = base - h if stacked else top + ph - h
                out.append(f'<rect class="bar" style="--i:{j}" x="{x:.1f}" y="{y:.1f}" width="{bw - 4:.1f}" height="{h:.1f}" rx="3" fill="#{col}"/>')
                if stacked:
                    if h > 22:
                        out.append(f'<text class="lbl in" x="{x + (bw - 4) / 2:.1f}" y="{y + h / 2 + 5:.1f}">{fmt(v)}</text>')
                    base -= h
                else:
                    out.append(f'<text class="lbl" x="{x + (bw - 4) / 2:.1f}" y="{y - 8:.1f}">{fmt(v)}</text>')
            if stacked:
                out.append(f'<text class="lbl" x="{x0 + (bw - 4) / 2:.1f}" y="{base - 8:.1f}">{fmt(totals[j])}</text>')
            out.append(f'<text class="cat" x="{slot * j + slot / 2:.1f}" y="{top + ph + 26}">{esc(cat)}</text>')
        return f'<svg viewBox="0 0 {cw} {ch}" class="svgc">{"".join(out)}</svg>'

    def _hbar(self, d, cw=760, ch=400):
        cats, s = d["categories"], d["series"][0]
        vmax = max(s["values"])
        lw, slot = 150, ch / len(cats)
        bh = min(slot * 0.6, 44)
        hl = d.get("highlight")
        out = []
        for j, (cat, v) in enumerate(zip(cats, s["values"])):
            y = slot * j + (slot - bh) / 2
            w = (cw - lw - 80) * v / vmax if vmax else 0
            col = self.c["accent"] if j == hl else (self.c["chart_muted"] if hl is not None else self.pal(0))
            out.append(f'<text class="cat r" x="{lw - 14}" y="{y + bh / 2 + 6:.1f}">{esc(cat)}</text>')
            out.append(f'<rect class="hbar" style="--i:{j}" x="{lw}" y="{y:.1f}" width="{w:.1f}" height="{bh:.1f}" rx="3" fill="#{col}"/>')
            out.append(f'<text class="lbl l" x="{lw + w + 10:.1f}" y="{y + bh / 2 + 6:.1f}">{fmt(v)}</text>')
        return f'<svg viewBox="0 0 {cw} {ch}" class="svgc">{"".join(out)}</svg>'

    def _line(self, d, cw=760, ch=400):
        cats, series = d["categories"], d["series"]
        vals = [v for s in series for v in s["values"]]
        vmin, vmax = min(0, min(vals)), max(vals)
        top, bot, padx = 36, 40, 50
        ph = ch - top - bot
        sx = (cw - 2 * padx) / max(len(cats) - 1, 1)
        Y = lambda v: top + ph - ph * (v - vmin) / (vmax - vmin or 1)  # noqa: E731
        out = [f'<line x1="0" y1="{top + ph}" x2="{cw}" y2="{top + ph}" class="axis"/>']
        for j, cat in enumerate(cats):
            out.append(f'<text class="cat" x="{padx + sx * j:.1f}" y="{top + ph + 26}">{esc(cat)}</text>')
        for i, s in enumerate(series):
            col = self.pal(i)
            pts = [(padx + sx * j, Y(v)) for j, v in enumerate(s["values"])]
            path = "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts)
            out.append(f'<path class="line" d="{path}" pathLength="1" stroke="#{col}" fill="none"/>')
            for j, ((x, y), v) in enumerate(zip(pts, s["values"])):
                col_vals = [t["values"][j] for t in series]  # 같은 x에서 큰 값은 위, 작은 값은 아래에 표기
                dy = -14 if v >= max(col_vals) or len(series) == 1 else 26
                out.append(f'<circle class="dot" style="--i:{j}" cx="{x:.1f}" cy="{y:.1f}" r="5" fill="#{col}"/>')
                out.append(f'<text class="lbl dot" style="--i:{j}" x="{x:.1f}" y="{y + dy:.1f}" fill="#{col}">{fmt(v)}</text>')
        return f'<svg viewBox="0 0 {cw} {ch}" class="svgc">{"".join(out)}</svg>'

    def _pie(self, d, cw=760, ch=400):
        vals = d["series"][0]["values"]
        total = sum(vals)
        r, cx, cy = 140, cw / 2, ch / 2
        out, acc = [], 0.0
        for i, v in enumerate(vals):
            frac = v / total
            out.append(f'<circle class="seg" cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="#{self.pal(i)}" stroke-width="64" '
                       f'pathLength="100" stroke-dasharray="{frac * 100:.2f} {100 - frac * 100:.2f}" '
                       f'stroke-dashoffset="{-acc * 100:.2f}" transform="rotate(-90 {cx} {cy})"/>')
            ang = (acc + frac / 2) * 2 * math.pi - math.pi / 2
            lx, ly = cx + (r + 70) * math.cos(ang), cy + (r + 70) * math.sin(ang)
            out.append(f'<text class="lbl dot" style="--i:{i}" x="{lx:.1f}" y="{ly + 6:.1f}">{fmt(frac, pct=True)}</text>')
            acc += frac
        return f'<svg viewBox="0 0 {cw} {ch}" class="svgc pie">{"".join(out)}</svg>'

    # ---------- 조립 ----------
    def build(self, out):
        lint_sequence(self.spec)
        slides = []
        for d in self.spec["slides"]:
            self.page += 1
            self.n = 0
            tone, inner = getattr(self, d["type"])(d)
            notes = html.escape(d.get("notes", ""), quote=True)
            slides.append(f'<section class="slide {tone} t-{d["type"]}" data-notes="{notes}">{inner}</section>')
        title = self.spec.get("title") or next((s.get("title") for s in self.spec["slides"] if s["type"] == "cover"), "Slides")
        page = TEMPLATE.replace("/*FONTS*/", font_faces()).replace("/*VARS*/", self.css_vars())
        page = page.replace("<!--TITLE-->", html.escape(str(title).replace("\n", " ")))
        page = page.replace("<!--SLIDES-->", "\n".join(slides))
        Path(out).write_text(page, encoding="utf-8")

    def css_vars(self):
        c = self.c
        return (f"--primary:#{c['primary']};--accent:#{c['accent']};--bg:#{c['background']};--surface:#{c['surface']};"
                f"--text:#{c['text']};--muted:#{c['muted']};--line:#{c['line']};--onp:#{c['on_primary_muted']};"
                f"--hlbg:#{c['highlight_bg']};--warn:#{c['warning']};--cmuted:#{c['chart_muted']};")


TEMPLATE = r"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title><!--TITLE--></title>
<style>
/*FONTS*/
:root{/*VARS*/}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;background:#0b0d12;overflow:hidden}
body{font-family:'Pretendard',system-ui,sans-serif;color:var(--text);word-break:keep-all;-webkit-font-smoothing:antialiased}
.deck{position:absolute;left:50%;top:50%;width:1280px;height:720px;transform-origin:center;transform:translate(-50%,-50%) scale(var(--s,1))}
.slide{position:absolute;inset:0;display:none;background:var(--bg);overflow:hidden;padding:52px 72px 0}
.slide.active{display:block}
.slide.dark{background:var(--primary);color:#fff}
.slide.light2{background:var(--surface)}
header{height:118px}
.chapter{font-size:15px;font-weight:700;color:var(--accent);letter-spacing:.02em;margin-bottom:10px}
.title{font-size:34px;font-weight:800;color:var(--primary);line-height:1.25;letter-spacing:-.02em}
.subtitle{font-size:18px;color:var(--muted);margin-top:8px}
main.body{position:absolute;left:72px;right:72px;top:190px;bottom:72px}
footer{position:absolute;left:72px;right:72px;bottom:22px;display:flex;justify-content:space-between;font-size:13px;color:var(--muted);border-top:1px solid var(--line);padding-top:10px}
mark.chk{background:none;color:var(--warn);font-weight:700}
/* cover / closing / section */
.cover{position:absolute;left:96px;right:96px;top:250px}
.rule{width:72px;height:6px;background:var(--accent);margin-bottom:28px}
.cover h1{font-size:60px;font-weight:800;line-height:1.18;letter-spacing:-.03em}
.cover-sub{font-size:22px;color:var(--onp);margin-top:22px;line-height:1.5}
.meta{position:absolute;left:96px;bottom:56px;font-size:16px;color:var(--onp)}
.closing{top:280px}
.section{position:absolute;left:96px;top:220px}
.sec-no{font-size:120px;font-weight:800;color:var(--accent);line-height:1;margin-bottom:24px}
.section .rule{background:var(--primary)}
.section h2{font-size:44px;font-weight:800;color:var(--primary)}
/* agenda */
.agenda{list-style:none}
.agenda li{display:flex;align-items:baseline;gap:32px;padding:20px 0;border-bottom:1px solid var(--line);font-size:26px;font-weight:600}
.agenda .no{color:var(--accent);font-weight:800;font-size:28px;width:48px}
/* bullets */
.message{font-size:28px;font-weight:700;border-left:6px solid var(--accent);padding:6px 0 6px 22px;margin-bottom:40px;line-height:1.35}
.bullets{list-style:none}
.bullets li{font-size:24px;line-height:1.45;padding:12px 0 12px 30px;position:relative}
.bullets li::before{content:"";position:absolute;left:4px;top:26px;width:9px;height:9px;border-radius:50%;background:var(--accent)}
.bullets b{color:var(--primary);margin-right:6px}
/* kpi */
.cards{display:grid;gap:24px;height:100%;align-content:center}
.cards.n1{grid-template-columns:1fr}.cards.n2{grid-template-columns:repeat(2,1fr)}.cards.n3{grid-template-columns:repeat(3,1fr)}.cards.n4{grid-template-columns:repeat(4,1fr)}
.card{background:var(--surface);border-radius:14px;padding:36px 30px;min-height:280px;display:flex;flex-direction:column;justify-content:center}
.kv{font-size:64px;font-weight:800;color:var(--primary);letter-spacing:-.03em;line-height:1.05}
.cards.n1 .kv,.cards.n2 .kv{font-size:96px}
.card:first-child .kv{color:var(--accent)}
.kl{font-size:20px;font-weight:700;margin-top:18px}
.kd{font-size:15px;color:var(--muted);margin-top:8px}
/* chart */
.chartwrap{display:flex;gap:40px;height:100%}
.chart{flex:1;position:relative;display:flex;flex-direction:column}
.svgc{width:100%;flex:1;min-height:0}
.unit{font-size:13px;color:var(--muted);position:absolute;left:0;top:-26px}
.axis{stroke:var(--line);stroke-width:1.5}
.lbl{font-size:17px;font-weight:700;fill:var(--text);text-anchor:middle;font-family:'Pretendard'}
.lbl.in{fill:#fff;font-size:14px}.lbl.l{text-anchor:start}
.cat{font-size:16px;fill:var(--muted);text-anchor:middle;font-family:'Pretendard'}.cat.r{text-anchor:end}
.line{stroke-width:4;stroke-linecap:round;stroke-linejoin:round}
.legend{display:flex;gap:24px;justify-content:center;font-size:15px;color:var(--muted);padding-top:6px}
.legend i{display:inline-block;width:12px;height:12px;border-radius:3px;margin-right:6px;vertical-align:-1px}
.withside .chart{flex:0 0 66%}
.takeaway{flex:1;align-self:center;font-size:22px;line-height:1.55;border-left:6px solid var(--accent);padding-left:22px;font-weight:600}
/* table */
.tbl{width:100%;border-collapse:collapse;font-size:19px}
.tbl th{background:var(--primary);color:#fff;font-weight:700;padding:14px 18px;text-align:center}
.tbl th:first-child,.tbl td:first-child{text-align:left;font-weight:700}
.tbl td{padding:14px 18px;text-align:center;border-bottom:1px solid var(--line)}
.tbl td.hl{background:var(--hlbg);font-weight:800;color:var(--primary)}
.tbl th.hl{background:var(--accent)}
/* compare */
.compare{display:grid;grid-template-columns:1fr 1fr;gap:72px;height:100%;position:relative}
.ccol{background:var(--surface);border-radius:14px;padding:34px 38px}
.ccol h3{font-size:24px;font-weight:800;margin-bottom:22px;color:var(--muted)}
.ccol.after{background:var(--primary);color:#fff}.ccol.after h3{color:var(--accent)}
.ccol ul{list-style:none}.ccol li{font-size:22px;padding:10px 0;line-height:1.4}
.vs{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);font-size:40px;font-weight:800;color:var(--accent)}
/* process */
.process{display:flex;align-items:flex-start;gap:0;height:100%;padding-top:30px}
.step{flex:1;background:var(--surface);border-radius:14px;padding:30px 24px;height:280px}
.step.last{background:var(--primary);color:#fff}
.sn{width:44px;height:44px;border-radius:50%;background:var(--primary);color:#fff;font-weight:800;font-size:20px;display:flex;align-items:center;justify-content:center;margin-bottom:22px}
.step.last .sn{background:var(--accent)}
.st{font-size:22px;font-weight:800;margin-bottom:12px}
.sd{font-size:17px;line-height:1.5;color:var(--muted)}.step.last .sd{color:var(--onp)}
.arrow{flex:0 0 40px;height:280px;position:relative}
.arrow::after{content:"";position:absolute;left:12px;top:50%;width:14px;height:14px;border-top:4px solid var(--accent);border-right:4px solid var(--accent);transform:translateY(-50%) rotate(45deg)}
/* motion */
@keyframes up{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:none}}
@keyframes grow{from{transform:scaleY(0)}to{transform:scaleY(1)}}
@keyframes growx{from{transform:scaleX(0)}to{transform:scaleX(1)}}
@keyframes draw{from{stroke-dashoffset:1}to{stroke-dashoffset:0}}
@keyframes fade{from{opacity:0}to{opacity:1}}
@keyframes spin{from{opacity:0;transform:rotate(-30deg) scale(.85)}to{opacity:1;transform:none}}
.slide.active .a{animation:up .6s cubic-bezier(.2,.7,.2,1) both;animation-delay:calc(var(--i,0)*90ms)}
.slide.active .bar{transform-box:fill-box;transform-origin:bottom;animation:grow .8s cubic-bezier(.2,.7,.2,1) both;animation-delay:calc(300ms + var(--i,0)*110ms)}
.slide.active .hbar{transform-box:fill-box;transform-origin:left;animation:growx .8s cubic-bezier(.2,.7,.2,1) both;animation-delay:calc(300ms + var(--i,0)*110ms)}
.slide.active .line{stroke-dasharray:1;animation:draw 1.2s ease-out both;animation-delay:.3s}
.slide.active .dot{animation:fade .4s both;animation-delay:calc(700ms + var(--i,0)*120ms)}
.slide.active .bar+.lbl,.slide.active .lbl{animation:fade .5s both;animation-delay:900ms}
.slide.active .pie{transform-origin:center;transform-box:fill-box;animation:spin .9s cubic-bezier(.2,.7,.2,1) both;animation-delay:.25s}
@media (prefers-reduced-motion:reduce){.slide.active *{animation:none!important}}
/* hud */
.hud{position:fixed;right:14px;bottom:12px;display:flex;gap:6px;align-items:center;font:13px 'Pretendard',sans-serif;color:#9aa3b2;opacity:.15;transition:opacity .2s;z-index:10}
.hud:hover{opacity:1}
.hud button{font:inherit;background:#1d2230;color:#dfe5ee;border:1px solid #333b4d;border-radius:6px;padding:4px 9px;cursor:pointer}
body.editing .hud{opacity:1}
body.editing .ed{outline:1px dashed rgba(232,131,58,.6);outline-offset:3px;cursor:text}
body.editing #b-edit{background:var(--accent);color:#fff}
#notes{position:fixed;left:0;right:0;bottom:0;max-height:30vh;overflow:auto;background:rgba(10,12,18,.94);color:#e8ecf2;font:18px/1.6 'Pretendard',sans-serif;padding:16px 28px;display:none;z-index:9}
body.show-notes #notes{display:block}
/* print = PDF */
@page{size:1280px 720px;margin:0}
@media print{
 html,body{background:#fff;overflow:visible;height:auto}
 .deck{position:static;transform:none;width:1280px;height:auto}
 .slide{position:relative;display:block!important;width:1280px;height:720px;page-break-after:always;break-after:page}
 .slide *{animation:none!important}
 .hud,#notes{display:none!important}
}
</style></head>
<body>
<div class="deck" id="deck">
<!--SLIDES-->
</div>
<div id="notes"></div>
<div class="hud" id="hud"><span id="pg"></span>
<button id="b-edit" title="E">편집</button><button id="b-save" title="Ctrl+S">저장</button>
<button id="b-notes" title="N">노트</button><button id="b-full" title="F">전체화면</button><button id="b-print" title="P">PDF</button></div>
<script>
(()=>{
const S=[...document.querySelectorAll('.slide')];let cur=0;
const deck=document.getElementById('deck'),pg=document.getElementById('pg'),notes=document.getElementById('notes');
const reduce=matchMedia('(prefers-reduced-motion: reduce)').matches;
function fit(){deck.style.setProperty('--s',Math.min(innerWidth/1280,innerHeight/720))}
function countUp(el){
  if(reduce||document.body.classList.contains('editing'))return;
  const t=el.dataset.final||el.textContent;el.dataset.final=t;
  const m=t.match(/^([^\d]*)([\d,]*\.?\d+)(.*)$/s);if(!m)return;
  const num=parseFloat(m[2].replace(/,/g,'')),dec=(m[2].split('.')[1]||'').length,comma=m[2].includes(',');
  const t0=performance.now(),dur=900;
  (function step(now){const p=Math.min((now-t0)/dur,1),e=1-Math.pow(1-p,3);
    let v=(num*e).toFixed(dec);if(comma)v=Number(v).toLocaleString('en-US',{minimumFractionDigits:dec,maximumFractionDigits:dec});
    el.textContent=m[1]+v+m[3];if(p<1)requestAnimationFrame(step);else el.textContent=t;})(t0);
}
function show(i){
  cur=Math.max(0,Math.min(S.length-1,i));
  S.forEach((s,k)=>s.classList.toggle('active',k===cur));
  pg.textContent=(cur+1)+' / '+S.length;
  notes.textContent=S[cur].dataset.notes||'(노트 없음)';
  history.replaceState(null,'','#'+(cur+1));
  S[cur].querySelectorAll('.count').forEach(countUp);
}
function toggleEdit(){
  const on=document.body.classList.toggle('editing');
  document.querySelectorAll('.ed').forEach(e=>on?e.setAttribute('contenteditable','true'):e.removeAttribute('contenteditable'));
}
function save(){
  if(document.body.classList.contains('editing'))toggleEdit();
  document.querySelectorAll('.count').forEach(e=>{if(e.dataset.final){e.textContent=e.dataset.final;delete e.dataset.final}});
  const doc=document.documentElement.cloneNode(true);
  doc.querySelectorAll('.slide').forEach(s=>s.classList.remove('active'));
  doc.querySelector('body').className='';
  const blob=new Blob(['<!doctype html>\n'+doc.outerHTML],{type:'text/html'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download=(location.pathname.split('/').pop()||'slides.html').replace(/(_edited)?\.html?$/,'')+'_edited.html';a.click();
}
document.addEventListener('keydown',e=>{
  const editing=document.body.classList.contains('editing');
  if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();save();return}
  if(e.key==='Escape'&&editing){toggleEdit();return}
  if(editing)return;
  const k=e.key;
  if(['ArrowRight','PageDown',' ','Enter'].includes(k)){e.preventDefault();show(cur+1)}
  else if(['ArrowLeft','PageUp','Backspace'].includes(k)){e.preventDefault();show(cur-1)}
  else if(k==='Home')show(0);else if(k==='End')show(S.length-1);
  else if(k==='f'||k==='F'){document.fullscreenElement?document.exitFullscreen():document.documentElement.requestFullscreen()}
  else if(k==='e'||k==='E')toggleEdit();
  else if(k==='n'||k==='N')document.body.classList.toggle('show-notes');
  else if(k==='p'||k==='P')print();
});
deck.addEventListener('click',e=>{if(document.body.classList.contains('editing'))return;
  const r=deck.getBoundingClientRect();show(e.clientX<r.left+r.width*0.3?cur-1:cur+1)});
document.getElementById('b-edit').onclick=toggleEdit;document.getElementById('b-save').onclick=save;
document.getElementById('b-notes').onclick=()=>document.body.classList.toggle('show-notes');
document.getElementById('b-full').onclick=()=>document.documentElement.requestFullscreen();
document.getElementById('b-print').onclick=()=>print();
addEventListener('resize',fit);fit();
show((parseInt(location.hash.slice(1))||1)-1);
})();
</script>
</body></html>
"""


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    spec = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    deck = HtmlDeck(spec)
    deck.build(sys.argv[2])
    size = Path(sys.argv[2]).stat().st_size / 1e6
    print(f"생성 완료: {sys.argv[2]} ({deck.page}장, 브랜드={spec.get('brand', 'default')}, {size:.1f}MB, 폰트 내장)")
    if warnings:
        print("\n[분량·규칙 경고]")
        for w in warnings:
            print(" -", w)
    if checks:
        print(f"\n[{CHECK_TAG} 항목 {len(checks)}건 — 발표 전 반드시 확인]")
        for c in checks:
            print(" -", c)


if __name__ == "__main__":
    main()
