"""Build the TCM kitchen recipe book PDF.

Usage:  python3 build.py
Reads ../../TCM Menu 3.0.xlsx for menu order, recipes.py for ingredients/plating,
images/<S.No>.jpg for plate photos (drop your own photo with the same name to replace one),
and credits.json for photo attributions.  Writes ../TCM Recipe Book.pdf (+ .html).
"""
import html, json, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
BOOK = os.path.dirname(HERE)
ROOT = os.path.dirname(BOOK)
sys.path.insert(0, HERE)
from recipes import R

PER_PAGE = 4
e = html.escape


def load_menu():
    ws = openpyxl.load_workbook(os.path.join(ROOT, "TCM Menu 3.0.xlsx")).active
    rows = []
    for sno, cat, item, desc in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        if sno is None or item is None:
            continue
        rows.append(dict(sno=int(sno), cat=cat.strip(), name=item.strip(), desc=(desc or "").strip()))
    return rows


def load_credits():
    p = os.path.join(HERE, "credits.json")
    return json.load(open(p)) if os.path.exists(p) else {}


def card(it, credits):
    r = R.get(it["sno"], dict(ing=["(add ingredients)"], plate=""))
    img = f"images/{it['sno']}.jpg"
    has_img = os.path.exists(os.path.join(BOOK, img))
    c = credits.get(str(it["sno"]))
    tag = '<span class="ref">Reference photo</span>' if c else ""
    photo = (f'<div class="photo" style="background-image:url(\'{img}\')">{tag}</div>' if has_img
             else '<div class="photo empty">Add plating photo<br><small>images/%d.jpg</small></div>' % it["sno"])
    ings = "".join(f"<li>{e(x)}</li>" for x in r["ing"])
    cols = " two" if len(r["ing"]) >= 6 and sum(len(x) for x in r["ing"]) < 260 else ""
    desc = f'<p class="desc">{e(it["desc"])}</p>' if it["desc"] else ""
    verify = '<span class="verify" title="Name is open to interpretation — check against your recipe">✓ verify</span>' if r.get("verify") else ""
    return f"""
    <article class="card">
      {photo}
      <div class="body">
        <div class="meta"><span class="cat">{e(it['cat'])}</span><span class="sno">#{it['sno']}</span>{verify}</div>
        <h2>{e(it['name'])}</h2>{desc}
        <h3>Main ingredients</h3>
        <ul class="ing{cols}">{ings}</ul>
        <div class="plate"><b>Plating</b> {e(r['plate'])}</div>
      </div>
    </article>"""


CSS = """
@page { size: A4; margin: 0; }
* { box-sizing: border-box; }
:root { --ink:#1f1a17; --muted:#6b625b; --accent:#b4441f; --paper:#fffdf9; --line:#e8dfd5; --chip:#f6ebe2; }
body { margin:0; font-family: 'Inter', 'Helvetica Neue', Arial, sans-serif; color:var(--ink); background:#ddd; }
.page { width:210mm; height:297mm; background:var(--paper); page-break-after:always; position:relative;
        padding:11mm 10mm 12mm; display:flex; flex-direction:column; overflow:hidden; margin:0 auto; }
@media screen { .page { margin:8mm auto; box-shadow:0 2px 12px rgba(0,0,0,.15); } }
.phead { display:flex; justify-content:space-between; align-items:baseline; border-bottom:1.5px solid var(--ink);
         padding-bottom:2mm; margin-bottom:4mm; }
.phead .brand { font-family:'Playfair Display', Georgia, serif; font-weight:700; font-size:13pt; letter-spacing:.3px; }
.phead .cats { font-size:8.5pt; color:var(--muted); text-transform:uppercase; letter-spacing:1px; }
.pfoot { position:absolute; bottom:5mm; left:10mm; right:10mm; display:flex; justify-content:space-between;
         font-size:7.5pt; color:var(--muted); }
.pfoot .pn { font-weight:700; color:var(--ink); font-size:9pt; }
.grid { flex:1; display:grid; grid-template-columns:1fr 1fr; grid-template-rows:1fr 1fr; gap:4mm; min-height:0; }
.card { border:1px solid var(--line); border-radius:3mm; overflow:hidden; display:flex; flex-direction:column; background:#fff; min-height:0; }
.photo { height:68mm; flex:none; background-size:cover; background-position:center; position:relative; background-color:#eee; }
.photo.empty { display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; color:#999; font-size:9pt;
               border-bottom:1px dashed #ccc; background:repeating-linear-gradient(45deg,#f6f6f6 0 6px,#efefef 6px 12px); }
.ref { position:absolute; left:2mm; bottom:2mm; background:rgba(0,0,0,.55); color:#fff; font-size:6pt; padding:.6mm 1.6mm;
       border-radius:1mm; letter-spacing:.3px; text-transform:uppercase; }
.body { padding:3mm 3.5mm 3mm; display:flex; flex-direction:column; flex:1; min-height:0; }
.meta { display:flex; gap:2mm; align-items:center; font-size:7pt; }
.cat { background:var(--chip); color:var(--accent); padding:.5mm 2mm; border-radius:5mm; text-transform:uppercase; letter-spacing:.8px; font-weight:700; }
.sno { color:var(--muted); font-weight:600; }
.verify { margin-left:auto; color:#8a6d00; background:#fff3c4; padding:.3mm 1.6mm; border-radius:1mm; font-weight:600; }
h2 { font-family:'Playfair Display', Georgia, serif; font-size:15pt; margin:1.4mm 0 .8mm; line-height:1.1; }
.desc { font-size:7.5pt; color:var(--muted); margin:0 0 1mm; font-style:italic; }
h3 { font-size:6.8pt; text-transform:uppercase; letter-spacing:1px; color:var(--muted); margin:1.2mm 0 .8mm; }
ul.ing { margin:0; padding-left:3.6mm; font-size:8.8pt; line-height:1.3; }
ul.ing.two { columns:2; column-gap:4mm; }
ul.ing li { break-inside:avoid; margin-bottom:.3mm; }
ul.ing li::marker { color:var(--accent); }
.plate { margin-top:auto; font-size:8pt; line-height:1.3; background:#faf6f1; border-left:2px solid var(--accent);
         padding:1.4mm 2mm; border-radius:0 1mm 1mm 0; }
.plate b { color:var(--accent); text-transform:uppercase; font-size:6.6pt; letter-spacing:.8px; margin-right:1mm; }
/* cover & index */
.cover { justify-content:center; align-items:flex-start; padding:30mm 22mm; }
.cover .kicker { font-size:10pt; letter-spacing:3px; text-transform:uppercase; color:var(--accent); font-weight:700; }
.cover h1 { font-family:'Playfair Display', Georgia, serif; font-size:44pt; line-height:1; margin:4mm 0 6mm; }
.cover p { font-size:11pt; color:var(--muted); max-width:130mm; line-height:1.5; }
.cover .rule { width:30mm; height:3px; background:var(--accent); margin:8mm 0; }
.index h1 { font-family:'Playfair Display', Georgia, serif; font-size:22pt; margin:0 0 5mm; }
.toc { columns:2; column-gap:10mm; font-size:9pt; }
.toc .grp { break-inside:avoid; margin-bottom:3.5mm; }
.toc .gh { display:flex; justify-content:space-between; font-weight:700; border-bottom:1px solid var(--line); padding-bottom:.6mm; margin-bottom:.8mm; }
.toc .gh span:last-child { color:var(--accent); }
.toc .it { display:flex; justify-content:space-between; color:var(--muted); font-size:8pt; line-height:1.45; }
.credits { font-size:6pt; color:var(--muted); line-height:1.35; columns:3; column-gap:5mm; }
.credits div { break-inside:avoid; }
"""


def page(inner, n, cats, total):
    return f"""<section class="page">
  <div class="phead"><span class="brand">TCM · Kitchen Recipe Book</span><span class="cats">{e(' · '.join(cats))}</span></div>
  {inner}
  <div class="pfoot"><span>Main ingredients &amp; plating reference · for kitchen use</span><span class="pn">Page {n} / {total}</span></div>
</section>"""


def build():
    menu, credits = load_menu(), load_credits()
    chunks = [menu[i:i + PER_PAGE] for i in range(0, len(menu), PER_PAGE)]
    # page numbering: 1 cover, 2..(1+index_pages) index, then recipe pages, then credits
    index_pages = 2
    first = 2 + index_pages
    total = first + len(chunks) - 1 + (1 if credits else 0)
    item_page = {it["sno"]: first + i for i, ch in enumerate(chunks) for it in ch}

    out = []
    out.append(f"""<section class="page cover">
  <div class="kicker">Kitchen reference</div><h1>TCM<br>Recipe Book</h1><div class="rule"></div>
  <p>{len(menu)} menu items · main ingredients and final plating for every dish. Keep at the pass for quick reference.</p>
  <p style="font-size:9pt">Photos marked <b>Reference photo</b> are stand-ins showing the style of the dish — replace them with photos of our own plates
  by saving a picture as <code>images/&lt;item no&gt;.jpg</code> and rebuilding. Items tagged <b>✓ verify</b> have names open to interpretation; please check them.</p>
</section>""")

    groups = {}
    for it in menu:
        groups.setdefault(it["cat"], []).append(it)
    glist = list(groups.items())
    half, cnt, split = sum(len(v) + 2 for _, v in glist) / 2, 0, len(glist)
    for gi, (_, v) in enumerate(glist):
        cnt += len(v) + 2
        if cnt >= half: split = gi + 1; break
    tocs = ["".join(
        f'<div class="grp"><div class="gh"><span>{e(c)}</span><span>p. {item_page[its[0]["sno"]]}</span></div>'
        + "".join(f'<div class="it"><span>{it["sno"]} · {e(it["name"])}</span><span>{item_page[it["sno"]]}</span></div>' for it in its)
        + "</div>" for c, its in part) for part in (glist[:split], glist[split:])]
    for k, toc in enumerate(tocs):
        out.append(page(f'<div class="index"><h1>Contents{" (contd.)" if k else ""}</h1><div class="toc">{toc}</div></div>', 2 + k, ["Index"], total))

    for i, ch in enumerate(chunks):
        cats = list(dict.fromkeys(it["cat"] for it in ch))
        out.append(page(f'<div class="grid">{"".join(card(it, credits) for it in ch)}</div>', first + i, cats, total))

    if credits:
        rows = "".join(f'<div><b>#{k}</b> {e((v.get("artist") or "unknown")[:40])}, {e(v.get("lic") or "")}</div>'
                       for k, v in sorted(credits.items(), key=lambda kv: int(kv[0])))
        out.append(page(f'<div class="index"><h1>Photo credits</h1><p style="font-size:8pt;color:#6b625b">Reference photos from Wikimedia Commons, used under their stated licences.</p><div class="credits">{rows}</div></div>',
                        total, ["Credits"], total))

    doc = f"""<!doctype html><html><head><meta charset="utf-8"><title>TCM Recipe Book</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=Playfair+Display:wght@700&display=swap" rel="stylesheet">
<style>{CSS}</style></head><body>{''.join(out)}</body></html>"""
    html_path = os.path.join(BOOK, "TCM Recipe Book.html")
    open(html_path, "w").write(doc)

    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page()
        pg.goto("file://" + html_path, wait_until="networkidle")
        pg.pdf(path=os.path.join(ROOT, "TCM Recipe Book.pdf"), format="A4", print_background=True, prefer_css_page_size=True)
        # overflow check: any card whose body content is clipped
        bad = pg.evaluate("""() => [...document.querySelectorAll('.card')].filter(c => {
            const b = c.querySelector('.body'); return b.scrollHeight > b.clientHeight + 1; }).map(c => c.querySelector('h2').textContent)""")
        b.close()
    print(f"pages: {total}; overflowing cards: {bad or 'none'}")


if __name__ == "__main__":
    build()
