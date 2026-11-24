import os, math
import pandas as pd
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image

BASE = "/home/claude/spacex"
TEMPLATE = "/mnt/user-data/uploads/ds-capstone-template-coursera.pptx"
FIG = f"{BASE}/outputs/figures"; SQLD = f"{BASE}/outputs/sql"
OUT = f"{BASE}/deck/Data Science Capstone Project Report.pptx"
os.makedirs(f"{BASE}/deck", exist_ok=True)

TITLE_C = RGBColor(0x00, 0x54, 0x93); ACC = RGBColor(0x00, 0x70, 0xC0)
TEXT = RGBColor(0x26, 0x26, 0x26); GREY = RGBColor(0x59, 0x59, 0x59)
GREEN = RGBColor(0x3E, 0x8E, 0x62); RED = RGBColor(0xC8, 0x44, 0x55); ORANGE = RGBColor(0xD9, 0x82, 0x2B)
LIGHT = RGBColor(0xEA, 0xF2, 0xFA); WHITE = RGBColor(0xFF, 0xFF, 0xFF); MID = ACC
TITLE_FONT = "IBM Plex Mono SemiBold"
GITHUB = "https://github.com/[your-username]/spacex-falcon9-capstone"
LEFT, CW = 770011, 10515600

# ---------- data for text ----------
df = pd.read_csv(f"{BASE}/data/spacex_clean.csv", parse_dates=["Date"])
df["Year"] = df.Date.dt.year
mod = df[df.Landing_class.notna()].copy()
N, NM = len(df), len(mod)
succ, fail = int(mod.Landing_class.sum()), int((mod.Landing_class == 0).sum())
rate_all = mod.Landing_class.mean()
pre = mod[mod.Year < 2016]; post = mod[mod.Year >= 2016]
sites = df.groupby("Launch_Site").agg(launches=("Date", "size"), att=("Landing_class", "count"), ok=("Landing_class", "sum"))
sites["rate"] = sites.ok / sites.att
mc = pd.read_csv(f"{BASE}/outputs/model_comparison.csv")
best = mc.iloc[0]
sql = {i: pd.read_csv(f"{SQLD}/task{i}.csv") for i in range(1, 11)}

# ---------- helpers ----------
def delete_slide(prs, slide):
    lst = prs.slides._sldIdLst
    for sld in list(lst):
        if sld.id == slide.slide_id:
            rId = sld.rId; lst.remove(sld); prs.part.drop_rel(rId); return

def title(slide, text, size=30):
    tb = slide.shapes.add_textbox(Emu(LEFT), Emu(538650), Emu(CW), Emu(640000))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.name = TITLE_FONT; r.font.color.rgb = TITLE_C
    return tb

def text(slide, x, y, w, h, paras, size=15, color=TEXT, fill=None, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    if fill is not None:
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
        shp.adjustments[0] = 0.06; shp.fill.solid(); shp.fill.fore_color.rgb = fill; shp.line.fill.background()
    else:
        shp = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = shp.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(100000); tf.margin_top = tf.margin_bottom = Emu(60000)
    for i, p_ in enumerate(paras):
        t, o = (p_, {}) if isinstance(p_, str) else p_
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = o.get("align", align); para.space_after = Pt(o.get("space", 6))
        r = para.add_run(); r.text = t
        r.font.size = Pt(o.get("size", size)); r.font.bold = o.get("bold", False)
        r.font.color.rgb = o.get("color", color)
    return shp

def heading(t, sz=17):
    return (t, {"bold": True, "size": sz, "color": TITLE_C, "space": 8})

def bullets(items, size=15):
    return [("•  " + i, {"size": size, "space": 7}) for i in items]

def img(slide, path, x, y, max_w, max_h, caption=None):
    w0, h0 = Image.open(path).size
    scale = min(max_w / w0, max_h / h0)
    w, h = int(w0 * scale), int(h0 * scale)
    px = x + int((max_w - w) / 2)
    slide.shapes.add_picture(path, Emu(px), Emu(y), width=Emu(w), height=Emu(h))
    if caption:
        text(slide, x, y + h + 40000, max_w, 300000, [(caption, {"size": 11, "color": GREY})], size=11)
    return h

def table(slide, x, y, w, data, col_w=None, size=12, row_h=330000):
    rows, cols = len(data), len(data[0])
    gt = slide.shapes.add_table(rows, cols, Emu(x), Emu(y), Emu(w), Emu(row_h * rows))
    t = gt.table
    if col_w:
        tot = sum(col_w)
        for i, cw in enumerate(col_w): t.columns[i].width = Emu(int(w * cw / tot))
    for r_i, rv in enumerate(data):
        t.rows[r_i].height = Emu(row_h)
        for c_i, v in enumerate(rv):
            cell = t.cell(r_i, c_i); cell.text = str(v)
            cell.margin_left = cell.margin_right = Emu(60000); cell.margin_top = cell.margin_bottom = Emu(25000)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]; p.font.size = Pt(size)
            cell.fill.solid()
            if r_i == 0:
                cell.fill.fore_color.rgb = TITLE_C; p.font.bold = True; p.font.color.rgb = WHITE
            else:
                cell.fill.fore_color.rgb = LIGHT if r_i % 2 == 0 else WHITE; p.font.color.rgb = TEXT
    return gt

KIND = {"data": (LIGHT, TITLE_C), "model": (ACC, WHITE), "eval": (GREEN, WHITE), "output": (ORANGE, WHITE)}
def arrow(c):
    ln = c.line._get_or_add_ln(); t = etree.SubElement(ln, qn("a:tailEnd")); t.set("type", "triangle")

def flow(slide, steps, per_row=4, top=1900000, bottom=5700000, left=770011, right=11286000):
    n = len(steps); rows = math.ceil(n / per_row); gap = 420000
    bw = int((right - left - (per_row - 1) * gap) / per_row)
    pitch = (bottom - top) / rows; bh = int(min(950000, pitch * 0.6))
    pos = []
    for r in range(rows):
        items = steps[r * per_row:(r + 1) * per_row]
        y = int(top + r * pitch + (pitch - bh) / 2)
        for c, lab in enumerate(items):
            col = c if r % 2 == 0 else (per_row - 1 - c)
            pos.append((int(left + col * (bw + gap)), y, lab))
    for x, y, (lab, kind) in pos:
        fill, tc = KIND[kind]
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(x), Emu(y), Emu(bw), Emu(bh))
        shp.adjustments[0] = 0.15; shp.fill.solid(); shp.fill.fore_color.rgb = fill
        shp.line.color.rgb = ACC if kind == "data" else fill
        tf = shp.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for li, line in enumerate(lab.split("\n")):
            p = tf.paragraphs[0] if li == 0 else tf.add_paragraph(); p.alignment = PP_ALIGN.CENTER
            r = p.add_run(); r.text = line; r.font.size = Pt(13 if li == 0 else 11)
            r.font.bold = (li == 0); r.font.color.rgb = tc
    for i in range(len(pos) - 1):
        x1, y1, _ = pos[i]; x2, y2, _ = pos[i + 1]
        if abs(y1 - y2) < 10:
            if x2 > x1: c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x1 + bw), Emu(y1 + bh // 2), Emu(x2), Emu(y2 + bh // 2))
            else: c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x1), Emu(y1 + bh // 2), Emu(x2 + bw), Emu(y2 + bh // 2))
        else:
            c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x1 + bw // 2), Emu(y1 + bh), Emu(x2 + bw // 2), Emu(y2))
        c.line.color.rgb = GREY; c.line.width = Pt(1.75); arrow(c)

def notes(slide, t): slide.notes_slide.notes_text_frame.text = t

# ---------- open template; keep cover and section slides ----------
prs = Presentation(TEMPLATE)
T = list(prs.slides)
cover, sec1, sec2, sec3, sec4, sec5 = T[0], T[4], T[16], T[33], T[37], T[41]
blank = [l for l in prs.slide_layouts if l.name == "Blank"][0]

# cover
for sh in cover.shapes:
    if sh.name == "TextBox 5":
        tf = sh.text_frame
        tf.paragraphs[0].runs[0].text = "Muhammad Enzi Muzakki"
        for p_ in list(tf.paragraphs)[1:]: p_._p.getparent().remove(p_._p)
        p2 = tf.add_paragraph(); r2 = p2.add_run(); r2.text = "November 24, 2026"
        sh.width = Emu(4300000); sh.top = Emu(4450000)
        for p_ in tf.paragraphs:
            for r_ in p_.runs: r_.font.size = Pt(20)

def section(sl, label):
    for sh in sl.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip().startswith("Section"):
            sh.text_frame.paragraphs[0].runs[0].text = label
            for r_ in sh.text_frame.paragraphs[0].runs[1:]: r_.text = ""
            sh.width = Emu(7000000)
    return sl


def new(): return prs.slides.add_slide(blank)
slides = {}

# ===== Outline =====
s = new(); title(s, "Outline")
text(s, LEFT + 150000, 1750000, 9800000, 4300000, [
    ("Executive summary and introduction", {"size": 20, "space": 14}),
    ("Section 1: Methodology (data collection, wrangling, EDA and predictive design)", {"size": 20, "space": 14}),
    ("Section 2: Exploratory data analysis (visualisation and SQL results)", {"size": 20, "space": 14}),
    ("Section 3: Launch sites analysis (Folium map)", {"size": 20, "space": 14}),
    ("Section 4: Dashboard with Plotly Dash", {"size": 20, "space": 14}),
    ("Section 5: Predictive analysis (classification) and conclusions", {"size": 20, "space": 14}),
])
slides["outline"] = s

# ===== Executive summary =====
s = new(); title(s, "Executive Summary")
text(s, LEFT, 1400000, 5100000, 4500000, [
    heading("What we did"),
    *bullets([
        "Documented the SpaceX REST API and Wikipedia scraping methods; analysed the processed dataset of 101 launches (2010–2020).",
        "Cleaned the data, fixed inconsistent category labels, and labelled each landing as success or failure.",
        "Explored the data with visualisations and 10 SQL queries, built an interactive Folium map and a Plotly Dash dashboard, and trained four classifiers.",
    ], 16)], size=16)
text(s, 6150000, 1400000, 5300000, 4500000, [
    heading("What we found"),
    *bullets([
        f"{rate_all*100:.0f}% of attempted first-stage landings succeeded ({succ} of {NM}).",
        f"Landing success was {pre.Landing_class.mean()*100:.0f}% before 2016 and {post.Landing_class.mean()*100:.0f}% from 2016 onward, so time is the strongest signal.",
        f"Booster generation matters: F9 v1.0 and v1.1 landed 0 of 5, while F9 B5 landed 39 of 42.",
        f"Logistic regression was the best classifier by cross-validation ({best.cv_accuracy*100:.1f}%), with {best.test_accuracy*100:.0f}% test accuracy.",
    ], 16)], size=16, fill=LIGHT)
notes(s, "Executive summary for a senior audience: what we did, and the four results that matter most. Detail follows in the methodology and results sections.")
slides["exec"] = s

# ===== Introduction =====
s = new(); title(s, "Introduction")
text(s, LEFT, 1400000, 5100000, 4700000, [
    heading("Project background and context"),
    *bullets([
        "SpaceX is the only private company to return a spacecraft from low-Earth orbit.",
        "Falcon 9 launches are advertised at about $62 million, against $165 million or more for other providers.",
        "Much of the saving comes from reusing the first stage. If we can predict whether it will land, we can estimate the cost of a launch.",
        "An alternate bidder can use the same reasoning to price its own launch offers.",
    ], 14)], size=14)
text(s, 6150000, 1400000, 5300000, 4700000, [
    heading("Problems we want to answer"),
    *bullets([
        "Which factors (launch site, orbit, booster version, payload mass, year) relate to a successful landing?",
        "How have launch volume and landing success changed over time?",
        "Can a classifier predict landing success from information known before launch?",
        "Which visual tools (map, dashboard) help a stakeholder act on these results?",
    ], 14)], size=14, fill=LIGHT)
notes(s, "Background from the course materials: SpaceX advertises Falcon 9 at about 62 million dollars, and competitors cost upward of 165 million. The business question is whether the first stage will land.")
slides["intro"] = s

# ===== Methodology overview =====
s = new(); title(s, "Methodology Overview")
flow(s, [
    ("Data collection\nSpaceX REST API + Wikipedia", "data"),
    ("Data wrangling\nclean, trim, label outcomes", "data"),
    ("EDA: visualisation\nand SQL (10 queries)", "model"),
    ("Interactive analytics\nFolium map + Plotly Dash", "model"),
    ("Predictive analysis\nclassification, cross-validated", "eval"),
    ("Conclusions\nrecommendations and next steps", "output"),
], per_row=3, top=1850000, bottom=5350000)
text(s, LEFT, 5550000, CW, 500000, [(f"Repository: {GITHUB}", {"size": 12, "color": GREY})], size=12)
slides["meth"] = s

# ===== Data collection =====
s = new(); title(s, "Data Collection: SpaceX REST API and Web Scraping")
text(s, LEFT, 1350000, 4600000, 4500000, [
    heading("Key steps"),
    *bullets([
        "Called SpaceX REST API v4 endpoints for rockets, launchpads, payloads, cores and launches using requests.get().",
        "Converted the JSON responses to a DataFrame with json_normalize().",
        "Scraped Wikipedia's 'List of Falcon 9 and Falcon Heavy launches' with BeautifulSoup, reading the table elements.",
        "Combined both sources and kept Falcon 9 records for analysis.",
    ], 13),
    ("Note: the API and scraping notebooks could not be re-run in this environment, so the analysis uses the processed Spacex.csv they produce.", {"size": 11, "color": GREY, "space": 4}),
    ("GitHub: " + GITHUB, {"size": 11, "color": ACC}),
], size=13)
flow(s, [
    ("SpaceX REST API v4\nrockets, launchpads, payloads", "data"),
    ("requests.get()\nJSON responses", "model"),
    ("json_normalize()\nto DataFrame", "model"),
    ("Wikipedia launch list\nBeautifulSoup tables", "data"),
    ("Combine and filter\nFalcon 9 records", "eval"),
    ("Spacex.csv\nprocessed dataset", "output"),
], per_row=2, top=1500000, bottom=5600000, left=5350000, right=11286000)
notes(s, "The API notebook queries SpaceX REST v4 for rockets, launchpads, payloads, cores and launches, then flattens the JSON. "
         "The web-scraping notebook reads the Falcon 9 launch-list table from Wikipedia. The processed result is Spacex.csv, which this analysis uses.")
slides["collect"] = s

# ===== Data wrangling =====
s = new(); title(s, "Data Wrangling")
flow(s, [
    ("Load Spacex.csv\n101 rows, 10 columns", "data"),
    ("Trim text fields\nfix 'Success ' labels", "data"),
    ("Drop missing dates\n0 rows removed", "model"),
    ("Parse Date, Year\nbooster family from version", "model"),
    ("Label landing_class\n1 success, 0 failure", "eval"),
    ("Save spacex_clean.csv\n71 modelling rows", "output"),
], per_row=3, top=1450000, bottom=4950000)
text(s, LEFT, 5100000, CW, 1200000, [
    *bullets([
        "Labelling: success = any 'Success' outcome (ground pad or drone ship); failure = any 'Failure' outcome. 'No attempt', 'Precluded' and 'Controlled/Uncontrolled (ocean)' are not clear outcomes and are kept for EDA but excluded from modelling.",
        f"Result: {NM} landings with a clear outcome ({succ} successes, {fail} failures).  GitHub: {GITHUB}",
    ], 13)], size=13)
notes(s, "Raw labels contained trailing spaces that would split one category into two, so every text field is trimmed. The label definition keeps modelling honest by excluding launches with no clear landing outcome.")
slides["wrangle"] = s

# ===== EDA methodology =====
s = new(); title(s, "EDA and Interactive Visual Analytics Methodology")
text(s, LEFT, 1400000, 5100000, 4500000, [
    heading("Exploratory data analysis"),
    *bullets([
        "Visualisation with Matplotlib and Seaborn: yearly trend, site, orbit, booster family and payload.",
        "SQL with SQLite: ten lab tasks covering filters, aggregates, subqueries, date parsing and ranking.",
        "Each chart answers one question about landing success.",
    ], 14)], size=14)
text(s, 6150000, 1400000, 5300000, 4500000, [
    heading("Interactive visual analytics"),
    *bullets([
        "Folium map: launch sites as markers, circles sized by launch count, popups with success rates.",
        "Plotly Dash dashboard: site dropdown, outcome pie chart, payload range slider and scatter plot.",
        "Both tools let a stakeholder filter by site or payload without touching the code.",
    ], 14)], size=14, fill=LIGHT)
text(s, LEFT, 5700000, CW, 350000, [(f"Notebooks: 01_SQL_Tasks.ipynb and 02_EDA_Map_Prediction.ipynb  ·  GitHub: {GITHUB}", {"size": 11, "color": GREY})], size=11)
notes(s, "EDA uses two complementary routes: visual charts for patterns and SQL for precise answers. Interactive tools are built on the same cleaned data so the numbers match.")
slides["eda_meth"] = s

# ===== Predictive methodology =====
s = new(); title(s, "Predictive Analysis Methodology")
flow(s, [
    ("Features\nsite, orbit, booster, payload, year", "data"),
    ("Pipeline\none-hot encode, scale payload and year", "model"),
    ("Stratified 80/20 split\n56 train, 15 test", "data"),
    ("5-fold CV grid search\n4 model families", "model"),
    ("Select by CV accuracy\nnot by test accuracy", "eval"),
    ("Test once and report\nconfusion matrix, importance", "output"),
], per_row=3, top=1450000, bottom=4950000)
text(s, LEFT, 5100000, CW, 1200000, [
    *bullets([
        "Models: logistic regression, SVM, decision tree and k-nearest neighbours, each with its own hyper-parameter grid.",
        f"Choosing by cross-validation avoids tuning on the test set. GitHub: {GITHUB}",
    ], 13)], size=13)
notes(s, "A stratified split keeps the success share the same in training and test. Every model is tuned with 5-fold cross-validation on the training set only. The test set is used once, to report the final accuracy.")
slides["pred_meth"] = s

# ===== Results divider handled by section slide =====
# ===== EDA visual results =====
def is_wide(path):
    w0, h0 = Image.open(path).size
    return w0 / h0 >= 1.6

def img_slide(key, ttl, path, expl, note_txt, cap=None):
    s = new(); title(s, ttl, size=26)
    if is_wide(path):
        h = img(s, path, LEFT, 1250000, CW, 3700000, caption=cap)
        ty = 1250000 + h + (400000 if cap else 150000)
        paras = [(e, {"size": 14, "space": 8}) if isinstance(e, str) else e for e in expl]
        text(s, LEFT, ty, CW, 5900000 - ty, paras, size=14, fill=LIGHT)
    else:
        img(s, path, LEFT, 1300000, 6900000, 4650000, caption=cap)
        text(s, 7900000, 1350000, 3400000, 4600000, [(e, {"size": 13, "space": 9}) if isinstance(e, str) else e for e in expl], size=13, fill=LIGHT)
    notes(s, note_txt); slides[key] = s
    return s

yr = mod.groupby("Year").agg(att=("Landing_class", "count"), ok=("Landing_class", "sum"))
yr["rate"] = yr.ok / yr.att
img_slide("yearly", "Launch Success Yearly Trend", f"{FIG}/yearly_trend.png",
    [heading("What it shows"),
     f"Landing success rose from 0% in 2010 to 100% in 2017 and 2019, and has stayed at 91% or higher since 2017.",
     f"Launches grew from 2 in 2010 to 23 in 2020, so experience and scale grew together.",
     ("Why this chart: a line chart shows a trend over time, and the bars show how much data each year has."), ],
    "Yearly landing success rate (line) and launch count (bars). Years with no attempted landings are not plotted.")

img_slide("flight_site", "Flight Number vs. Launch Site", f"{FIG}/flight_vs_site.png",
    [heading("What it shows"),
     "Each dot is one launch, ordered chronologically across the four sites.",
     "CCAFS SLC-40 hosts the most launches (34). VAFB SLC-4E hosted only polar, sun-synchronous and LEO missions.",
     "Why this chart: it shows how each site's use changed over the flight sequence."],
    "Flight number is the chronological index of each launch in the dataset.")

img_slide("payload_site", "Payload vs. Launch Site", f"{FIG}/payload_vs_site.png",
    [heading("What it shows"),
     "Green is a successful landing, red is a failure, and each dot is one launch.",
     "Failures appear at many payload sizes (0 to 15,600 kg), so payload mass alone does not separate outcomes.",
     "Why this chart: it checks whether heavier payloads at a site fail more often."],
    "Payload mass against launch site, coloured by landing outcome. Only launches with a clear outcome are shown.")

so = mod.groupby("Orbit").agg(n=("Landing_class", "size"), r=("Landing_class", "mean")).sort_values("r", ascending=False)
img_slide("orbit_rate", "Success Rate vs. Orbit Type", f"{FIG}/success_by_orbit.png",
    [heading("What it shows"),
     f"Orbits with a clear landing record reach 78% to 100% success.",
     f"GTO missions are the most common ({int(so.loc['GTO','n'])} with a clear outcome) at {so.loc['GTO','r']*100:.0f}% success.",
     "Why this chart: a bar chart makes it easy to compare the success rate across orbit types, with sample size shown."],
    "Landing success rate by orbit type. The number in brackets is the count of launches with a clear outcome.")

img_slide("flight_orbit", "Flight Number vs. Orbit Type", f"{FIG}/flight_vs_orbit.png",
    [heading("What it shows"),
     "Each dot is one launch, placed by its flight number and orbit type.",
     "The first 20 launches are mostly LEO (ISS) and GTO (15 of 20); the last 20 are mostly LEO (12 of 20).",
     "Why this chart: it shows whether the mix of missions changed over time."],
    "Flight number against orbit type. Colours distinguish orbits.")

img_slide("payload_orbit", "Payload vs. Orbit Type", f"{FIG}/payload_vs_orbit.png",
    [heading("What it shows"),
     f"Mean payload is highest for LEO ({df[df.Orbit=='LEO'].Payload_kg.mean():,.0f} kg) and lowest for HEO and SSO.",
     "Heavier payloads go to LEO and Polar LEO; GTO carries mid-sized loads (about 5,000 kg); LEO (ISS) and SSO are lighter.",
     "Why this chart: payload and orbit are linked, so both must be considered in the model."],
    "Payload mass against orbit type, coloured by landing outcome.")

# Booster family and an innovative insight
bf = mod.groupby("Booster_family").agg(n=("Landing_class", "size"), ok=("Landing_class", "sum"))
bf = bf.reindex(["F9 v1.0", "F9 v1.1", "F9 FT", "F9 B4", "F9 B5"])
img_slide("booster", "Landing Success by Booster Generation", f"{FIG}/booster_family_success.png",
    [heading("What it shows"),
     f"F9 v1.0 and v1.1 landed 0 of {int(bf.loc['F9 v1.0','n'] + bf.loc['F9 v1.1','n'])} attempts.",
     f"F9 FT landed {int(bf.loc['F9 FT','ok'])} of {int(bf.loc['F9 FT','n'])}, B4 {int(bf.loc['F9 B4','ok'])} of {int(bf.loc['F9 B4','n'])}, and B5 {int(bf.loc['F9 B5','ok'])} of {int(bf.loc['F9 B5','n'])}.",
     "Why this chart: it compares generations, which reflect design changes over time."],
    "Landing success by booster family, with the number of attempts in brackets.")

s = new(); title(s, "Innovative Insight: Time Explains Most of the Variation", size=26)
text(s, LEFT, 1350000, 5300000, 4700000, [
    heading("Before and after 2016"),
    *bullets([
        f"Before 2016: {int(pre.Landing_class.sum())} of {len(pre)} landings succeeded ({pre.Landing_class.mean()*100:.0f}%).",
        f"From 2016: {int(post.Landing_class.sum())} of {len(post)} succeeded ({post.Landing_class.mean()*100:.0f}%).",
        f"CCAFS LC-40 hosted 19 of its 26 launches before 2016, so its lower rate ({sites.loc['CCAFS LC-40','rate']*100:.0f}%) reflects its early history.",
        f"Other sites together reached {mod[mod.Launch_Site!='CCAFS LC-40'].Landing_class.mean()*100:.0f}% success.",
    ], 14)], size=14)
table(s, 6400000, 1500000, 5100000, [
    ["Launch site", "Launches", "Landings", "Success %"],
    *[[k, int(r.launches), f"{int(r.ok)}/{int(r.att)}", f"{r.rate*100:.0f}%"] for k, r in sites.sort_values("rate").iterrows()],
], col_w=[2.2, 1.2, 1.2, 1.2], size=12, row_h=380000)
text(s, 6400000, 3750000, 5100000, 1500000, [
    ("Why this matters: a model that ignores time would confuse early learning with site quality. The year feature is the top predictor in Section 4.", {"size": 13})],
    size=13, fill=LIGHT)
notes(s, "Innovative insight: the apparent site differences are largely a time effect. Early launches, mostly from CCAFS LC-40, happened while the booster design was still changing.")
slides["insight"] = s

# ===== SQL results =====
def sql_table_slide(key, ttl, rows, expl_lines, note_txt, col_w=None, size=12):
    s = new(); title(s, ttl, size=26)
    table(s, LEFT, 1350000, CW, rows, col_w=col_w, size=size, row_h=360000)
    text(s, LEFT, 1350000 + 360000 * len(rows) + 150000, CW, 1200000,
         [(e, {"size": 13, "space": 6}) for e in expl_lines], size=13)
    notes(s, note_txt); slides[key] = s

sql_table_slide("sql1", "EDA with SQL: Tasks 1–5", [
    ["Task", "SQL result", "Answer"],
    ["1  Unique launch sites", f"{len(sql[1])} distinct sites", ", ".join(sql[1].Launch_Site)],
    ["2  Records starting 'CCA'", "5 records", "First five launches from CCAFS LC-40, 2010–2013"],
    ["3  NASA (CRS) total payload", f"{sql[3].iloc[0,0]:,.0f} kg", "Sum of PAYLOAD_MASS__KG_ where Customer = 'NASA (CRS)'"],
    ["4  Average payload, F9 v1.1", f"{sql[4].iloc[0,0]:,.1f} kg", "Mean across the 15 'F9 v1.1' flights"],
    ["5  First ground-pad success", str(sql[5].iloc[0,0]), "MIN(Date) where outcome = 'Success (ground pad)'"],
], [
    "SQL on the cleaned SpaceX table; all text fields trimmed first. Queries and results: notebooks/01_SQL_Tasks.ipynb.",
], "Results of the first five SQL tasks. Task 2 uses LIKE 'CCA%'; task 5 uses MIN on the date text.",
    col_w=[3, 2, 5.5], size=12)

sql_table_slide("sql2", "EDA with SQL: Tasks 6–10", [
    ["Task", "SQL result", "Answer"],
    ["6  Drone-ship success, payload 4,000–6,000 kg", f"{len(sql[6])} boosters", ", ".join(sql[6].Booster_Version.str.strip())],
    ["7  Mission outcomes", "3 categories", "Success 99; Success (payload status unclear) 1; Failure (in flight) 1"],
    ["8  Maximum payload (15,600 kg)", f"{len(sql[8])} boosters", "Subquery on MAX(PAYLOAD_MASS__KG_); all are F9 B5 variants"],
    ["9  2015 failed drone-ship landings", f"{len(sql[9])} records", "January (F9 v1.1 B1012) and April (F9 v1.1 B1015), both CCAFS LC-40"],
    ["10 Landing outcomes, 2010-06-04 to 2017-03-20", f"{len(sql[10])} outcome types", "Ranked by count: No attempt 10, then Success (drone ship) 5 and Failure (drone ship) 5"],
], [
    "Task 9 uses substr(Date, 6, 2) for the month, since SQLite has no month-name function.",
], "Results of SQL tasks 6 to 10. Task 8 uses a subquery with MAX. Task 10 counts landing outcomes in the date window.",
    col_w=[3.2, 1.8, 5.5], size=11)

# ===== Interactive: Folium (divider handled) =====
img_slide("folium1", "Folium Map: All Launch Sites", f"{FIG}/map_all_sites.png",
    [heading("What it shows"),
     "Four launch sites: three in Florida (CCAFS LC-40, CCAFS SLC-40, KSC LC-39A) and VAFB SLC-4E in California.",
     "Circles scale with the number of launches at each site.",
     "Why: a map shows geographic spread, and the zoomed panel separates the close Cape Canaveral sites."],
    "Static snapshot of the Folium map (outputs/launch_sites_map.html). Coordinates are approximate published locations.",
    cap="Global view (left) and Cape Canaveral zoom (right). Site coordinates are approximate.")

img_slide("folium2", "Folium Map: Landing Success by Site", f"{FIG}/map_outcome_colours.png",
    [heading("What it shows"),
     f"Green is a site with at least 80% landing success: KSC LC-39A ({sites.loc['KSC LC-39A','rate']*100:.0f}%), CCAFS SLC-40 ({sites.loc['CCAFS SLC-40','rate']*100:.0f}%) and VAFB ({sites.loc['VAFB SLC-4E','rate']*100:.0f}%).",
     f"Red is CCAFS LC-40 at {sites.loc['CCAFS LC-40','rate']*100:.0f}%.",
     "Why: colour turns the rate into an immediate signal for a stakeholder."],
    "Colour encodes the landing success rate for each site (green 80% or more, orange 60% or more, red below 60%).",
    cap="Marker colour is the landing success rate for each site.")

s = new(); title(s, "Folium Map: Proximity to Reference Point", size=26)
img(s, f"{FIG}/proximity_jsc.png", LEFT, 1300000, 6600000, 4300000)
text(s, 7650000, 1350000, 3700000, 4600000, [
    heading("What it shows"),
    "Great-circle distances from each launch site to NASA Johnson Space Center, the reference point used in the course lab.",
    "Distances are about 1,400 km for the Florida sites and about 2,460 km for VAFB.",
    ("Limitation: distances to coastline, railways and highways need external geodata that is not in the provided files, so they are not calculated here.", {"size": 12, "color": GREY}),
], size=13, fill=LIGHT)
notes(s, "This slide demonstrates the proximity calculation with a haversine distance. Coastline, railway and highway proximity would follow the same method once those datasets are added.")
slides["folium3"] = s

# ===== Dash dashboard =====
def dash_slide(key, ttl, path, expl, note_txt):
    img_slide(key, ttl, path, expl, note_txt)

ks = df[df.Launch_Site == "KSC LC-39A"].Landing_Outcome.value_counts()
dash_slide("dash1", "Dashboard: Landing Outcomes, All Sites", f"{FIG}/dash_pie_all.png",
    [heading("What it shows"),
     f"Outcome share across all {N} launches, including those with no landing attempt.",
     "Pie chart: shows the part of the whole that each outcome makes up.",
     "Filter: the site dropdown changes the pie to a single site."],
    "Dashboard view 1: the pie chart from the Dash app for all sites. The live app is in scripts/spacex_dash_app.py.")

dash_slide("dash2", "Dashboard: Highest Success Ratio, KSC LC-39A", f"{FIG}/dash_pie_ksc.png",
    [heading("What it shows"),
     f"KSC LC-39A has the highest landing success ratio of the four sites ({sites.loc['KSC LC-39A','rate']*100:.0f}% of attempted landings).",
     f"Its {int(sites.loc['KSC LC-39A','launches'])} launches include {int(ks.get('No attempt', 0))} with no landing attempt, so the pie includes them.",
     "Why: selecting this site shows whether the success share holds up when filtered."],
    "Dashboard view 2: KSC LC-39A selected in the dropdown.")

dash_slide("dash3", "Dashboard: Payload vs. Outcome with Range Slider", f"{FIG}/dash_payload_range.png",
    [heading("What it shows"),
     "The range slider limits payload mass; the scatter plot colours launches by outcome and marks orbit.",
     "Low payloads (0–5,000 kg) include both failures (F9 v1.0 and v1.1) and successes.",
     "Heavier payloads (5,000–10,000 kg) are almost all successes, mostly F9 B5."],
    "Dashboard view 3: two payload windows, showing how the range slider changes what is visible.")

# ===== Predictive results =====
s = new(); title(s, "Classification Accuracy by Model", size=26)
img(s, f"{FIG}/model_accuracy.png", LEFT, 1300000, 6500000, 4600000)
table(s, 7500000, 1450000, 3900000, [
    ["Model", "CV %", "Test %"],
    *[[r.model, f"{r.cv_accuracy*100:.1f}", f"{r.test_accuracy*100:.1f}"] for _, r in mc.iterrows()],
], col_w=[2.6, 1, 1], size=11, row_h=380000)
text(s, 7500000, 3700000, 3900000, 2000000, [
    (f"Best model: {best.model}, chosen by cross-validation ({best.cv_accuracy*100:.1f}%).", {"size": 13, "bold": True, "color": TITLE_C}),
    ("The test set has only 15 landings, so one error moves test accuracy by about 6.7 points. The cross-validation score is more reliable here.", {"size": 12, "space": 4}),
], size=13, fill=LIGHT)
notes(s, "All four models were tuned by 5-fold cross-validation on the training data. Logistic regression had the highest cross-validation accuracy. Test accuracies are close because the test set is tiny.")
slides["pred_acc"] = s

s = new(); title(s, "Confusion Matrix and Feature Importance", size=26)
img(s, f"{FIG}/confusion_matrix.png", LEFT, 1300000, 4700000, 3900000)
img(s, f"{FIG}/feature_importance.png", 5700000, 1300000, 5600000, 3500000)
text(s, LEFT, 5300000, CW, 900000, [
    ("Confusion matrix (test set, 15 landings): all 13 successes were predicted correctly. One of the two failures was predicted correctly and one was predicted as a success.", {"size": 12, "space": 4}),
    ("Feature importance (depth-3 tree): launch year is the strongest feature (0.42), followed by the FT booster (0.32) and payload mass (0.26).", {"size": 12}),
], size=12)
notes(s, "The confusion matrix shows the model misses no successful landings. Its one false alarm is a failure predicted as a success. Feature importance from a shallow tree shows year and booster generation carry most of the signal.")
slides["pred_cm"] = s

# ===== Conclusions =====
s = new(); title(s, "Conclusions")
text(s, LEFT, 1400000, CW, 4500000, [
    *bullets([
        f"Time is the strongest signal: landing success was {pre.Landing_class.mean()*100:.0f}% before 2016 and {post.Landing_class.mean()*100:.0f}% from 2016 onward.",
        "Booster generation matters more than payload: F9 v1.0 and v1.1 landed none of their attempts, while B5 landed 39 of 42.",
        "CCAFS LC-40's low success rate mostly reflects its early launches, not its location.",
        f"Logistic regression was the best classifier ({best.cv_accuracy*100:.1f}% cross-validated accuracy) and misses no successful landings on the test set.",
        "Next steps: add coastline, rail and highway distances; enlarge the dataset so the test set is more than 15 landings; and deploy the Dash dashboard for stakeholders.",
    ], 16)], size=16)
notes(s, "Summary for a senior audience. The recommendation: use the classifier and the dashboard as decision support, and collect more landings before relying on exact probabilities.")
slides["concl"] = s

# ===== Appendix =====
s = new(); title(s, "Appendix: Assets and Reproducibility")
text(s, LEFT, 1400000, CW, 4700000, [
    heading("Repository and notebooks"),
    *bullets([
        f"GitHub repository: {GITHUB}",
        "notebooks/01_SQL_Tasks.ipynb: ten SQL tasks on the cleaned table (executed).",
        "notebooks/02_EDA_Map_Prediction.ipynb: wrangling, EDA charts, Folium map, classifiers (executed).",
        "scripts/spacex_dash_app.py: Plotly Dash dashboard (run with python scripts/spacex_dash_app.py).",
        "data/Spacex.csv (source), data/spacex_clean.csv (cleaned), outputs/ (figures, SQL results, model comparison).",
    ], 14),
    ("Reproduce: run the notebooks in order, then start the dashboard. Dependencies: pandas, numpy, matplotlib, seaborn, scikit-learn, folium, plotly, dash.", {"size": 12, "color": GREY}),
], size=14)
notes(s, "Everything shown in this report is reproducible from the repository. Replace the GitHub placeholder with your repository address before submitting.")
slides["appendix"] = s

s = new(); title(s, "Appendix: Model Comparison and Site Performance", size=26)
img(s, f"{FIG}/success_by_site_and_orbit.png", LEFT, 1300000, 6200000, 4500000)
table(s, 7250000, 1450000, 4100000, [
    ["Model", "Best parameters"],
    ["Logistic regression", "C = 1"],
    ["SVM", "C = 10, gamma = scale"],
    ["K-nearest neighbours", "k = 3, uniform weights"],
    ["Decision tree", "max depth = 2"],
], col_w=[2.2, 2.6], size=11, row_h=360000)
text(s, 7250000, 3600000, 4100000, 1800000, [
    ("Hyper-parameters were chosen by 5-fold cross-validation on the 56-landing training set. Only the best setting per model is shown.", {"size": 12}),
], size=12, fill=LIGHT)
slides["appendix2"] = s

# ----- reorder, keep only chosen slides -----
order = [cover, slides["outline"], slides["exec"], slides["intro"], sec1, slides["meth"], slides["collect"], slides["wrangle"],
         slides["eda_meth"], slides["pred_meth"],
         sec2, slides["yearly"], slides["flight_site"], slides["payload_site"], slides["orbit_rate"], slides["flight_orbit"],
         slides["payload_orbit"], slides["booster"], slides["insight"], slides["sql1"], slides["sql2"],
         sec3, slides["folium1"], slides["folium2"], slides["folium3"],
         sec4, slides["dash1"], slides["dash2"], slides["dash3"],
         sec5, slides["pred_acc"], slides["pred_cm"], slides["concl"], slides["appendix"], slides["appendix2"]]
order_ids = [sl.slide_id for sl in order]
keep_ids = set(order_ids)
lst = prs.slides._sldIdLst
id_elems = {sld.id: sld for sld in list(lst)}
for sld in list(lst): lst.remove(sld)
for oid in order_ids:
    lst.append(id_elems[oid])
for sld_id, sld in id_elems.items():
    if sld_id not in keep_ids:
        prs.part.drop_rel(sld.rId)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
