import os, math
import numpy as np, pandas as pd
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.patches import Circle

BASE = "/home/claude/spacex"; FIG = f"{BASE}/outputs/figures"; os.makedirs(FIG, exist_ok=True)
sns.set_theme(style="whitegrid", context="talk")
BLUE, DBLUE, GREEN, RED, GREY = "#0070C0", "#005493", "#3E8E62", "#C84455", "#7F7F7F"

df = pd.read_csv(f"{BASE}/data/spacex_clean.csv", parse_dates=["Date"])
df["Flight_number"] = np.arange(1, len(df) + 1)
mod = df[df.Landing_class.notna()].copy()
mod["Outcome"] = mod.Landing_class.map({1.0: "Success", 0.0: "Failure"})

def save(fig, name):
    fig.tight_layout(); fig.savefig(f"{FIG}/{name}.png", dpi=170, facecolor="white"); plt.close(fig)

# --- Flight number vs launch site (all launches, outcome coloured where attempted)
fig, ax = plt.subplots(figsize=(11, 5.2))
sites = sorted(df.Launch_Site.unique())
for i, s in enumerate(sites):
    d = df[df.Launch_Site == s]
    ax.scatter(d.Flight_number, [i] * len(d), s=60, color=BLUE, alpha=0.6, edgecolor="white")
ax.set_yticks(range(len(sites))); ax.set_yticklabels(sites)
ax.set_xlabel("Flight number (chronological)"); ax.set_title("Flight number vs launch site")
save(fig, "flight_vs_site")

# --- Payload vs launch site
fig, ax = plt.subplots(figsize=(11, 5.2))
m = mod.copy()
sns.stripplot(data=m, x="Payload_kg", y="Launch_Site", hue="Outcome", palette={"Success": GREEN, "Failure": RED},
              size=9, alpha=0.85, ax=ax, order=sites)
ax.set_xlabel("Payload mass (kg)"); ax.set_ylabel(""); ax.set_title("Payload mass vs launch site (landing outcome)")
ax.legend(title="Landing", loc="lower right")
save(fig, "payload_vs_site")

# --- Success rate by orbit
o = mod.groupby("Orbit").agg(n=("Landing_class", "size"), rate=("Landing_class", "mean")).sort_values("rate")
fig, ax = plt.subplots(figsize=(10, 5.2))
bars = ax.barh(o.index, o.rate * 100, color=BLUE)
for b, n, r in zip(bars, o.n, o.rate):
    ax.text(r * 100 + 1, b.get_y() + b.get_height() / 2, f"{r*100:.0f}%  (n={n})", va="center")
ax.set_xlim(0, 118); ax.set_xlabel("Landing success rate (%)"); ax.set_title("Landing success rate by orbit type")
save(fig, "success_by_orbit")

# --- Flight number vs orbit type (launch order coloured by orbit)
fig, ax = plt.subplots(figsize=(11, 5.2))
orbits = sorted(df.Orbit.unique())
pal = sns.color_palette("tab10", len(orbits))
for i, orb in enumerate(orbits):
    d = df[df.Orbit == orb]
    ax.scatter(d.Flight_number, [i] * len(d), s=60, color=pal[i], alpha=0.8, edgecolor="white")
ax.set_yticks(range(len(orbits))); ax.set_yticklabels(orbits)
ax.set_xlabel("Flight number (chronological)"); ax.set_title("Flight number vs orbit type")
save(fig, "flight_vs_orbit")

# --- Payload vs orbit type
fig, ax = plt.subplots(figsize=(11, 5.2))
sns.stripplot(data=df, x="Payload_kg", y="Orbit", hue="Landing_Outcome", size=8, alpha=0.8, ax=ax, legend=False)
ax.set_xlabel("Payload mass (kg)"); ax.set_ylabel(""); ax.set_title("Payload mass vs orbit type")
save(fig, "payload_vs_orbit")

# --- Yearly success trend (line) with launches
yr = df.groupby(df.Date.dt.year).agg(launches=("Flight_number", "size"),
                                      attempted=("Landing_class", "count"), succ=("Landing_class", "sum"))
yr["rate"] = yr.succ / yr.attempted
fig, ax = plt.subplots(figsize=(11, 5.2))
ax.plot(yr.index, yr.rate * 100, "o-", color=RED, lw=3, ms=9, label="Landing success rate (%)")
for x, r, a in zip(yr.index, yr.rate, yr.attempted):
    if not np.isnan(r): ax.annotate(f"{r*100:.0f}%", (x, r * 100), textcoords="offset points", xytext=(0, 10), ha="center")
ax.set_ylim(-5, 110); ax.set_xlabel("Year"); ax.set_ylabel("Landing success (%)")
ax2 = ax.twinx(); ax2.bar(yr.index, yr.launches, alpha=0.25, color=BLUE); ax2.set_ylabel("Launches")
ax.set_title("Yearly landing success rate and launch count")
save(fig, "yearly_trend")

# --- Model comparison bar (CV and test)
mc = pd.read_csv(f"{BASE}/outputs/model_comparison.csv")
fig, ax = plt.subplots(figsize=(10, 5.2))
x = np.arange(len(mc)); w = 0.38
ax.bar(x - w / 2, mc.cv_accuracy * 100, w, color=BLUE, label="Cross-validation accuracy")
ax.bar(x + w / 2, mc.test_accuracy * 100, w, color=DBLUE, alpha=0.6, label="Test accuracy (n=15)")
for i, (c, t) in enumerate(zip(mc.cv_accuracy, mc.test_accuracy)):
    ax.text(i - w / 2, c * 100 + 1, f"{c*100:.1f}", ha="center", fontsize=11)
    ax.text(i + w / 2, t * 100 + 1, f"{t*100:.1f}", ha="center", fontsize=11)
ax.set_xticks(x); ax.set_xticklabels(mc.model, rotation=10); ax.set_ylim(0, 110)
ax.set_ylabel("Accuracy (%)"); ax.set_title("Classification accuracy by model"); ax.legend(loc="lower right")
save(fig, "model_accuracy")

# --- Folium-style map snapshots (static). Sites drawn at approximate public coordinates.
COORDS = {"CCAFS LC-40": (28.5623, -80.5774), "CCAFS SLC-40": (28.5632, -80.5768),
          "KSC LC-39A": (28.5733, -80.6469), "VAFB SLC-4E": (34.6328, -120.6108)}
site = df.groupby("Launch_Site").agg(launches=("Flight_number", "size"),
                                     attempted=("Landing_class", "count"), succ=("Landing_class", "sum"))
site["rate"] = site.succ / site.attempted
def snapshot(name, title, color_by_rate):
    fig = plt.figure(figsize=(14, 6.4))
    gs = fig.add_gridspec(1, 2, width_ratios=[1.35, 1])
    ax = fig.add_subplot(gs[0, 0]); zx = fig.add_subplot(gs[0, 1])
    def colour(s):
        if not color_by_rate: return DBLUE
        r = site.loc[s].rate
        return GREEN if r >= 0.8 else ("#D9822B" if r >= 0.6 else RED)
    for a in (ax, zx):
        a.set_facecolor("#EAF2FA"); a.grid(True, color="white", lw=1)
    # global view
    ax.set_xlim(-180, 180); ax.set_ylim(-60, 75)
    ax.set_xlabel("Longitude"); ax.set_ylabel("Latitude"); ax.set_title("Global view: all four launch sites")
    for s_, (lat, lon) in COORDS.items():
        ax.scatter(lon, lat, s=140, color=colour(s_), edgecolor="white", zorder=5)
    ax.annotate("CCAFS LC-40 / SLC-40\n(Cape Canaveral, Florida)", (-80.6, 28.6), xytext=(-150, 45),
                arrowprops=dict(arrowstyle="-", color=GREY), fontsize=10,
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=GREY))
    ax.annotate("KSC LC-39A\n(Florida)", (-80.65, 28.57), xytext=(-10, 60), arrowprops=dict(arrowstyle="-", color=GREY),
                fontsize=10, bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=GREY))
    ax.annotate("VAFB SLC-4E\n(California)", (-120.6, 34.6), xytext=(-175, 15), arrowprops=dict(arrowstyle="-", color=GREY),
                fontsize=10, bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=GREY))
    # zoomed Cape Canaveral / KSC view
    zx.set_xlim(-80.75, -80.45); zx.set_ylim(28.45, 28.7)
    zx.set_xlabel("Longitude"); zx.set_ylabel("Latitude"); zx.set_title("Zoom: Cape Canaveral and KSC")
    offs = {"CCAFS LC-40": (-95, 40), "CCAFS SLC-40": (40, 75), "KSC LC-39A": (-150, -55)}
    for s_ in ["CCAFS LC-40", "CCAFS SLC-40", "KSC LC-39A"]:
        lat, lon = COORDS[s_]; r = site.loc[s_]
        zx.scatter(lon, lat, s=max(120, 14 * r.launches), color=colour(s_), edgecolor="white", zorder=5, alpha=0.9)
        label = f"{s_}\n{int(r.launches)} launches" + (f"\n{r.rate*100:.0f}% landed" if color_by_rate else "")
        zx.annotate(label, (lon, lat), xytext=offs[s_], textcoords="offset points", fontsize=10,
                    arrowprops=dict(arrowstyle="-", color=GREY),
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=colour(s_)))
    fig.suptitle(title, fontsize=15, y=0.99)
    fig.tight_layout()
    fig.savefig(f"{FIG}/{name}.png", dpi=170, facecolor="white"); plt.close(fig)
snapshot("map_all_sites", "Launch sites on a global map and at Cape Canaveral (marker size = launch count)", False)
snapshot("map_outcome_colours", "Launch sites coloured by landing success rate (green ≥80%, orange ≥60%, red <60%)", True)

# --- Proximity (distance from each site to a reference point, haversine)
def haversine_km(a, b):
    R = 6371.0
    la1, lo1, la2, lo2 = map(math.radians, [a[0], a[1], b[0], b[1]])
    d = 2 * R * math.asin(math.sqrt(math.sin((la2 - la1) / 2) ** 2 +
                                    math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2))
    return d
JSC = (29.5502, -95.0977)
dist = {s: haversine_km(c, JSC) for s, c in COORDS.items()}
pd.Series(dist, name="km_to_JSC").round(0).to_csv(f"{BASE}/outputs/proximity_km.csv")
fig, ax = plt.subplots(figsize=(9, 5))
ks = sorted(dist, key=dist.get)
ax.barh(ks, [dist[k] for k in ks], color=BLUE)
for i, k in enumerate(ks): ax.text(dist[k] + 40, i, f"{dist[k]:,.0f} km", va="center")
ax.set_xlabel("Great-circle distance (km)"); ax.set_xlim(0, max(dist.values()) * 1.2)
ax.set_title("Distance from each launch site to NASA Johnson Space Center (reference)")
save(fig, "proximity_jsc")

# --- Dash dashboard views (same logic as scripts/spacex_dash_app.py)
oc = df.groupby("Landing_Outcome").size().sort_values(ascending=False)
fig, ax = plt.subplots(figsize=(8.5, 6.2))
wedges, _, _ = ax.pie(oc.values, labels=None, autopct="%1.0f%%", startangle=90, pctdistance=0.8,
                      colors=sns.color_palette("Blues_r", len(oc)), wedgeprops=dict(width=0.45))
ax.legend(wedges, [f"{k} ({v})" for k, v in oc.items()], loc="center left", bbox_to_anchor=(1, 0.5), fontsize=10)
ax.set_title("Dashboard view 1: landing outcomes, all sites")
save(fig, "dash_pie_all")

ksc = df[df.Launch_Site == "KSC LC-39A"].groupby("Landing_Outcome").size().sort_values(ascending=False)
fig, ax = plt.subplots(figsize=(8.5, 6.2))
wedges, _, _ = ax.pie(ksc.values, autopct="%1.0f%%", startangle=90, pctdistance=0.8,
                      colors=sns.color_palette("Greens_r", len(ksc)), wedgeprops=dict(width=0.45))
ax.legend(wedges, [f"{k} ({v})" for k, v in ksc.items()], loc="center left", bbox_to_anchor=(1, 0.5), fontsize=10)
ax.set_title("Dashboard view 2: KSC LC-39A (highest landing success rate)")
save(fig, "dash_pie_ksc")

fig, axes = plt.subplots(1, 2, figsize=(14, 5.2), sharey=True)
for ax, (lo, hi), t in zip(axes, [(0, 5000), (5000, 10000)], ["0–5,000 kg", "5,000–10,000 kg"]):
    d = mod[mod.Payload_kg.between(lo, hi)]
    sns.scatterplot(data=d, x="Payload_kg", y="Booster_family", hue="Outcome", style="Orbit",
                    palette={"Success": GREEN, "Failure": RED}, s=90, ax=ax, legend=(ax is axes[1]))
    ax.set_title(f"Dashboard view 3: payload {t}"); ax.set_xlabel("Payload (kg)"); ax.set_ylabel("")
    ax.set_xlim(lo - 200, hi + 200)
if axes[1].get_legend(): axes[1].legend(bbox_to_anchor=(1.01, 1), loc="upper left", fontsize=9)
save(fig, "dash_payload_range")

print("figures:", sorted(os.listdir(FIG)))
print("distances km:", {k: round(v) for k, v in dist.items()})
print("site table:\n", site.round(3))
