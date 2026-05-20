import pandas as pd, numpy as np, glob, os, json

root = r"C:\Users\AC717\背景振動量測數據"
files = sorted(glob.glob(os.path.join(root, "0416", "*.rnd")))

dfs = []
for f in files:
    try:
        df = pd.read_csv(f, skiprows=1, na_values=["--","UN","OL",""], dtype=str, low_memory=False)
        df.columns = [c.strip() for c in df.columns]
        df["Start Time"] = pd.to_datetime(df["Start Time"].str.strip(),
                                           format="%Y/%m/%d %H:%M:%S.%f", errors="coerce")
        df.dropna(subset=["Start Time"], inplace=True)
        mask = (df["Start Time"] >= "2026-04-16 10:00:00") & (df["Start Time"] < "2026-04-16 11:00:00")
        sub = df.loc[mask]
        if len(sub) > 0:
            dfs.append(sub)
    except:
        pass

df = pd.concat(dfs, ignore_index=True)
df.sort_values("Start Time", inplace=True)
df.reset_index(drop=True, inplace=True)
print(f"Total rows: {len(df)}")
print(f"Time: {df['Start Time'].iloc[0]} ~ {df['Start Time'].iloc[-1]}")

AXES = ["X", "Y", "Z"]
FREQ_BANDS = [
    (1,"1 Hz"),(1.25,"1.25 Hz"),(1.6,"1.6 Hz"),(2,"2 Hz"),(2.5,"2.5 Hz"),
    (3.15,"3.15 Hz"),(4,"4 Hz"),(5,"5 Hz"),(6.3,"6.3 Hz"),(8,"8 Hz"),
    (10,"10 Hz"),(12.5,"12.5 Hz"),(16,"16 Hz"),(20,"20 Hz"),(25,"25 Hz"),
    (31.5,"31.5 Hz"),(40,"40 Hz"),(50,"50 Hz"),(63,"63 Hz"),(80,"80 Hz"),
    (100,"100 Hz"),
]

# Convert to numeric
for fv, fn in FREQ_BANDS:
    for ax in AXES:
        col = f"{ax}_{fn}"
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

# Compute XYZ composite per band
for fv, fn in FREQ_BANDS:
    cols = [f"{ax}_{fn}" for ax in AXES if f"{ax}_{fn}" in df.columns]
    if cols:
        mat = df[cols].values.astype(np.float64)
        df[f"XYZ_{fn}"] = np.sqrt(np.sum(mat**2, axis=1))

# Stats per band (XYZ composite) - raw 0.1s data
def calc_stats(vals):
    v = vals[np.isfinite(vals) & (vals > 0)]
    if len(v) == 0:
        return {s: np.nan for s in ["Leq","L5","L10","L50","L90","L95","Lmax"]}
    return {
        "Leq": float(np.sqrt(np.mean(v**2))),
        "L5":  float(np.percentile(v, 95)),
        "L10": float(np.percentile(v, 90)),
        "L50": float(np.percentile(v, 50)),
        "L90": float(np.percentile(v, 10)),
        "L95": float(np.percentile(v, 5)),
        "Lmax":float(np.max(v)),
    }

# Acceleration stats per band (XYZ)
acc_stats = {}
for fv, fn in FREQ_BANDS:
    col = f"XYZ_{fn}"
    if col in df.columns:
        acc_stats[fn] = calc_stats(df[col].values)

# Velocity conversion: v = a / (2*pi*f) => um/s = (m/s²) / (2*pi*f) * 1e6
vel_stats = {}
for fv, fn in FREQ_BANDS:
    col = f"XYZ_{fn}"
    if col in df.columns:
        vel_vals = df[col].values / (2.0 * np.pi * fv) * 1e6  # um/s
        vel_stats[fn] = calc_stats(vel_vals)

# dB acceleration (ref 1e-6 m/s²)
db_acc_stats = {}
for fv, fn in FREQ_BANDS:
    if fn in acc_stats:
        db_acc_stats[fn] = {}
        for s in ["Leq","L5","L10","L50","L90","L95","Lmax"]:
            v = acc_stats[fn][s]
            db_acc_stats[fn][s] = 20*np.log10(v/1e-6) if v > 0 else np.nan

# dB velocity (ref 2.54e-8 m/s)
db_vel_stats = {}
for fv, fn in FREQ_BANDS:
    if fn in vel_stats:
        db_vel_stats[fn] = {}
        for s in ["Leq","L5","L10","L50","L90","L95","Lmax"]:
            v_ums = vel_stats[fn][s]
            v_ms = v_ums * 1e-6
            db_vel_stats[fn][s] = 20*np.log10(v_ms/2.54e-8) if v_ums > 0 else np.nan

# VC grade evaluation (4-80 Hz, using velocity um/s Leq)
VC_CURVES = {"工廠 (ISO)":803.2,"辦公室 (ISO)":402.6,"住宅 (ISO)":201.8,
             "手術室 (ISO)":101.1,"VC-A":50.7,"VC-B":25.4,"VC-C":12.7,"VC-D":6.38,"VC-E":3.20}
VC_ORDER = ["VC-E","VC-D","VC-C","VC-B","VC-A","手術室 (ISO)","住宅 (ISO)","辦公室 (ISO)","工廠 (ISO)"]

def vc_val(flat, f):
    if f >= 8: return flat
    steps = round(np.log2(8.0/f)*3)
    return flat * 10**(steps*2/20)

# Evaluate per stat
vc_freqs = [4,5,6.3,8,10,12.5,16,20,25,31.5,40,50,63,80]
vc_freq_names = [f"{f} Hz" if f != 3.15 else "3.15 Hz" for f in vc_freqs]

for stat_key in ["Leq","L5","L10","L50","L90","L95","Lmax"]:
    band_vels = []
    for fv in vc_freqs:
        fn = f"{fv} Hz"
        if fn in vel_stats:
            v = vel_stats[fn][stat_key]
            if np.isfinite(v) and v > 0:
                band_vels.append((fv, v))
    grade = "超過工廠 (ISO)"
    for name in VC_ORDER:
        flat = VC_CURVES[name]
        if all(vel <= vc_val(flat, fv) for fv, vel in band_vels):
            grade = name; break
    peak = max(v for _, v in band_vels) if band_vels else 0
    print(f"VC grade ({stat_key}): {grade}, peak={peak:.2f} um/s")

# Save results to JSON for the Excel script
result = {
    "n_rows": len(df),
    "time_start": str(df["Start Time"].iloc[0]),
    "time_end": str(df["Start Time"].iloc[-1]),
    "acc_stats": {fn: {s: round(v, 8) for s, v in st.items()} for fn, st in acc_stats.items()},
    "vel_stats": {fn: {s: round(v, 4) for s, v in st.items()} for fn, st in vel_stats.items()},
    "db_acc_stats": {fn: {s: round(v, 2) if np.isfinite(v) else None for s, v in st.items()} for fn, st in db_acc_stats.items()},
    "db_vel_stats": {fn: {s: round(v, 2) if np.isfinite(v) else None for s, v in st.items()} for fn, st in db_vel_stats.items()},
}

# Also get a few raw samples for illustration
samples = {}
for fn in ["10 Hz", "31.5 Hz"]:
    col = f"XYZ_{fn}"
    if col in df.columns:
        arr = df[col].values[:20].tolist()
        samples[fn] = [round(v, 8) for v in arr]
result["samples"] = samples

# Broadband
bb_vals = np.zeros(len(df))
for fv, fn in FREQ_BANDS:
    col = f"XYZ_{fn}"
    if col in df.columns:
        bb_vals += df[col].values ** 2
bb_vals = np.sqrt(bb_vals)
result["broadband"] = {s: round(v, 8) for s, v in calc_stats(bb_vals).items()}

with open(r"C:\Users\AC717\背景振動量測數據\sample_result.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
print("Saved sample_result.json")
