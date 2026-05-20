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
        mask = (df["Start Time"] >= "2026-04-16 10:00:00") & (df["Start Time"] < "2026-04-16 10:00:20.1")
        sub = df.loc[mask]
        if len(sub) > 0:
            dfs.append(sub)
    except:
        pass

df = pd.concat(dfs, ignore_index=True)
df.sort_values("Start Time", inplace=True)
df.reset_index(drop=True, inplace=True)

bands = ["10 Hz", "31.5 Hz"]
result = {}
for fn in bands:
    rows = []
    for i in range(min(200, len(df))):
        t = str(df["Start Time"].iloc[i])
        x_col, y_col, z_col = f"X_{fn}", f"Y_{fn}", f"Z_{fn}"
        x = float(df[x_col].iloc[i]) if pd.notna(df[x_col].iloc[i]) else 0.0
        y = float(df[y_col].iloc[i]) if pd.notna(df[y_col].iloc[i]) else 0.0
        z = float(df[z_col].iloc[i]) if pd.notna(df[z_col].iloc[i]) else 0.0
        xyz = np.sqrt(x**2 + y**2 + z**2)
        rows.append({"time": t, "X": x, "Y": y, "Z": z, "XYZ": round(xyz, 8)})
    result[fn] = rows

with open(os.path.join(root, "raw_1s_demo.json"), "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
print(f"OK: {len(result['10 Hz'])} rows per band (= {len(result['10 Hz'])//10} seconds)")
