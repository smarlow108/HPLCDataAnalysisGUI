"""
Recursive ChemStation .CH Processor (with integration)
- Pick a root folder in a dialog
- Recursively process all .CH files
- Uses calibrated parameters and applies transform = "int"
- Saves *_processed.png near each file
"""

import os
import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import medfilt, savgol_filter
from scipy.interpolate import interp1d
from tkinter import Tk, filedialog

# ── FINAL CALIBRATION from A5_VWD1A ─────────────────────────
CALIBRATION = {
    "offset": 616,
    "dtype": ">i2",
    "invert": False,
    "scale_guess": 10000,
    "trim": 2838,
    "time_scale": 0.980,
    "scale_a": 0.5933582708385587,
    "offset_b": 0.17597110192247636,
    "transform": "int",      #"int", "raw"   # <-- IMPORTANT
    "comb_filter": True
}

# ── helpers ─────────────────────────────────────────────────
def subtract_linear_baseline(y, robust=True):
    n = len(y)
    x = np.arange(n)
    if robust:
        nseg = max(1, n // 20)
        fit_x = np.concatenate([x[:nseg], x[-nseg:]])
        fit_y = np.concatenate([y[:nseg], y[-nseg:]])
        slope, intercept = np.polyfit(fit_x, fit_y, 1)
    else:
        slope = (y[-1] - y[0]) / (n - 1)
        intercept = y[0]
    baseline = intercept + slope * x
    return y - baseline, slope

def resample_xscale(y, scale):
    x_old = np.linspace(0, 1, len(y))
    x_new = np.linspace(0, 1, int(round(len(y) * scale)))
    f = interp1d(x_old, y, kind="linear", fill_value="extrapolate")
    return f(x_new)

# ── core processing ─────────────────────────────────────────
def process_ch_file(ch_path, P=CALIBRATION, show_plot=True):
    print(f"🔍 Processing {ch_path}")

    # 1) read binary
    with open(ch_path, "rb") as f:
        arr = np.frombuffer(f.read()[P["offset"]:], dtype=P["dtype"]).astype(float)
    if P["invert"]:
        arr = -arr
    arr /= P["scale_guess"]

    # 2) trim beginning
    if len(arr) > P["trim"]:
        arr = arr[P["trim"]:]

    # 3) optional de-comb filtering
    if P.get("comb_filter", True):
        arr = medfilt(arr, 5)
        arr = savgol_filter(arr, 15, 2)

    # 4) APPLY TRANSFORM (this is what was missing)
    tname = (P.get("transform") or "raw").lower()
    if tname == "diff":
        arr = np.gradient(arr)
    elif tname == "int":
        # integrate after mean-centering to limit drift; baseline removal handles the rest
        arr = np.cumsum(arr - np.mean(arr))
    # else "raw": do nothing

    # 5) x-axis rescale (time compression/expansion)
    arr = resample_xscale(arr, P["time_scale"])

    # 6) amplitude rescale (a) + offset (b)
    arr = arr * P["scale_a"] + P["offset_b"]

    # 7) final linear baseline removal
    processed, slope = subtract_linear_baseline(arr, robust=True)

    # 8) time axis in minutes (arbitrary uniform sampling)
    time = np.arange(len(processed), dtype=float) / 60.0

    # 9) save plot
    if show_plot:
        plt.figure(figsize=(10, 5))
        plt.plot(time, processed, lw=1.2)
        plt.xlabel("Time (min)")
        plt.ylabel("Intensity (a.u.)")
        tit = f"{os.path.basename(ch_path)} | transform={tname} | slope={slope:.2e}"
        plt.title(tit)
        out_path = os.path.splitext(ch_path)[0] + "_processed.png"
        plt.savefig(out_path, dpi=150)
        plt.close()
        print(f"📈 Saved: {out_path}")

    return {"file": ch_path, "baseline_slope": slope, "samples": len(processed)}

# ── folder UI + recursion ───────────────────────────────────
def select_folder():
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Select Root Folder Containing .CH Files")
    if not folder:
        print("❌ No folder selected. Exiting.")
        raise SystemExit
    return folder

def find_ch_files(root):
    ch_files = []
    for dirpath, _, filenames in os.walk(root):
        for f in filenames:
            if f.lower().endswith(".ch"):
                ch_files.append(os.path.join(dirpath, f))
    return ch_files

# ── main ────────────────────────────────────────────────────
if __name__ == "__main__":
    print("📂 Please select the root folder containing .CH files...")
    root = select_folder()
    files = find_ch_files(root)
    print(f"🔎 Found {len(files)} .CH files under {root}")

    if not files:
        print("⚠️ No .CH files found.")
        raise SystemExit

    results = [process_ch_file(p) for p in files]

    print("\n✅ All files processed.")
    for r in results:
        print(f"  {os.path.basename(r['file']):30s}  baseline={r['baseline_slope']:.3e}")
