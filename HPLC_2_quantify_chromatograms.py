import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import sparse
from scipy.sparse.linalg import spsolve
from scipy.signal import find_peaks, peak_widths

# ---------- Baseline correction ----------
def baseline_als(y, lam=1e6, p=0.01, niter=10):
    L = len(y)
    D = sparse.diags([1, -2, 1], [0, -1, -2], shape=(L, L-2))
    w = np.ones(L)
    for _ in range(niter):
        W = sparse.spdiags(w, 0, L, L)
        Z = W + lam * D.dot(D.transpose())
        z = spsolve(Z, w * y)
        w = p * (y > z) + (1 - p) * (y < z)
    return z

# ---------- Peak integration ----------
def integrate_peak(x, y, left, right):
    return np.trapz(y[left:right], x[left:right])

# ---------- Process detector ----------
def process_detector(detector_file, calib_df, detector_name):
    print(f"\n📈 Processing {detector_name} ...")
    df = pd.read_csv(detector_file)
    t, y = df["Time_min"].values, df["Signal"].values
    print(f"   ↳ Time range: {t[0]:.2f} – {t[-1]:.2f} min  ({len(t)} points)")
    print(f"   ↳ Signal range: {np.min(y):.4g} – {np.max(y):.4g}")

    # Autoscale to 0–1 range
    y_min, y_max = np.min(y), np.max(y)
    if y_max - y_min == 0:
        print("⚠️ Flat signal, skipping.")
        return pd.DataFrame()
    y_scaled = (y - y_min) / (y_max - y_min)

    # Baseline correction on scaled signal
    baseline_scaled = baseline_als(y_scaled, lam=1e6, p=0.01)
    y_corr_scaled = y_scaled - baseline_scaled

    # Restore to original scale
    baseline = baseline_scaled * (y_max - y_min) + y_min
    y_corr = y_corr_scaled * (y_max - y_min)

    # Adaptive prominence for peak detection
    prominence = np.std(y_corr) * 2
    peaks, _ = find_peaks(y_corr, prominence=prominence)
    widths, _, left_ips, right_ips = peak_widths(y_corr, peaks, rel_height=0.5)
    print(f"   ↳ Found {len(peaks)} peaks")

    # Quantify peaks vs calibration
    results = []
    for _, row in calib_df.iterrows():
        name = row["Compound"]
        rt = row["Expected_RT_min"]
        rt_low = float(row.get("RT_Window_Low", 0.1) or 0.1)
        rt_high = float(row.get("RT_Window_High", 0.1) or 0.1)
        slope = float(row.get("Curve_Slope", np.nan))
        intercept = float(row.get("Curve_Intercept", 0))

        # match peaks
        mask = (t[peaks] >= (rt - rt_low)) & (t[peaks] <= (rt + rt_high))
        if not any(mask):
            # widen search window automatically if no match
            mask = (t[peaks] >= (rt - rt_low*5)) & (t[peaks] <= (rt + rt_high*5))
            if not any(mask):
                continue

        idx = np.where(mask)[0][0]
        left = int(left_ips[idx])
        right = int(right_ips[idx])
        area = integrate_peak(t, y_corr, left, right)
        peak_rt = t[peaks[idx]]

        conc = np.nan
        if np.isfinite(slope) and slope != 0:
            conc = (area - intercept) / slope

        results.append({
            "Compound": name,
            f"{detector_name}_Peak_RT_min": peak_rt,
            f"{detector_name}_Peak_Area": area,
            f"{detector_name}_Concentration": conc
        })

    # Plot chromatogram
    plt.figure(figsize=(10, 5))
    plt.plot(t, y, color="gray", lw=1, label="Raw Signal")
    plt.plot(t, baseline, "k--", lw=1, label="Baseline")
    plt.plot(t, y_corr, color="blue", lw=1, label="Corrected Signal")
    plt.scatter(t[peaks], y_corr[peaks], color="red", s=25, zorder=5)
    plt.title(f"{detector_name} Chromatogram")
    plt.xlabel("Time (min)")
    plt.ylabel("Signal (a.u.)")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.show()

    return pd.DataFrame(results)

# ---------- Main workflow ----------
def main():
    folder = input("📁 Enter the folder containing RID1A.csv, VWD1A.csv, and calibration_table.csv:\n> ").strip('" ')
    calib_path = os.path.join(folder, "calibration_table.csv")
    if not os.path.exists(calib_path):
        print("❌ calibration_table.csv not found.")
        return

    calib_df = pd.read_csv(calib_path)
    print(f"✅ Loaded calibration with {len(calib_df)} compounds.")
    combined_df = calib_df[["Compound", "Expected_RT_min"]].copy()

    for detector_name, file_name in [("RID1A", "RID1A.csv"), ("VWD1A", "VWD1A.csv")]:
        path = os.path.join(folder, file_name)
        if os.path.exists(path):
            det_results = process_detector(path, calib_df, detector_name)
            combined_df = combined_df.merge(det_results, on="Compound", how="left")
        else:
            print(f"⚠️ {file_name} not found in folder.")

    out_csv = os.path.join(folder, "combined_results.csv")
    combined_df.to_csv(out_csv, index=False)
    print(f"\n✅ Combined quantitation results saved to: {out_csv}")

if __name__ == "__main__":
    main()
