import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress
import os

# ---------------- CONFIG ----------------
CH_PATH  = r"C:\Users\smarlow\Downloads\ChatGPTAutomation\A5_VWD1A.ch"
TXT_PATH = r"C:\Users\smarlow\Downloads\ChatGPTAutomation\A5_VWD.txt"

PERIOD = 3
TOTAL_TIME = 40.0  # minutes
TRIM_START_MIN = 13.5  # only trim start

# ---------------- LOAD TXT ----------------
txt = np.loadtxt(TXT_PATH)
time_txt = txt[:, 0]
signal_txt = txt[:, 1]

# ---------------- LOAD RAW .CH ----------------
raw = np.fromfile(CH_PATH, dtype=np.int16)
n_cycles = len(raw) // PERIOD
raw = raw[:n_cycles * PERIOD].reshape(n_cycles, PERIOD)
signal_ch = np.mean(raw, axis=1)

# Build full CH time axis (0–40 min)
time_ch = np.linspace(0, TOTAL_TIME, len(signal_ch))

# --- Trim only the beginning ---
mask = time_ch >= TRIM_START_MIN
signal_ch = signal_ch[mask]
time_ch = time_ch[mask]

# --- Rescale CH time axis to match TXT (0–40 min) ---
# Stretch CH so its trimmed range maps onto full 0–40 min
time_ch_scaled = np.linspace(0, TOTAL_TIME, len(signal_ch))

# Interpolate CH signal onto TXT time grid
signal_ch_interp = np.interp(time_txt, time_ch_scaled, signal_ch)

# --- Linear scaling fit ---
slope, intercept, r, p, se = linregress(signal_ch_interp, signal_txt)
converted = signal_ch_interp * slope + intercept

# --- Print summary ---
print(f"🔍 Conversion from .CH → .TXT (trimmed start only):")
print(f"    signal_txt ≈ {slope:.6f} * signal_ch + {intercept:.6f}")
print(f"    Correlation R² = {r**2:.6f}")
print(f"    CH range: {signal_ch.min():.1f} .. {signal_ch.max():.1f}")
print(f"    TXT range: {signal_txt.min():.6f} .. {signal_txt.max():.6f}")

# --- Plot comparison ---
plt.figure(figsize=(9,4))
plt.plot(time_txt, signal_txt, lw=1.2, label="ChemStation TXT (mAU)")
plt.plot(time_txt, converted, lw=1.2, color='crimson', alpha=0.8,
         label="Scaled CH (trimmed 13.5 min, time-scaled)")
plt.xlabel("Time (min)")
plt.ylabel("Signal (mAU)")
plt.title(f"Aligned comparison of {os.path.basename(CH_PATH)} and TXT")
plt.legend()
plt.grid(alpha=0.3)
plt.tight_layout()
plt.show()
