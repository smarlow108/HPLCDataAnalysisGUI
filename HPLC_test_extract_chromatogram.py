import numpy as np
import matplotlib.pyplot as plt
import os

# --- CONFIG ---
files = [
    r"C:\path\to\sample1\RID1A.ch",
    r"C:\path\to\sample2\RID1A.ch",
    r"C:\path\to\sample3\RID1A.ch"
]

period = 3  # same structure
window = 500  # smoothing for RMS

def estimate_wait_time(ch_file):
    raw = np.fromfile(ch_file, dtype=np.int16)
    raw = raw - np.median(raw)
    n_cycles = len(raw) // period
    signal = np.mean(raw[:n_cycles * period].reshape(n_cycles, period), axis=1)
    time = np.linspace(0, 40, len(signal))
    
    rms = np.sqrt(np.convolve(signal**2, np.ones(window)/window, mode='same'))
    threshold = 3 * np.median(rms[:2000])
    start_idx = np.argmax(rms > threshold)
    
    wait_time = time[start_idx]
    
    plt.figure(figsize=(8,3))
    plt.plot(time, signal, lw=0.8)
    plt.axvline(wait_time, color='r', linestyle='--', label=f"Start ~{wait_time:.2f} min")
    plt.title(os.path.basename(ch_file))
    plt.xlabel("Time (min)")
    plt.ylabel("Signal (a.u.)")
    plt.legend()
    plt.tight_layout()
    plt.show()
    
    return wait_time

# --- Run on all files ---
for f in files:
    if os.path.exists(f):
        wt = estimate_wait_time(f)
        print(f"{os.path.basename(f)} — Estimated wait time: {wt:.2f} min")
    else:
        print(f"File not found: {f}")
