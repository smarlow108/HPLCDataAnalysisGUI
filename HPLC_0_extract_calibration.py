import os
import struct
import csv
import re
from tkinter import Tk, filedialog

def utf16le_strings(raw, minlen=3):
    """Extract readable UTF-16LE strings from binary data."""
    out, buf = [], []
    for i in range(0, len(raw)-1, 2):
        ch, nz = raw[i], raw[i+1]
        if nz == 0 and (32 <= ch <= 126 or ch in (9,10,13)):
            buf.append(chr(ch))
        else:
            if len(buf) >= minlen:
                out.append(''.join(buf))
            buf = []
    if len(buf) >= minlen:
        out.append(''.join(buf))
    return out

def find_offsets_utf16le(raw, term):
    """Find positions of a UTF-16LE term inside a binary file."""
    pattern = term.encode('utf-16-le')
    i, offs = 0, []
    while True:
        j = raw.find(pattern, i)
        if j == -1: break
        offs.append(j); i = j + 2
    return offs

def nearby_doubles(raw, off, span=256, limit=(0,300)):
    """Scan nearby bytes for plausible float/double values."""
    vals, win = [], raw[off:off+span]
    for j in range(0, max(0,len(win)-8), 4):
        for fmt in ("<d","<f"):
            try:
                v = struct.unpack_from(fmt, win, j)[0]
                if limit[0] <= v <= limit[1]:
                    vals.append(round(v,6))
            except: pass
    return sorted(set(vals))

def parse_da_method(path):
    """Extract calibration table from DAMETHOD.REG"""
    with open(path, "rb") as f:
        raw = f.read()
    strings = utf16le_strings(raw, 3)
    comp_idx = [i for i,s in enumerate(strings) if s.strip() == "Compound"]
    compounds = []
    for i, idx in enumerate(comp_idx):
        end = comp_idx[i+1] if i+1 < len(comp_idx) else len(strings)
        block = strings[idx:end]
        # guess compound name
        name = next((s.strip() for s in block[1:10]
                     if re.fullmatch(r"[A-Za-z0-9 _\-\(\)\.\+]{2,}", s)), None)
        # extract RT info
        exp_rt = nearby_doubles(raw, raw.find(b"ExpRetTime"), limit=(0,300))
        rt_low = nearby_doubles(raw, raw.find(b"RTWinLow"), limit=(0,300))
        rt_high = nearby_doubles(raw, raw.find(b"RTWinHigh"), limit=(0,300))
        # curve parameters
        slope = nearby_doubles(raw, raw.find(b"CurveParm1"), limit=(-1e6,1e6))
        intercept = nearby_doubles(raw, raw.find(b"CurveParm2"), limit=(-1e6,1e6))
        compounds.append({
            "Compound": name or f"Compound_{len(compounds)+1}",
            "Expected_RT_min": exp_rt[0] if exp_rt else "",
            "RT_Window_Low": rt_low[0] if rt_low else "",
            "RT_Window_High": rt_high[0] if rt_high else "",
            "Curve_Slope": slope[0] if slope else "",
            "Curve_Intercept": intercept[0] if intercept else ""
        })
    return compounds

def save_to_csv(compounds, out_path):
    """Save calibration table to CSV."""
    if not compounds:
        print("⚠️ No compounds found in the method file.")
        return
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=compounds[0].keys())
        writer.writeheader()
        writer.writerows(compounds)
    print(f"✅ Calibration table saved to: {out_path}")

def main():
    # GUI folder picker
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Select Agilent Method or Data Folder")
    if not folder:
        print("❌ No folder selected.")
        return

    print(f"📁 Scanning folder: {folder}")
    method_file = None
    for root, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith("damethod.reg"):
                method_file = os.path.join(root, f)
                break
        if method_file: break

    if not method_file:
        print("❌ No DAMETHOD.REG found in this folder.")
        return

    print(f"🔍 Found method: {method_file}")
    compounds = parse_da_method(method_file)
    out_path = os.path.join(folder, "calibration_table.csv")
    save_to_csv(compounds, out_path)

if __name__ == "__main__":
    main()
