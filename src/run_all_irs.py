import subprocess
import os
from pathlib import Path

# ── CONFIG ──────────────────────────────────────────────
SCRIPTS = {
    "aditya birla": "adityaBirlaIRS.py",
    "axis":         "axisIrs.py",
    "bandhan":      "bandhanIRS.py",
    "dsp":          "extract_dsp_irs.py",
    "hdfc":         "extract_hdfc_irs.py",
    "icici":        "extract_icici_irs.py",
    "kotak":        "extract_kotak_irs.py",
    "nippon":       "extract_nippon_irs.py",
    "sbi":          "sbiIRS.py",
    "tata":         "tataIRS.py",
}
OUTPUT_FOLDER = "outputs"
# ────────────────────────────────────────────────────────

Path(OUTPUT_FOLDER).mkdir(exist_ok=True)

for folder, script_name in SCRIPTS.items():
    script = os.path.join(folder, script_name)
    if not os.path.exists(script):
        print(f"[SKIP] {script} not found")
        continue

    print(f"[RUN] {script}")
    result = subprocess.run(
        ["python", script_name],
        capture_output=True,
        text=True,
        cwd=os.path.abspath(folder)  # run with the script's folder as cwd
    )

    out_file = os.path.join(OUTPUT_FOLDER, f"{Path(folder).name}_output.txt")
    with open(out_file, "w") as f:
        f.write(f"=== STDOUT ===\n{result.stdout}\n")
        f.write(f"=== STDERR ===\n{result.stderr}\n")
        f.write(f"=== EXIT CODE ===\n{result.returncode}\n")

    print(f"[SAVED] → {out_file}")

print("Done.")