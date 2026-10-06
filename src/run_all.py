import subprocess
import sys
from pathlib import Path

# List of python files to run (order matters)
FILES = [
    "adityaBirlaCompilation.py",
    "axisCompilation.py",
    "bandhanCompilation.py",
    "dspAuto.py",
    "HDFCCompilation.py",
    "ICICICompilation.py",
    "kotakCompilation.py",
    "miraeCompilation.py",
    "nipponCompilation.py",
    "SBICompilation.py",
]

python_exe = sys.executable  # uses the same Python that's running this file
base_path = Path(__file__).parent

for file in FILES:import subprocess
import sys
from pathlib import Path

python_exe = sys.executable
base_path = Path(__file__).parent

# Scripts with NO arguments
simple_scripts = [
    "adityaBirlaCompilation.py",
    "axisCompilation.py",
    "bandhanCompilation.py",
    "HDFCCompilation.py",
    "ICICICompilation.py",
    "kotakCompilation.py",
    "miraeCompilation.py",
    "nipponCompilation.py",
    "SBICompilation.py",
]

# DSP with arguments
dsp_command = [
    python_exe,
    str(base_path / "dspAuto.py"),
    str(base_path / "DSP.xlsx"),
    str(base_path / "CompiledFiles" / "DSPCompiled.xlsx"),
]

# Run normal scripts
for script in simple_scripts:
    print(f"\n🔥 Running {script}...")
    result = subprocess.run(
        [python_exe, str(base_path / script)],
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        print(f"❌ {script} FAILED")
        print(result.stderr)
        sys.exit(1)
    else:
        print(f"✅ {script} completed")

# Run DSP separately
print("\n🔥 Running dspAuto.py with input/output files...")
result = subprocess.run(
    dsp_command,
    capture_output=True,
    text=True
)

if result.returncode != 0:
    print("❌ dspAuto.py FAILED")
    print(result.stderr)
    sys.exit(1)
else:
    print("✅ dspAuto.py completed successfully")

print("\n🏁 ALL COMPILATIONS FINISHED SUCCESSFULLY")
