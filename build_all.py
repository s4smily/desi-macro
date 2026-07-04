#!/usr/bin/env python3
"""
build_all.py — one command for the whole weekly build.

    python build_all.py bd_macro_data.xlsx

Steps:
  1. Excel  -> data.json          (make_data.py)
  2. data.json embedded into index.html as offline fallback
  3. data.json -> DCM_Weekly_<week>.pptx  +  .pdf   (make_deck.py, weekly data only)

Upload to GitHub afterwards:  index.html, data.json   (the site)
Email to colleagues:          the .pptx and/or .pdf

Notes:
  * Works on Windows, macOS and Linux.
  * Run it from anywhere; it always finds its own folder and writes outputs there.
  * Uses the SAME Python that launched it (no dependence on a 'python3' command).
"""
import sys, json, subprocess, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
PY   = sys.executable or "python"          # the exact interpreter running this file

def run(cmd, cwd):
    print("»", " ".join(cmd))
    subprocess.run(cmd, check=True, cwd=cwd)

def main():
    # Accept the Excel file as given (absolute, relative, or bare name).
    arg = sys.argv[1] if len(sys.argv) > 1 else "bd_macro_data.xlsx"

    # Resolve the Excel path: try as given, then next to this script.
    candidates = [os.path.abspath(arg), os.path.join(HERE, os.path.basename(arg))]
    xlsx = next((p for p in candidates if os.path.isfile(p)), None)
    if not xlsx:
        print("ERROR: could not find the Excel file.")
        print("  Looked for:")
        for c in candidates:
            print("   -", c)
        print("Put bd_macro_data.xlsx in the same folder as build_all.py, then run again.")
        sys.exit(1)

    print("Using workbook:", xlsx)
    data_json = os.path.join(HERE, "data.json")

    # 1. Excel -> data.json  (run inside HERE so imports + outputs resolve there)
    run([PY, os.path.join(HERE, "make_data.py"), xlsx, data_json], cwd=HERE)

    # 2. embed fallback into index.html
    with open(data_json, encoding="utf-8") as f:
        data = json.load(f)
    idx = os.path.join(HERE, "index.html")
    with open(idx, encoding="utf-8") as f:
        html = f.read()
    payload = json.dumps(data, ensure_ascii=False)
    new, n = re.subn(r"^const FALLBACK_DATA\s*=.*;\s*$",
                     "const FALLBACK_DATA = " + payload.replace(chr(92), chr(92)*2) + ";",
                     html, count=1, flags=re.M)
    if n == 0:
        new = html.replace("/*__DATA__*/null/*__END__*/", payload, 1)
    with open(idx, "w", encoding="utf-8") as f:
        f.write(new)
    print("\u2713 Embedded latest week into index.html as offline fallback")

    # 3. data.json -> PPTX + PDF
    run([PY, os.path.join(HERE, "make_deck.py"), data_json, "--pdf"], cwd=HERE)

    print("\nDONE. Upload index.html + data.json to GitHub. Email the .pptx / .pdf.")

if __name__ == "__main__":
    main()
