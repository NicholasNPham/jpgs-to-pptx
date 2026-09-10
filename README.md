# JPGs to PPTX

Converts a folder of scanned document pages (JPGs) into a single PowerPoint file, one page per slide, automatically sized and centered, ready to review or present without manual copy-pasting.

## What problem it solves

Manually copying dozens of scanned document images into PowerPoint, one at a time, resizing and centering each one by hand, is slow and error-prone. This tool automates that entire process into a single click.

## How it works

Point the tool at a folder of scanned `.jpg`/`.jpeg` pages (or just drop the built executable into that folder). It finds every image, places each one on its own slide against a black background, and automatically shrinks any oversized scan to fit the slide without stretching or distorting it, while leaving smaller images at their true size. The result is a single PowerPoint file that looks like each page was neatly pasted in by hand.

## Tech used

- **Python 3**
- [**python-pptx**](https://python-pptx.readthedocs.io/) — builds the PowerPoint file
- [**Pillow (PIL)**](https://pillow.readthedocs.io/) — reads image dimensions for scaling
- [**PyInstaller**](https://pyinstaller.org/) — packages the script into a standalone Windows `.exe`

## How to run it

### Option A: Run as a Python script

```powershell
pip install -r requirements.txt
python jpgs_to_pptx.py
```

Place `jpgs_to_pptx.py` in the same folder as your `.jpg`/`.jpeg` images before running. Output is saved as `output.pptx` in that same folder.

### Option B: Build and run as a standalone .exe

No Python installation needed for the end user, just the compiled `.exe`.

```powershell
pip install -r requirements.txt
pip install pyinstaller
python -m PyInstaller --onefile --console --name jpgs_to_pptx --icon=icon.ico --hidden-import=pptx --hidden-import=pptx.oxml --collect-all=pptx jpgs_to_pptx.py
```

This creates `dist/jpgs_to_pptx.exe`. Drop that `.exe` into any folder containing `.jpg`/`.jpeg` images and run it. It detects its own location, scans that folder, and saves `output.pptx` right next to itself, no configuration needed.
