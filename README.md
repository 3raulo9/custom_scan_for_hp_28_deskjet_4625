# HP DeskJet 4620 Scan

A one-window scanner app for the HP DeskJet 4620 series (DeskJet Ink Advantage 4625).
It scans from the flatbed and saves a JPEG. There are four options: Scan Size, Output Type, Resolution and Item Type.

- **Resolution** defaults to 1200 dpi, the scanner's maximum optical resolution.
- The scanner sends uncompressed data, and the app compresses it only once: JPEG quality 95, no chroma subsampling, real DPI and an sRGB profile in the file.
- **Item Type → Document** turns the paper pure white and lightly sharpens text. **Photo** keeps the colors exactly as scanned.
- Scans go to `Pictures\Scans` by default. The app remembers your last settings.

## Run

`dist\HP4620Scan.exe`

## Rebuild the exe

```powershell
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
.\.venv\Scripts\pyinstaller --noconfirm --onefile --windowed --name HP4620Scan hp_scan.py
```

## If scanning fails with "Could not communicate with the scanner"

Turn the printer off, wait about 10 seconds and turn it back on. The HP driver's USB scan link often stops responding after the PC has been asleep. HP's own HP Scan app fails the same way when this happens.
