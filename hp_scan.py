"""One-button scanner app for the HP DeskJet 4620 series (DeskJet Ink Advantage 4625).

Scans from the flatbed through Windows Image Acquisition (WIA) and saves a
high-quality JPEG. Defaults to the scanner's maximum optical resolution.
"""

import ctypes
import itertools
import json
import os
import queue
import subprocess
import tempfile
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pythoncom
import pywintypes
import win32com.client
from PIL import Image, ImageFilter
from win32com.shell import shell, shellcon

APP_TITLE = "HP DeskJet 4620 Scan"
SETTINGS_FILE = Path(os.environ.get("APPDATA", Path.home())) / "HP4620Scan" / "settings.json"
SRGB_PROFILE = Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32/spool/drivers/color/sRGB Color Space Profile.icm"

Image.MAX_IMAGE_PIXELS = None  # a full page at 1200 dpi is ~143 megapixels

WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"  # the only format this driver offers
WIA_SCANNER_DEVICE = 1
WIA_FLATBED = 2
WIA_SUBTYPE_RANGE = 1

# WIA property IDs
PROP_DOCUMENT_HANDLING_SELECT = "3088"
PROP_DATA_TYPE = "4103"
PROP_BITS_PER_PIXEL = "4104"
PROP_X_RESOLUTION = "6147"
PROP_Y_RESOLUTION = "6148"
PROP_X_START = "6149"
PROP_Y_START = "6150"
PROP_X_EXTENT = "6151"
PROP_Y_EXTENT = "6152"

# label -> (WIA data type, bits per pixel)
OUTPUT_TYPES = {
    "Color": (3, 24),
    "Grayscale": (2, 8),
    "Black and White": (0, 1),
}

# label -> (width, height) in inches; None scans the whole glass
SCAN_SIZES = {
    "Entire Scan Area": None,
    "Letter (8.5 x 11 in)": (8.5, 11.0),
    "A4 (210 x 297 mm)": (210 / 25.4, 297 / 25.4),
    "A5 (148 x 210 mm)": (148 / 25.4, 210 / 25.4),
    "8 x 10 in": (8.0, 10.0),
    "5 x 7 in": (5.0, 7.0),
    "4 x 6 in": (4.0, 6.0),
}
GLASS_SIZE = (8.5, 11.69)  # inches, as reported by the driver

ITEM_TYPES = {
    "Document": "Paper is cleaned to pure white and text is lightly sharpened.",
    "Photo": "Colors are kept exactly as scanned, no processing.",
}

DEFAULT_RESOLUTIONS = [75, 100, 200, 300, 600, 1200]
JPEG_QUALITY = 95

WIA_ERRORS = {
    0x80210001: "The scanner reported a general error.",
    0x80210002: "Paper jam. Clear the jam and try again.",
    0x80210005: "The scanner is offline. Check that the printer is on and the USB cable is connected.",
    0x80210006: "The scanner is busy. Wait for it to finish and try again.",
    0x80210007: "The scanner is still warming up. Try again in a moment.",
    0x80210008: "The printer needs attention. Check its display.",
    0x8021000A: "Could not communicate with the scanner.\n\n"
                "Turn the printer off, wait 10 seconds and turn it back on. "
                "This often happens after the PC has been asleep.",
    0x8021000D: "The scanner is in use by another program.",
    0x80210015: "Scanner not found. Check that the printer is on and the USB cable is connected.",
}
RETRYABLE_ERRORS = {0x80210006, 0x80210007, 0x8021000A}


class ScanError(Exception):
    pass


# ---------------------------------------------------------------- scanner

def wia_error_code(err):
    """WIA puts its own HRESULT inside the generic 'exception occurred' COM error."""
    code = err.hresult
    if err.excepinfo and err.excepinfo[5]:
        code = err.excepinfo[5]
    return code & 0xFFFFFFFF


def device_name(info):
    return str(info.Properties.Item("Name").Value)


def find_scanner():
    """Return the DeskJet's WIA DeviceInfo, any other scanner as a fallback, or None."""
    manager = win32com.client.Dispatch("WIA.DeviceManager")
    infos = [manager.DeviceInfos.Item(i) for i in range(1, manager.DeviceInfos.Count + 1)]
    scanners = [info for info in infos if info.Type == WIA_SCANNER_DEVICE]
    for info in scanners:
        if "4620" in device_name(info):
            return info
    return scanners[0] if scanners else None


def supported_resolutions(item):
    prop = item.Properties.Item(PROP_X_RESOLUTION)
    if prop.SubType == WIA_SUBTYPE_RANGE:
        return [dpi for dpi in DEFAULT_RESOLUTIONS if prop.SubTypeMin <= dpi <= prop.SubTypeMax]
    values = prop.SubTypeValues
    return sorted(int(values.Item(i)) for i in range(1, values.Count + 1)) or DEFAULT_RESOLUTIONS


def read_device():
    """Name and resolutions of the connected scanner, or (None, []) if there is none."""
    info = find_scanner()
    if info is None:
        return None, []
    device = info.Connect()
    return device_name(info), supported_resolutions(device.Items.Item(1))


def hp_scan_running():
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq HPScan.exe", "/NH"],
        capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return "HPScan.exe" in result.stdout


def describe_error(code):
    message = WIA_ERRORS.get(code, f"The scanner reported an error (0x{code:08X}).")
    if code in (0x80210006, 0x8021000D) and hp_scan_running():
        message += "\n\nHP Scan is open and may be using the scanner. Close HP Scan and try again."
    return message


def set_prop(props, prop_id, value):
    props.Item(prop_id).Value = value


def acquire(bmp_path, dpi, output_type, size):
    """Scan one page from the flatbed into an uncompressed BMP file."""
    info = find_scanner()
    if info is None:
        raise ScanError(WIA_ERRORS[0x80210015])
    device = info.Connect()
    try:
        set_prop(device.Properties, PROP_DOCUMENT_HANDLING_SELECT, WIA_FLATBED)
    except pywintypes.com_error:
        pass  # scanners without a document feeder don't have this property

    item = device.Items.Item(1)
    props = item.Properties
    data_type, bits = OUTPUT_TYPES[output_type]
    set_prop(props, PROP_DATA_TYPE, data_type)
    set_prop(props, PROP_BITS_PER_PIXEL, bits)
    # Resolution goes first: the allowed start/extent ranges are in pixels at the current resolution.
    set_prop(props, PROP_X_RESOLUTION, dpi)
    set_prop(props, PROP_Y_RESOLUTION, dpi)
    set_prop(props, PROP_X_START, 0)
    set_prop(props, PROP_Y_START, 0)
    max_width = props.Item(PROP_X_EXTENT).SubTypeMax
    max_height = props.Item(PROP_Y_EXTENT).SubTypeMax
    if size is None:
        width, height = max_width, max_height
    else:
        width = min(round(size[0] * dpi), max_width)
        height = min(round(size[1] * dpi), max_height)
    set_prop(props, PROP_X_EXTENT, width)
    set_prop(props, PROP_Y_EXTENT, height)

    image = item.Transfer(WIA_FORMAT_BMP)
    image.SaveFile(str(bmp_path))


def acquire_with_retry(bmp_path, dpi, output_type, size, report, attempts=3):
    for attempt in range(1, attempts + 1):
        try:
            acquire(bmp_path, dpi, output_type, size)
            return
        except pywintypes.com_error as err:
            code = wia_error_code(err)
            if code in RETRYABLE_ERRORS and attempt < attempts:
                report(f"Scanner not ready, retrying ({attempt}/{attempts - 1})")
                time.sleep(5)
                continue
            raise ScanError(describe_error(code)) from err


# ---------------------------------------------------------------- image

def enhance_document(img, dpi):
    """Map the paper tone to pure white, deepen the blacks, and sharpen text slightly."""
    hist = img.reduce(4).convert("L").histogram()
    total = sum(hist)
    # The paper is the dominant bright tone. Without one (e.g. a photo) the levels are left alone.
    window = [sum(hist[max(v - 3, 0):v + 4]) for v in range(256)]
    paper = max(range(140, 256), key=window.__getitem__)
    if window[paper] >= total * 0.10:
        cumulative = list(itertools.accumulate(hist))
        black = min(next(v for v, count in enumerate(cumulative) if count >= total * 0.002), 50)
        white = paper - 6
        if white - black > 100:
            scale = 255 / (white - black)
            lut = [max(0, min(255, round((v - black) * scale))) for v in range(256)]
            img = img.point(lut * len(img.getbands()))
    radius = max(0.5, dpi / 300)  # same physical sharpening radius at every resolution
    return img.filter(ImageFilter.UnsharpMask(radius=radius, percent=70, threshold=3))


def save_jpeg(img, path, dpi):
    extra = {}
    if img.mode == "RGB" and SRGB_PROFILE.exists():
        extra["icc_profile"] = SRGB_PROFILE.read_bytes()
    # 4:4:4 (no chroma subsampling) keeps colored text and fine detail sharp.
    img.save(path, "JPEG", quality=JPEG_QUALITY, subsampling=0, dpi=(dpi, dpi), **extra)


def unique_path(folder):
    stem = f"scan_{datetime.now():%Y-%m-%d_%H-%M-%S}"
    path = folder / f"{stem}.jpg"
    for n in itertools.count(2):
        if not path.exists():
            return path
        path = folder / f"{stem}_{n}.jpg"


def run_scan(folder, dpi, output_type, size, item_type, report):
    """Scan, clean up and save one page. Returns the JPEG path and its pixel size."""
    folder.mkdir(parents=True, exist_ok=True)
    bmp_path = Path(tempfile.gettempdir()) / f"hp4620scan_{os.getpid()}_{time.time_ns()}.bmp"
    try:
        report("Scanning")
        acquire_with_retry(bmp_path, dpi, output_type, size, report)
        report("Processing")
        img = Image.open(bmp_path)
        img.load()
    finally:
        bmp_path.unlink(missing_ok=True)

    mode = "RGB" if output_type == "Color" else "L"
    if img.mode != mode:
        img = img.convert(mode)
    if item_type == "Document" and output_type != "Black and White":
        img = enhance_document(img, dpi)

    report("Saving JPEG")
    path = unique_path(folder)
    save_jpeg(img, path, dpi)
    return path, img.size


# ---------------------------------------------------------------- settings

def default_folder():
    try:
        pictures = Path(shell.SHGetKnownFolderPath(shellcon.FOLDERID_Pictures))
    except pywintypes.com_error:
        pictures = Path.home() / "Pictures"
    return pictures / "Scans"


def load_settings():
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_settings(settings):
    try:
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")
    except OSError:
        pass


# ---------------------------------------------------------------- UI

def format_duration(seconds):
    seconds = int(seconds)
    return f"{seconds // 60}:{seconds % 60:02d}"


class ScanApp:
    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.busy = False
        self.phase = ""
        self.started = 0.0
        self.last_scan = None
        self.settings = load_settings()

        root.title(APP_TITLE)
        root.resizable(False, False)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        style = ttk.Style()
        style.configure("Scan.TButton", font=("Segoe UI", 10, "bold"), padding=(20, 6))
        style.configure("Hint.TLabel", foreground="gray40")

        main = ttk.Frame(root, padding=16)
        main.grid(sticky="nsew")
        main.columnconfigure(0, weight=1)

        self.device_label = ttk.Label(main, text="Looking for the scanner…")
        self.device_label.grid(row=0, column=0, sticky="w", pady=(0, 10))

        box = ttk.LabelFrame(main, text="Save as JPEG", padding=(12, 4, 12, 12))
        box.grid(row=1, column=0, sticky="ew")
        box.columnconfigure(0, weight=1)
        self.scan_size = self.dropdown(box, 0, "Scan Size:", list(SCAN_SIZES), self.settings.get("scan_size"))
        self.output_type = self.dropdown(box, 2, "Output Type:", list(OUTPUT_TYPES), self.settings.get("output_type"))
        self.resolution = self.dropdown(box, 4, "Resolution:", [], None)
        self.item_type = self.dropdown(box, 6, "Item Type:", list(ITEM_TYPES), self.settings.get("item_type"))
        self.item_hint = ttk.Label(box, style="Hint.TLabel")
        self.item_hint.grid(row=8, column=0, sticky="w", pady=(4, 0))
        self.set_resolutions(DEFAULT_RESOLUTIONS)

        folder_row = ttk.Frame(main)
        folder_row.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(folder_row, text="Save to:").pack(side="left")
        self.folder = tk.StringVar(value=self.settings.get("folder") or str(default_folder()))
        ttk.Entry(folder_row, textvariable=self.folder, state="readonly").pack(side="left", fill="x", expand=True, padx=6)
        self.browse_button = ttk.Button(folder_row, text="Browse…", command=self.browse)
        self.browse_button.pack(side="left")

        self.info_label = ttk.Label(main, style="Hint.TLabel")
        self.info_label.grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Label(main, style="Hint.TLabel", text="Place the original face down in the corner marked on the glass.").grid(
            row=4, column=0, sticky="w")

        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.grid(row=5, column=0, sticky="ew", pady=(12, 4))
        self.status = ttk.Label(main, text="Ready.")
        self.status.grid(row=6, column=0, sticky="w")

        buttons = ttk.Frame(main)
        buttons.grid(row=7, column=0, sticky="ew", pady=(14, 0))
        self.open_scan_button = ttk.Button(buttons, text="Open Scan", command=self.open_scan, state="disabled")
        self.open_scan_button.pack(side="left")
        ttk.Button(buttons, text="Open Folder", command=self.open_folder).pack(side="left", padx=(6, 0))
        self.scan_button = ttk.Button(buttons, text="Scan", style="Scan.TButton", command=self.start_scan)
        self.scan_button.pack(side="right")
        root.bind("<Return>", lambda _: self.start_scan())

        self.refresh_info()
        self.check_device()
        self.pump()

    def dropdown(self, parent, row, label, values, saved):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(8, 2))
        combo = ttk.Combobox(parent, values=values, state="readonly", width=44)
        combo.grid(row=row + 1, column=0, sticky="ew")
        if values:
            combo.set(saved if saved in values else values[0])
        combo.bind("<<ComboboxSelected>>", lambda _: self.refresh_info())
        return combo

    def set_resolutions(self, resolutions):
        best = max(resolutions)
        labels = [f"{dpi} dpi" + ("  (maximum)" if dpi == best else "") for dpi in resolutions]
        saved = self.settings.get("resolution", best)
        self.resolution.config(values=labels)
        self.resolution.set(labels[resolutions.index(saved if saved in resolutions else best)])

    def selected_dpi(self):
        return int(self.resolution.get().split()[0])

    def refresh_info(self):
        dpi = self.selected_dpi()
        width, height = SCAN_SIZES[self.scan_size.get()] or GLASS_SIZE
        pixels_x, pixels_y = round(width * dpi), round(height * dpi)
        megabytes = pixels_x * pixels_y * OUTPUT_TYPES[self.output_type.get()][1] / 8 / 1e6
        self.info_label.config(
            text=f"{pixels_x:,} × {pixels_y:,} px  ·  {pixels_x * pixels_y / 1e6:.1f} MP  ·  {megabytes:,.0f} MB uncompressed")
        self.item_hint.config(text=ITEM_TYPES[self.item_type.get()])

    # -- background work ------------------------------------------------

    def check_device(self):
        def work():
            pythoncom.CoInitialize()
            try:
                self.events.put(("device", *read_device()))
            except pywintypes.com_error:
                self.events.put(("device", None, []))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=work, daemon=True).start()

    def start_scan(self):
        if self.busy:
            return
        self.settings.update(
            scan_size=self.scan_size.get(), output_type=self.output_type.get(),
            resolution=self.selected_dpi(), item_type=self.item_type.get(), folder=self.folder.get())
        save_settings(self.settings)
        options = dict(
            folder=Path(self.folder.get()), dpi=self.selected_dpi(), output_type=self.output_type.get(),
            size=SCAN_SIZES[self.scan_size.get()], item_type=self.item_type.get())
        self.set_busy(True)
        threading.Thread(target=self.scan_worker, args=(options,), daemon=True).start()

    def scan_worker(self, options):
        pythoncom.CoInitialize()
        try:
            path, size = run_scan(report=lambda text: self.events.put(("phase", text)), **options)
            self.events.put(("done", path, size))
        except ScanError as err:
            self.events.put(("error", str(err)))
        except Exception as err:  # keep the window alive whatever happens
            self.events.put(("error", f"Unexpected error: {type(err).__name__}: {err}"))
        finally:
            pythoncom.CoUninitialize()

    def pump(self):
        while True:
            try:
                event, *args = self.events.get_nowait()
            except queue.Empty:
                break
            getattr(self, f"on_{event}")(*args)
        if self.busy:
            self.status.config(text=f"{self.phase}…  {format_duration(time.monotonic() - self.started)}")
        self.root.after(200, self.pump)

    def on_device(self, name, resolutions):
        if name:
            self.device_label.config(text=f"●  {name}  —  ready", foreground="#1a7f37")
            if resolutions:
                self.set_resolutions(resolutions)
                self.refresh_info()
        else:
            self.device_label.config(text="●  Scanner not found — check power and the USB cable", foreground="#c62828")

    def on_phase(self, text):
        self.phase = text

    def on_done(self, path, size):
        elapsed = format_duration(time.monotonic() - self.started)
        self.set_busy(False)
        self.last_scan = path
        self.open_scan_button.config(state="normal")
        megabytes = path.stat().st_size / 1e6
        self.status.config(text=f"Saved {path.name}  ({size[0]:,} × {size[1]:,} px, {megabytes:.1f} MB) in {elapsed}")

    def on_error(self, message):
        self.set_busy(False)
        self.status.config(text="Scan failed.")
        messagebox.showerror(APP_TITLE, message)

    def set_busy(self, busy):
        self.busy = busy
        combo_state = "disabled" if busy else "readonly"
        for combo in (self.scan_size, self.output_type, self.resolution, self.item_type):
            combo.config(state=combo_state)
        self.scan_button.config(state="disabled" if busy else "normal")
        self.browse_button.config(state="disabled" if busy else "normal")
        if busy:
            self.phase = "Starting"
            self.started = time.monotonic()
            self.progress.start(12)
        else:
            self.progress.stop()

    # -- buttons --------------------------------------------------------

    def browse(self):
        folder = filedialog.askdirectory(initialdir=self.folder.get(), title="Save scans to")
        if folder:
            self.folder.set(str(Path(folder)))

    def open_scan(self):
        if self.last_scan and self.last_scan.exists():
            os.startfile(self.last_scan)

    def open_folder(self):
        if self.last_scan and self.last_scan.exists():
            subprocess.Popen(f'explorer /select,"{self.last_scan}"')
            return
        folder = Path(self.folder.get())
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(folder)

    def on_close(self):
        if self.busy and not messagebox.askyesno(APP_TITLE, "A scan is still running. Quit anyway?"):
            return
        self.root.destroy()


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # crisp text on high-DPI screens
    except (AttributeError, OSError):
        pass
    root = tk.Tk()
    ScanApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
