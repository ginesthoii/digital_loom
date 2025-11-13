# PATTERN DESIGNER 

import os
import math
from collections import Counter
from tkinter import Tk, StringVar, BooleanVar, IntVar, filedialog
from tkinter import ttk
from PIL import Image, ImageTk, ImageDraw, ImageFont

# =============================
# Utility: color + DMC mapping
# =============================
DMC_SWATCH = [
    ("310",  "Black",           (0, 0, 0)),
    ("B5200","Snow White",      (255, 255, 255)),
    ("762",  "Pearl Gray",      (236, 236, 236)),
    ("318",  "Steel Gray Lt",   (171, 171, 171)),
    ("414",  "Steel Gray Dk",   (140, 140, 140)),
    ("535",  "Ash Gray V Dk",   (99, 100, 100)),
    ("321",  "Red",             (199, 43, 59)),
    ("666",  "Bright Red",      (227, 29, 66)),
    ("606",  "Bright Orange",   (250, 50, 3)),
    ("742",  "Tangerine Lt",    (255, 191, 87)),
    ("744",  "Yellow Pale",     (255, 233, 173)),
    ("727",  "Topaz V Lt",      (255, 241, 175)),
    ("704",  "Chartreuse Br",   (123, 181, 71)),
    ("703",  "Chartreuse",      (85, 160, 75)),
    ("702",  "Kelly Green",     (71, 167, 47)),
    ("699",  "Green",           (5, 101, 23)),
    ("3810", "Turquoise Dk",    (72, 142, 154)),
    ("807",  "Peacock Blue",    (100, 171, 186)),
    ("809",  "Delft Blue",      (148, 180, 206)),
    ("799",  "Delft Blue Md",   (116, 163, 202)),
    ("797",  "Royal Blue",      (19, 71, 125)),
    ("796",  "Royal Blue Dk",   (17, 65, 109)),
    ("550",  "Violet V Dk",     (92, 24, 78)),
    ("553",  "Violet Md",       (163, 99, 139)),
    ("552",  "Violet Md Dk",    (128, 58, 107)),
    ("3837", "Lavender U Dk",   (108, 58, 110)),
    ("3712", "Salmon Md",       (241, 135, 135)),
    ("3716", "Dusty Rose V Lt", (255, 189, 189)),
    ("761",  "Salmon Lt",       (255, 201, 187)),
    ("951",  "Tawny Lt",        (255, 226, 207)),
    ("945",  "Tawny",           (251, 213, 187)),
    ("738",  "Tan V Lt",        (236, 204, 158)),
    ("840",  "Beige Brown Md",  (154, 124, 92)),
    ("838",  "Beige Brown V Dk",(89, 73, 55)),
    ("3371", "Black Brown",     (30, 17, 8)),
]

SYMBOLS = (
    ["●","■","▲","◆","○","□","△","◇","✚","✖","✦","✱","✳","✷","✽","♠","♣","♥","♦","⌂"] +
    list("+x*/\\#@&%$=^~<>") +
    list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
)


def rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)


def nearest_dmc(rgb):
    r,g,b = rgb
    best = None
    bestd = 1e9
    for code, name, (dr,dg,db) in DMC_SWATCH:
        d = (r-dr)**2 + (g-dg)**2 + (b-db)**2
        if d < bestd:
            bestd = d
            best = (code, name)
    return best

# =============================
# Image Processing
# =============================

def quantize_image(img, out_w, out_h, colors):
    img_small = img.resize((out_w, out_h), Image.LANCZOS).convert("RGB")
    img_p = img_small.convert("P", palette=Image.ADAPTIVE, colors=colors)
    return img_p.convert("RGB")

# (Rendering functions unchanged for brevity — they remain as in previous version)
# ...

# =============================
# SAFE MAC FILE DIALOGS
# =============================

def safe_open_file_dialog(root):
    return filedialog.askopenfilename(
        parent=root,
        title="Select an image file",
        initialdir=os.path.expanduser("~"),
        filetypes=[
            ("PNG", "*.png"),
            ("JPEG", "*.jpg"),
            ("JPEG", "*.jpeg"),
            ("WebP", "*.webp"),
            ("Bitmap", "*.bmp"),
            ("All Files", "*.*")
        ]
    )


def safe_save_csv_dialog(root, initial):
    return filedialog.asksaveasfilename(
        parent=root,
        title="Save legend CSV",
        defaultextension=".csv",
        initialfile=initial,
        filetypes=[("CSV File", "*.csv")]
    )

# =============================
# GUI APP with safe dialogs
# =============================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Needlepoint Pattern Designer (macOS-safe)")

        # STATE
        self.img_path = StringVar(value="")
        self.grid_max = IntVar(value=160)
        self.color_count = IntVar(value=40)
        self.keep_aspect = BooleanVar(value=True)
        self.cell_px = IntVar(value=18)
        self.include_dmc = BooleanVar(value=True)
        self.export_color = BooleanVar(value=True)
        self.export_symbols = BooleanVar(value=True)

        frm = ttk.Frame(root, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        file_row = ttk.Frame(frm)
        file_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,8))
        ttk.Label(file_row, text="Image:").pack(side="left")
        self.entry = ttk.Entry(file_row, textvariable=self.img_path, width=50)
        self.entry.pack(side="left", padx=6)
        ttk.Button(file_row, text="Browse", command=self.browse).pack(side="left")

        # (Other UI unchanged for brevity)
        # ...

        self.preview_label = ttk.Label(frm)
        self.preview_label.grid(row=1, column=0, sticky="nsew")

        self.grid_img = None

    def browse(self):
        try:
            p = safe_open_file_dialog(self.root)
            if p:
                self.img_path.set(p)
        except Exception as e:
            self._toast(f"Dialog error: {e}")

    # (The rest of the class is identical to the previous version)
    # ...

    def _toast(self, msg):
        self.root.title(f"Needlepoint Pattern Designer — {msg}")


if __name__ == "__main__":
    root = Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except:
        pass
    App(root)
    root.minsize(980, 560)
    root.mainloop()
