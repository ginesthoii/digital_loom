import os
import math
from collections import Counter
from tkinter import Tk, StringVar, BooleanVar, IntVar, filedialog
from tkinter import ttk
from PIL import Image, ImageTk, ImageDraw, ImageFont

# ============================================================
#  DMC COLOR SWATCH + SYMBOL SET
# ============================================================

DMC_SWATCH = [
    ("310","Black",(0,0,0)),
    ("B5200","Snow White",(255,255,255)),
    ("762","Pearl Gray",(236,236,236)),
    ("318","Steel Gray Lt",(171,171,171)),
    ("414","Steel Gray Dk",(140,140,140)),
    ("535","Ash Gray V Dk",(99,100,100)),
    ("321","Red",(199,43,59)),
    ("666","Bright Red",(227,29,66)),
    ("606","Bright Orange",(250,50,3)),
    ("742","Tangerine Lt",(255,191,87)),
    ("744","Yellow Pale",(255,233,173)),
    ("704","Chartreuse Br",(123,181,71)),
    ("703","Chartreuse",(85,160,75)),
    ("702","Kelly Green",(71,167,47)),
    ("699","Green",(5,101,23)),
    ("3810","Turquoise Dk",(72,142,154)),
    ("807","Peacock Blue",(100,171,186)),
    ("799","Delft Blue Md",(116,163,202)),
    ("797","Royal Blue",(19,71,125)),
    ("550","Violet V Dk",(92,24,78)),
    ("553","Violet Md",(163,99,139)),
    ("552","Violet Md Dk",(128,58,107)),
    ("3712","Salmon Md",(241,135,135)),
    ("761","Salmon Lt",(255,201,187)),
    ("951","Tawny Lt",(255,226,207)),
    ("945","Tawny",(251,213,187)),
    ("738","Tan V Lt",(236,204,158)),
    ("840","Beige Brown Md",(154,124,92)),
    ("838","Beige Brown V Dk",(89,73,55)),
    ("3371","Black Brown",(30,17,8)),
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

# ============================================================
#  IMAGE REDUCTION
# ============================================================

def quantize_image(img, out_w, out_h, colors):
    img_small = img.resize((out_w, out_h), Image.LANCZOS).convert("RGB")
    img_p = img_small.convert("P", palette=Image.ADAPTIVE, colors=colors)
    return img_p.convert("RGB")

# ============================================================
#  SAFE FILE DIALOGS FOR MACOS
# ============================================================

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

# ============================================================
#  CANVAS RENDERING (SINGLE BIG IMAGE)
# ============================================================

def render_pattern_image(arr, cell_px, numbered=True, symbols=False, palette_map=None):
    h = len(arr)
    w = len(arr[0])
    img = Image.new("RGB", (w*cell_px, h*cell_px), (255,255,255))
    draw = ImageDraw.Draw(img)

    font = ImageFont.load_default()

    for y in range(h):
        for x in range(w):
            if symbols:
                sym = palette_map[arr[y][x]]
                draw.rectangle(
                    (x*cell_px, y*cell_px, x*cell_px+cell_px, y*cell_px+cell_px),
                    fill=(255,255,255),
                    outline=(200,200,200)
                )
                tw,th = draw.textsize(sym, font=font)
                draw.text(
                    (x*cell_px+(cell_px-tw)//2, y*cell_px+(cell_px-th)//2),
                    sym, fill=(0,0,0), font=font
                )
            else:
                draw.rectangle(
                    (x*cell_px, y*cell_px, x*cell_px+cell_px, y*cell_px+cell_px),
                    fill=arr[y][x],
                    outline=(200,200,200)
                )

    if numbered:
        # every 10th row
        for y in range(0,h,10):
            draw.line((0, y*cell_px, w*cell_px, y*cell_px), fill=(0,0,0))
            draw.text((2, y*cell_px+2), str(y), fill=(0,0,0), font=font)

        # every 10th col
        for x in range(0,w,10):
            draw.line((x*cell_px, 0, x*cell_px, h*cell_px), fill=(0,0,0))
            draw.text((x*cell_px+2, 2), str(x), fill=(0,0,0), font=font)

    return img

# ============================================================
#  PDF TILING (Letter, 1/4-inch overlap)
# ============================================================

def export_tiled_pdf(big_img, filename):
    dpi = 300
    letter_w = 2550  # 8.5 * 300
    letter_h = 3300  # 11 * 300

    overlap = int(0.25 * dpi)  # 1/4 inch

    w,h = big_img.size

    pages = []
    y0 = 0
    while y0 < h:
        x0 = 0
        y1 = min(y0 + letter_h + overlap, h)
        while x0 < w:
            x1 = min(x0 + letter_w + overlap, w)
            crop = big_img.crop((x0, y0, x1, y1))
            pages.append(crop)
            x0 += letter_w
        y0 += letter_h

    if pages:
        pages[0].save(filename, save_all=True, append_images=pages[1:])

# ============================================================
#  GUI APP
# ============================================================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Needlepoint Pattern Designer — macOS Safe")

        self.img_path = StringVar(value="")
        self.grid_max = IntVar(value=160)
        self.color_count = IntVar(value=40)
        self.cell_px = IntVar(value=18)
        self.keep_aspect = BooleanVar(value=True)
        self.include_dmc = BooleanVar(value=True)
        self.export_color = BooleanVar(value=True)
        self.export_symbols = BooleanVar(value=True)

        frm = ttk.Frame(root, padding=12)
        frm.grid(row=0,column=0,sticky="nsew")

        # FILE
        row = ttk.Frame(frm)
        row.grid(row=0,column=0,columnspan=2,sticky="ew")
        ttk.Label(row,text="Image:").pack(side="left")
        ttk.Entry(row, textvariable=self.img_path, width=50).pack(side="left", padx=6)
        ttk.Button(row,text="Browse",command=self.browse).pack(side="left")

        # SETTINGS
        sfrm = ttk.Frame(frm)
        sfrm.grid(row=1,column=0,sticky="nw")

        ttk.Label(sfrm,text="Max Stitches (long side):").grid(row=0,column=0,sticky="w")
        ttk.Entry(sfrm,textvariable=self.grid_max,width=8).grid(row=0,column=1)

        ttk.Label(sfrm,text="Color Count:").grid(row=1,column=0,sticky="w")
        ttk.Entry(sfrm,textvariable=self.color_count,width=8).grid(row=1,column=1)

        ttk.Label(sfrm,text="Pixels per Stitch:").grid(row=2,column=0,sticky="w")
        ttk.Entry(sfrm,textvariable=self.cell_px,width=8).grid(row=2,column=1)

        ttk.Checkbutton(sfrm,text="Keep aspect ratio",variable=self.keep_aspect).grid(row=3,column=0,sticky="w")
        ttk.Checkbutton(sfrm,text="Include DMC match",variable=self.include_dmc).grid(row=4,column=0,sticky="w")
        ttk.Checkbutton(sfrm,text="Export Color Chart",variable=self.export_color).grid(row=5,column=0,sticky="w")
        ttk.Checkbutton(sfrm,text="Export Symbol Chart",variable=self.export_symbols).grid(row=6,column=0,sticky="w")

        ttk.Button(sfrm,text="Generate Preview",command=self.generate_preview).grid(row=7,column=0,pady=10)
        ttk.Button(sfrm,text="Export PDFs + CSV",command=self.export_all).grid(row=8,column=0,pady=10)

        # PREVIEW AREA
        self.preview_label = ttk.Label(frm)
        self.preview_label.grid(row=1,column=1,sticky="nsew")

        self.grid_arr = None

    # ---------------------------
    def browse(self):
        try:
            p = safe_open_file_dialog(self.root)
            if p:
                self.img_path.set(p)
        except Exception as e:
            self.root.title(f"Dialog Error: {e}")

    # ---------------------------
    def generate_preview(self):
        path = self.img_path.get()
        if not os.path.exists(path):
            self.root.title("Image not found.")
            return

        img = Image.open(path).convert("RGB")

        maxs = self.grid_max.get()
        w,h = img.size

        if self.keep_aspect.get():
            if w >= h:
                out_w = maxs
                out_h = int(h * (maxs / w))
            else:
                out_h = maxs
                out_w = int(w * (maxs / h))
        else:
            out_w = out_h = maxs

        colors = self.color_count.get()
        small = quantize_image(img, out_w, out_h, colors)

        arr = [[small.getpixel((x,y)) for x in range(out_w)] for y in range(out_h)]
        self.grid_arr = arr

        cell_px = self.cell_px.get()
        prev = render_pattern_image(arr, cell_px, numbered=True, symbols=False)

        prev_small = prev.resize((prev.width//2, prev.height//2))
        tkimg = ImageTk.PhotoImage(prev_small)
        self.preview_label.configure(image=tkimg)
        self.preview_label.image = tkimg

        self.root.title(f"Preview {out_w} x {out_h} stitches")

    # ---------------------------
    def export_all(self):
        if self.grid_arr is None:
            self.root.title("Generate preview first.")
            return

        arr = self.grid_arr

        # Build palette index
        flat = [tuple(px) for row in arr for px in row]
        counts = Counter(flat)
        sorted_cols = [c for c,k in counts.most_common()]

        # Map colors to symbols
        palette_map = {}
        for i, col in enumerate(sorted_cols):
            palette_map[col] = SYMBOLS[i]

        base = os.path.splitext(os.path.basename(self.img_path.get()))[0]
        outdir = os.path.dirname(self.img_path.get())

        # LEGEND CSV
        csv_path = safe_save_csv_dialog(self.root, base + "_legend.csv")
        if csv_path:
            with open(csv_path,"w") as f:
                f.write("symbol,r,g,b,hex,dmc_code,dmc_name\n")
                for col in sorted_cols:
                    sym = palette_map[col]
                    r,g,b = col
                    hx = rgb_to_hex(col)
                    dmc_code, dmc_name = nearest_dmc(col) if self.include_dmc.get() else ("","")
                    f.write(f"{sym},{r},{g},{b},{hx},{dmc_code},{dmc_name}\n")

        # EXPORT PDFs
        cell_px = self.cell_px.get()

        # Color chart
        if self.export_color.get():
            big = render_pattern_image(arr, cell_px, numbered=True, symbols=False)
            export_tiled_pdf(big, os.path.join(outdir, base + "_color.pdf"))

        # Symbols
        if self.export_symbols.get():
            bigs = render_pattern_image(arr, cell_px, numbered=True, symbols=True, palette_map=palette_map)
            export_tiled_pdf(bigs, os.path.join(outdir, base + "_symbols.pdf"))

        self.root.title("Export complete.")

# ============================================================
#  MAIN
# ============================================================

if __name__ == "__main__":
    root = Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except:
        pass
    App(root)
    root.minsize(950, 600)
    root.mainloop()
