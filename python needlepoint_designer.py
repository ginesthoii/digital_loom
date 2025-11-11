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
    ("945",  "Tawny",            (251, 213, 187)),
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
    return best  # (code, name)


# =============================
# Core image processing
# =============================

def quantize_image(img, out_w, out_h, colors):
    img_small = img.resize((out_w, out_h), Image.LANCZOS).convert("RGB")
    img_p = img_small.convert("P", palette=Image.ADAPTIVE, colors=colors)
    img_q = img_p.convert("RGB")
    return img_q


def draw_color_grid(base_grid, cell_px=20, grid=True, numbering=True):
    """Render a color chart image from a grid of RGB pixels.
    Includes optional gridlines and top/left numbering every 10 lines.
    """
    gw, gh = base_grid.size
    title_h = 28
    border = 40
    chart_w = gw * cell_px
    chart_h = gh * cell_px
    out = Image.new("RGB", (chart_w + border*2, chart_h + border*2 + title_h), (255,255,255))
    draw = ImageDraw.Draw(out)

    font, font_sm, font_title = load_fonts()

    # Title
    draw.text((border, 8), "Needlepoint Pattern — Color", fill=(0,0,0), font=font_title)

    grid_x0 = border
    grid_y0 = border + title_h

    # Cells
    for y in range(gh):
        for x in range(gw):
            c = base_grid.getpixel((x,y))
            x0 = grid_x0 + x*cell_px
            y0 = grid_y0 + y*cell_px
            draw.rectangle([x0, y0, x0+cell_px-1, y0+cell_px-1], fill=c)

    # Gridlines
    if grid:
        grid_color = (0,0,0)
        for x in range(gw+1):
            xg = grid_x0 + x*cell_px
            w = 2 if (x % 10 == 0) else 1
            draw.line([(xg, grid_y0), (xg, grid_y0+chart_h)], fill=grid_color, width=w)
        for y in range(gh+1):
            yg = grid_y0 + y*cell_px
            w = 2 if (y % 10 == 0) else 1
            draw.line([(grid_x0, yg), (grid_x0+chart_w, yg)], fill=grid_color, width=w)

    # Numbering top/left
    if numbering:
        fill = (0,0,0)
        for x in range(0, gw+1, 10):
            xg = grid_x0 + x*cell_px
            draw.text((xg+2, grid_y0-18), str(x), fill=fill, font=font_sm)
        for y in range(0, gh+1, 10):
            yg = grid_y0 + y*cell_px
            draw.text((grid_x0-28, yg-10), str(y), fill=fill, font=font_sm)

    return out


def draw_symbol_grid(base_grid, palette_order, symbol_map, cell_px=20, numbering=True):
    """Render a black-and-white symbolic chart.
    Each color index is drawn as a symbol centered in a cell.
    """
    gw, gh = base_grid.size
    title_h = 28
    border = 40
    chart_w = gw * cell_px
    chart_h = gh * cell_px
    out = Image.new("RGB", (chart_w + border*2, chart_h + border*2 + title_h), (255,255,255))
    draw = ImageDraw.Draw(out)
    font, font_sm, font_title = load_fonts()
    sym_font = best_symbol_font(cell_px)

    draw.text((border, 8), "Needlepoint Pattern — Symbols (B/W)", fill=(0,0,0), font=font_title)

    grid_x0 = border
    grid_y0 = border + title_h

    # Light cell background
    bg = (252,252,252)
    for y in range(gh):
        for x in range(gw):
            x0 = grid_x0 + x*cell_px
            y0 = grid_y0 + y*cell_px
            draw.rectangle([x0, y0, x0+cell_px-1, y0+cell_px-1], fill=bg)

    # Thin gridlines + bold every 10
    grid_color = (0,0,0)
    for x in range(gw+1):
        xg = grid_x0 + x*cell_px
        w = 2 if (x % 10 == 0) else 1
        draw.line([(xg, grid_y0), (xg, grid_y0+chart_h)], fill=grid_color, width=w)
    for y in range(gh+1):
        yg = grid_y0 + y*cell_px
        w = 2 if (y % 10 == 0) else 1
        draw.line([(grid_x0, yg), (grid_x0+chart_w, yg)], fill=grid_color, width=w)

    # Symbols
    for y in range(gh):
        for x in range(gw):
            rgb = base_grid.getpixel((x,y))
            try:
                idx = palette_order.index(rgb)
            except ValueError:
                idx = 0
            sym = symbol_map[idx]
            x0 = grid_x0 + x*cell_px
            y0 = grid_y0 + y*cell_px
            draw_symbol_centered(draw, (x0, y0, x0+cell_px-1, y0+cell_px-1), sym, sym_font)

    # Numbering top/left
    if numbering:
        for x in range(0, gw+1, 10):
            xg = grid_x0 + x*cell_px
            draw.text((xg+2, grid_y0-18), str(x), fill=(0,0,0), font=font_sm)
        for y in range(0, gh+1, 10):
            yg = grid_y0 + y*cell_px
            draw.text((grid_x0-28, yg-10), str(y), fill=(0,0,0), font=font_sm)

    return out


def load_fonts():
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 16)
        font_sm = ImageFont.truetype("DejaVuSans.ttf", 14)
        font_title = ImageFont.truetype("DejaVuSans.ttf", 20)
    except:
        font = ImageFont.load_default()
        font_sm = ImageFont.load_default()
        font_title = ImageFont.load_default()
    return font, font_sm, font_title


def best_symbol_font(cell_px):
    size = max(12, int(cell_px*0.8))
    try:
        return ImageFont.truetype("DejaVuSans.ttf", size)
    except:
        return ImageFont.load_default()


def draw_symbol_centered(draw, box, text, font):
    x0, y0, x1, y1 = box
    w = x1 - x0
    h = y1 - y0
    tw, th = draw.textsize(text, font=font)
    tx = x0 + (w - tw)//2
    ty = y0 + (h - th)//2
    draw.text((tx, ty), text, fill=(0,0,0), font=font)


# =============================
# Tiling (Letter paper) and export
# =============================

DPI = 300
LETTER_IN = (8.5, 11.0)
MARGIN_IN = 0.5
OVERLAP_IN = 0.25


def tile_to_pages(chart_img, paper_in=LETTER_IN, margin_in=MARGIN_IN, overlap_in=OVERLAP_IN, footer_text=None):
    pw_px = int(paper_in[0] * DPI)
    ph_px = int(paper_in[1] * DPI)
    margin_px = int(margin_in * DPI)
    overlap_px = int(overlap_in * DPI)

    content_w = pw_px - 2*margin_px
    content_h = ph_px - 2*margin_px

    cw, ch = chart_img.size

    pages = []
    y = 0
    row = 0
    while y < ch:
        x = 0
        col = 0
        h_end = min(y + content_h, ch)
        while x < cw:
            w_end = min(x + content_w, cw)
            crop = chart_img.crop((x, y, w_end, h_end))
            page = Image.new("RGB", (pw_px, ph_px), (255,255,255))
            page.paste(crop, (margin_px, margin_px))

            # Footer page number / label
            if footer_text:
                draw = ImageDraw.Draw(page)
                font, font_sm, font_title = load_fonts()
                draw.text((margin_px, ph_px - margin_px + 4 - 20), footer_text, fill=(0,0,0), font=font_sm)

            pages.append(page)
            if w_end == cw:
                break
            x = w_end - overlap_px
            col += 1
        if h_end == ch:
            break
        y = h_end - overlap_px
        row += 1

    return pages


# =============================
# Legend helpers
# =============================

def build_palette_counts(img):
    pixels = list(img.getdata())
    counts = Counter(pixels)
    palette = sorted(counts.keys(), key=lambda c: (-counts[c], c))
    return palette, counts


def make_symbol_map(num_colors):
    flat = []
    for group in SYMBOLS:
        flat.extend(group)
    if num_colors > len(flat):
        # fall back to letters cycling
        extra = [chr(ord('a') + (i % 26)) for i in range(num_colors - len(flat))]
        flat.extend(extra)
    return flat[:num_colors]


def write_csv_legend(path, palette, counts, include_dmc=True, include_symbols=True, symbol_map=None):
    with open(path, 'w', encoding='utf-8') as f:
        headers = ["count","hex","r","g","b"]
        if include_symbols:
            headers.insert(0, "symbol")
        if include_dmc:
            headers += ["dmc_code","dmc_name"]
        f.write(",".join(headers) + "\n")
        for i, rgb in enumerate(palette):
            row = []
            if include_symbols:
                row.append(symbol_map[i])
            row.extend([str(counts[rgb]), rgb_to_hex(rgb), str(rgb[0]), str(rgb[1]), str(rgb[2])])
            if include_dmc:
                code, name = nearest_dmc(rgb)
                row.extend([code, name])
            f.write(",".join(row) + "\n")


# =============================
# GUI App
# =============================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Needlepoint Pattern Designer (Color + Symbols, Tiled PDFs)")

        # State
        self.img_path = StringVar(value="")
        self.grid_max = IntVar(value=160)       # long side stitches
        self.color_count = IntVar(value=40)
        self.keep_aspect = BooleanVar(value=True)
        self.cell_px = IntVar(value=18)
        self.include_dmc = BooleanVar(value=True)
        self.export_color = BooleanVar(value=True)
        self.export_symbols = BooleanVar(value=True)

        # Layout
        frm = ttk.Frame(root, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        file_row = ttk.Frame(frm)
        file_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,8))
        ttk.Label(file_row, text="Image:").pack(side="left")
        self.entry = ttk.Entry(file_row, textvariable=self.img_path, width=50)
        self.entry.pack(side="left", padx=6)
        ttk.Button(file_row, text="Browse", command=self.browse).pack(side="left")

        # Options
        opt = ttk.LabelFrame(frm, text="Options", padding=8)
        opt.grid(row=1, column=0, sticky="nsew")

        ttk.Label(opt, text="Max stitches (longer side):").grid(row=0, column=0, sticky="w")
        ttk.Scale(opt, from_=60, to=300, variable=self.grid_max, orient="horizontal").grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Label(opt, textvariable=self.grid_max, width=4).grid(row=0, column=2)

        ttk.Label(opt, text="Color count:").grid(row=1, column=0, sticky="w")
        ttk.Scale(opt, from_=2, to=60, variable=self.color_count, orient="horizontal").grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Label(opt, textvariable=self.color_count, width=4).grid(row=1, column=2)

        ttk.Label(opt, text="Pixels per stitch (export):").grid(row=2, column=0, sticky="w")
        ttk.Scale(opt, from_=10, to=30, variable=self.cell_px, orient="horizontal").grid(row=2, column=1, sticky="ew", padx=6)
        ttk.Label(opt, textvariable=self.cell_px, width=4).grid(row=2, column=2)

        ttk.Checkbutton(opt, text="Keep aspect ratio", variable=self.keep_aspect).grid(row=3, column=0, sticky="w", pady=(6,0))
        ttk.Checkbutton(opt, text="Approximate DMC in legend", variable=self.include_dmc).grid(row=3, column=1, sticky="w", pady=(6,0))
        ttk.Checkbutton(opt, text="Export Color PDF", variable=self.export_color).grid(row=4, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Export Symbols PDF", variable=self.export_symbols).grid(row=4, column=1, sticky="w")

        for i in range(3):
            opt.columnconfigure(i, weight=1)

        # Preview
        prev = ttk.LabelFrame(frm, text="Preview (reduced colors)", padding=8)
        prev.grid(row=1, column=1, sticky="nsew", padx=(10,0))
        self.preview_label = ttk.Label(prev)
        self.preview_label.pack(fill="both", expand=True)

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8,0))
        ttk.Button(btns, text="Generate Preview", command=self.generate_preview).pack(side="left")
        ttk.Button(btns, text="Export PDFs + CSV", command=self.export_all).pack(side="left", padx=8)

        frm.columnconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(1, weight=1)

        self.preview_imgtk = None
        self.grid_img = None

    def browse(self):
        p = filedialog.askopenfilename(filetypes=[("Images","*.png;*.jpg;*.jpeg;*.webp;*.bmp")])
        if p:
            self.img_path.set(p)

    def compute_out_dims(self, base):
        max_st = int(self.grid_max.get())
        if self.keep_aspect.get():
            w, h = base.size
            if w >= h:
                out_w = max_st
                out_h = max(1, round(h * (max_st / w)))
            else:
                out_h = max_st
                out_w = max(1, round(w * (max_st / h)))
        else:
            out_w = out_h = max_st
        return out_w, out_h

    def generate_preview(self):
        if not os.path.isfile(self.img_path.get()):
            self._toast("Select a valid image.")
            return
        base = Image.open(self.img_path.get()).convert("RGB")
        out_w, out_h = self.compute_out_dims(base)
        colors = int(self.color_count.get())
        self.grid_img = quantize_image(base, out_w, out_h, colors)

        scale = max(2, min(6, 600 // max(out_w, out_h)))
        prev = self.grid_img.resize((out_w*scale, out_h*scale), Image.NEAREST)
        # Light preview grid every 10
        draw = ImageDraw.Draw(prev)
        for x in range(0, out_w+1, 10):
            X = x*scale
            draw.line([(X,0),(X,out_h*scale)], fill=(0,0,0), width=2)
        for y in range(0, out_h+1, 10):
            Y = y*scale
            draw.line([(0,Y),(out_w*scale,Y)], fill=(0,0,0), width=2)

        self.preview_imgtk = ImageTk.PhotoImage(prev)
        self.preview_label.configure(image=self.preview_imgtk)
        self._toast(f"Preview: {out_w}×{out_h} stitches, {colors} colors.")

    def export_all(self):
        if self.grid_img is None:
            self._toast("Generate a preview first.")
            return

        iname = os.path.splitext(os.path.basename(self.img_path.get()))[0]
        w, h = self.grid_img.size
        c = int(self.color_count.get())
        base = f"{iname}_{w}x{h}_{c}c"

        # Palette + counts for legend + symbols
        palette, counts = build_palette_counts(self.grid_img)
        symbol_map = make_symbol_map(len(palette))

        # CSV legend
        csv_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=base + "_legend.csv",
            filetypes=[("CSV","*.csv")]
        )
        if not csv_path:
            self._toast("Export canceled.")
            return
        write_csv_legend(csv_path, palette, counts, include_dmc=self.include_dmc.get(), include_symbols=True, symbol_map=symbol_map)

        # Build color chart and symbol chart images (full charts before tiling)
        cell_px = int(self.cell_px.get())
        color_chart = draw_color_grid(self.grid_img, cell_px=cell_px, grid=True, numbering=True)
        symbols_chart = draw_symbol_grid(self.grid_img, palette, symbol_map, cell_px=cell_px, numbering=True)

        pages_color = []
        pages_symbols = []

        # Tile to pages
        if self.export_color.get():
            pages_color = tile_to_pages(color_chart, footer_text="Color chart — Letter, 1/4 in overlap")
        if self.export_symbols.get():
            pages_symbols = tile_to_pages(symbols_chart, footer_text="Symbols chart — Letter, 1/4 in overlap")

        # Save PDFs
        if pages_color:
            pdf_path_color = os.path.splitext(csv_path)[0].replace("_legend", "_color") + ".pdf"
            pages_color[0].save(pdf_path_color, save_all=True, append_images=pages_color[1:], resolution=DPI)
        if pages_symbols:
            pdf_path_symbols = os.path.splitext(csv_path)[0].replace("_legend", "_symbols") + ".pdf"
            pages_symbols[0].save(pdf_path_symbols, save_all=True, append_images=pages_symbols[1:], resolution=DPI)

        msg_parts = [csv_path]
        if self.export_color.get():
            msg_parts.append(pdf_path_color)
        if self.export_symbols.get():
            msg_parts.append(pdf_path_symbols)
        self._toast("Saved: " + " | ".join(msg_parts))

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
