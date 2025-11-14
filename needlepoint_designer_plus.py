import os
import math
import csv
from collections import Counter
from tkinter import Tk, StringVar, BooleanVar, IntVar, filedialog
from tkinter import ttk
from PIL import Image, ImageTk, ImageDraw, ImageFont

# ============================================================
#  GLOBAL DMC PALETTE LOADER
# ============================================================

DMC_PALETTE = []  # list of dicts: number, name, type, r, g, b

def load_dmc_palette(csv_path="dmc_palette_full.csv"):
    """Load full DMC palette (all families) from CSV if present."""
    global DMC_PALETTE
    DMC_PALETTE = []
    if not os.path.exists(csv_path):
        return

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                num = str(row["number"]).strip()
                name = str(row.get("name", "")).strip()
                r = int(float(row["r"]))
                g = int(float(row["g"]))
                b = int(float(row["b"]))
                t = str(row.get("type", "regular")).strip().lower() or "regular"
                DMC_PALETTE.append({
                    "number": num,
                    "name": name,
                    "type": t,
                    "r": r,
                    "g": g,
                    "b": b,
                })
            except Exception:
                # skip bad rows silently
                continue

def nearest_dmc(rgb, allow_specialty=True):
    """
    Find nearest DMC color in loaded palette.
    If allow_specialty=False → only 'regular' rows are used.
    Returns (number, name, type) or ("","","") if no palette loaded.
    """
    if not DMC_PALETTE:
        return ("", "", "")

    r, g, b = rgb
    best = None
    bestd = 1e18

    for row in DMC_PALETTE:
        if (not allow_specialty) and row["type"] != "regular":
            continue
        dr = r - row["r"]
        dg = g - row["g"]
        db = b - row["b"]
        d = dr*dr + dg*dg + db*db
        if d < bestd:
            bestd = d
            best = row

    if best is None:
        return ("", "", "")
    return (best["number"], best["name"], best["type"])

def rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)

# ============================================================
#  BASIC SYMBOL SET (UNCHANGED)
# ============================================================

SYMBOLS = (
    ["●","■","▲","◆","○","□","△","◇","✚","✖","✦","✱","✳","✷","✽","♠","♣","♥","♦","⌂"] +
    list("+x*/\\#@&%$=^~<>") +
    list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
)

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
            ("JPEG", "*.jpg;*.jpeg"),
            ("WebP", "*.webp"),
            ("Bitmap", "*.bmp"),
            ("All Files", "*.*")
        ]
    )

def safe_save_base_dialog(root, initial_base):
    """Ask once, then derive _color.pdf, _symbols.pdf, _legend.csv, etc."""
    path = filedialog.asksaveasfilename(
        parent=root,
        title="Save pattern (base name)",
        defaultextension=".pdf",
        initialfile=f"{initial_base}_pattern.pdf",
        filetypes=[("PDF file", "*.pdf")]
    )
    if not path:
        return None
    base_root, _ = os.path.splitext(path)
    return base_root

# ============================================================
#  RENDER STITCH GRID AS IMAGE
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
                tw, th = draw.textsize(sym, font=font)
                draw.text(
                    (x*cell_px+(cell_px-tw)//2, y*cell_px+(cell_px-th)//2),
                    sym, fill=(0,0,0), font=font
                )
            else:
                draw.rectangle(
                    (x*cell_px, y*cell_px, x*cell_px+cell_px, y*cell_px+cell_px),
                    fill=arr[y][x],
                    outline=(60,60,60)
                )

    if numbered:
        # bold grid every 10
        for y in range(0, h, 10):
            draw.line((0, y*cell_px, w*cell_px, y*cell_px), fill=(0,0,0), width=2)
            draw.text((2, y*cell_px+2), str(y), fill=(255,255,255), font=font)

        for x in range(0, w, 10):
            draw.line((x*cell_px, 0, x*cell_px, h*cell_px), fill=(0,0,0), width=2)
            draw.text((x*cell_px+2, 2), str(x), fill=(255,255,255), font=font)

    return img

# ============================================================
#  PDF EXPORT (COLOR/SYMBOL) – WITH MARGINS + PAGES
# ============================================================

def export_tiled_pdf(big_img, filename):
    dpi = 300
    PAGE_W = int(8.5 * dpi)
    PAGE_H = int(11 * dpi)

    MARGIN = int(0.75 * dpi)   # 0.75" safe margin
    usable_w = PAGE_W - MARGIN*2
    usable_h = PAGE_H - MARGIN*2

    bw, bh = big_img.size

    pages = []
    page_index = 1

    y0 = 0
    while y0 < bh:
        x0 = 0
        while x0 < bw:
            x1 = min(x0 + usable_w, bw)
            y1 = min(y0 + usable_h, bh)
            crop = big_img.crop((x0, y0, x1, y1))

            # shrink a bit for nicer preview/printing
            scale = 0.9
            cw = int(crop.width * scale)
            ch = int(crop.height * scale)
            crop_scaled = crop.resize((cw, ch), Image.NEAREST)

            page = Image.new("RGB", (PAGE_W, PAGE_H), "white")

            # center within margins
            ox = MARGIN + (usable_w - cw)//2
            oy = MARGIN + (usable_h - ch)//2
            page.paste(crop_scaled, (ox, oy))

            draw = ImageDraw.Draw(page)
            border_box = (ox, oy, ox+cw, oy+ch)
            draw.rectangle(border_box, outline=(0,0,0), width=2)

            font = ImageFont.load_default()
            text = f"Page {page_index}"
            tw, th = draw.textsize(text, font=font)
            draw.text((PAGE_W//2 - tw//2, 20), text, fill=(0,0,0), font=font)

            pages.append(page)
            page_index += 1

            x0 += usable_w
        y0 += usable_h

    if pages:
        pages[0].save(filename, save_all=True, append_images=pages[1:])

def build_page_preview_image(big_img, title="Pattern Preview"):
    dpi = 150  # smaller for fast preview
    PAGE_W = int(8.5 * dpi)
    PAGE_H = int(11 * dpi)
    MARGIN = int(0.75 * dpi)
    usable_w = PAGE_W - MARGIN*2
    usable_h = PAGE_H - MARGIN*2

    bw, bh = big_img.size
    scale = min(usable_w / bw, usable_h / bh, 1.0)
    cw = int(bw * scale)
    ch = int(bh * scale)
    crop_scaled = big_img.resize((cw, ch), Image.NEAREST)

    page = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    ox = MARGIN + (usable_w - cw)//2
    oy = MARGIN + (usable_h - ch)//2
    page.paste(crop_scaled, (ox, oy))

    draw = ImageDraw.Draw(page)
    font = ImageFont.load_default()

    border_box = (ox, oy, ox+cw, oy+ch)
    draw.rectangle(border_box, outline=(0,0,0), width=2)

    draw.text((MARGIN, PAGE_H - MARGIN + 5), title, fill=(0,0,0), font=font)
    draw.text((PAGE_W//2 - 20, 20), "Page 1", fill=(0,0,0), font=font)

    return page

# ============================================================
#  GUI APP
# ============================================================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Needlepoint Pattern Designer — DMC Edition (macOS safe)")

        self.img_path = StringVar(value="")
        self.grid_max = IntVar(value=160)
        self.color_count = IntVar(value=40)
        self.cell_px = IntVar(value=18)
        self.keep_aspect = BooleanVar(value=True)
        self.include_dmc = BooleanVar(value=True)
        self.export_color = BooleanVar(value=True)
        self.export_symbols = BooleanVar(value=True)
        self.regular_only = BooleanVar(value=False)  # if True: restrict mapping to regular cotton

        frm = ttk.Frame(root, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        # File row
        row = ttk.Frame(frm)
        row.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(row, text="Image:").pack(side="left")
        ttk.Entry(row, textvariable=self.img_path, width=70).pack(side="left", padx=6)
        ttk.Button(row, text="Browse", command=self.browse).pack(side="left")

        # Settings panel
        sfrm = ttk.Frame(frm)
        sfrm.grid(row=1, column=0, sticky="nw")

        ttk.Label(sfrm, text="Max Stitches (long side):").grid(row=0, column=0, sticky="w")
        ttk.Entry(sfrm, textvariable=self.grid_max, width=8).grid(row=0, column=1)

        ttk.Label(sfrm, text="Color Count:").grid(row=1, column=0, sticky="w")
        ttk.Entry(sfrm, textvariable=self.color_count, width=8).grid(row=1, column=1)

        ttk.Label(sfrm, text="Pixels per Stitch:").grid(row=2, column=0, sticky="w")
        ttk.Entry(sfrm, textvariable=self.cell_px, width=8).grid(row=2, column=1)

        ttk.Checkbutton(sfrm, text="Keep aspect ratio", variable=self.keep_aspect).grid(row=3, column=0, sticky="w")
        ttk.Checkbutton(sfrm, text="Include DMC match", variable=self.include_dmc).grid(row=4, column=0, sticky="w")
        ttk.Checkbutton(sfrm, text="Export Color Chart", variable=self.export_color).grid(row=5, column=0, sticky="w")
        ttk.Checkbutton(sfrm, text="Export Symbol Chart", variable=self.export_symbols).grid(row=6, column=0, sticky="w")
        ttk.Checkbutton(sfrm, text="Limit to regular cotton only", variable=self.regular_only).grid(row=7, column=0, sticky="w")

        ttk.Button(sfrm, text="Generate Preview", command=self.generate_preview).grid(row=8, column=0, pady=10, sticky="w")
        ttk.Button(sfrm, text="Export PDFs + CSV", command=self.export_all).grid(row=9, column=0, pady=6, sticky="w")
        ttk.Button(sfrm, text="1-Page Preview", command=self.preview_one_page).grid(row=10, column=0, pady=6, sticky="w")

        # Preview area
        self.preview_label = ttk.Label(frm)
        self.preview_label.grid(row=1, column=1, sticky="nsew")

        self.grid_arr = None
        self._preview_full_img = None
        self._preview_page_img = None

    # ---------------------------
    def browse(self):
        try:
            p = safe_open_file_dialog(self.root)
            if p:
                self.img_path.set(p)
        except Exception as e:
            self.root.title(f"Dialog Error: {e}")

    # ---------------------------
    def _compute_grid(self):
        path = self.img_path.get()
        if not os.path.exists(path):
            self.root.title("Image not found.")
            return None, None, None

        img = Image.open(path).convert("RGB")
        maxs = self.grid_max.get()
        w, h = img.size

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
        arr = [[small.getpixel((x, y)) for x in range(out_w)] for y in range(out_h)]
        return img, arr, (out_w, out_h)

    # ---------------------------
    def generate_preview(self):
        img, arr, size = self._compute_grid()
        if arr is None:
            return
        self.grid_arr = arr
        out_w, out_h = size

        cell_px = self.cell_px.get()
        prev = render_pattern_image(arr, cell_px, numbered=True, symbols=False)

        max_w, max_h = 900, 500
        scale = min(max_w / prev.width, max_h / prev.height, 1.0)
        disp = prev.resize((int(prev.width * scale), int(prev.height * scale)), Image.NEAREST)

        tkimg = ImageTk.PhotoImage(disp)
        self.preview_label.configure(image=tkimg)
        self.preview_label.image = tkimg

        self.root.title(f"Preview {out_w} x {out_h} stitches")

    # ---------------------------
    def _build_palette_map(self):
        arr = self.grid_arr
        flat = [tuple(px) for row in arr for px in row]
        counts = Counter(flat)
        sorted_cols = [c for c, _ in counts.most_common()]
        palette_map = {}
        for i, col in enumerate(sorted_cols):
            sym = SYMBOLS[i % len(SYMBOLS)]
            palette_map[col] = sym
        return sorted_cols, palette_map, counts

    # ---------------------------
    def export_all(self):
        if self.grid_arr is None:
            self.root.title("Generate preview first.")
            return

        arr = self.grid_arr
        sorted_cols, palette_map, counts = self._build_palette_map()

        base = os.path.splitext(os.path.basename(self.img_path.get()))[0]
        base_root = safe_save_base_dialog(self.root, base)
        if not base_root:
            return

        csv_path = base_root + "_legend.csv"
        color_pdf_path = base_root + "_color.pdf"
        symbols_pdf_path = base_root + "_symbols.pdf"

        allow_specialty = not self.regular_only.get()

        # LEGEND CSV
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            f.write("symbol,r,g,b,hex,dmc_code,dmc_name,dmc_type,stitches\n")
            for col in sorted_cols:
                sym = palette_map[col]
                r, g, b = col
                hx = rgb_to_hex(col)

                if self.include_dmc.get() and DMC_PALETTE:
                    d_code, d_name, d_type = nearest_dmc(col, allow_specialty=allow_specialty)
                else:
                    d_code = d_name = d_type = ""

                st = counts[col]
                f.write(f"{sym},{r},{g},{b},{hx},{d_code},{d_name},{d_type},{st}\n")

        cell_px = self.cell_px.get()

        # COLOR PDF
        if self.export_color.get():
            big = render_pattern_image(arr, cell_px, numbered=True, symbols=False)
            export_tiled_pdf(big, color_pdf_path)

        # SYMBOL PDF
        if self.export_symbols.get():
            bigs = render_pattern_image(arr, cell_px, numbered=True, symbols=True, palette_map=palette_map)
            export_tiled_pdf(bigs, symbols_pdf_path)

        self.root.title(
            f"Exported: {os.path.basename(color_pdf_path)}, "
            f"{os.path.basename(symbols_pdf_path)}, {os.path.basename(csv_path)}"
        )

    # ---------------------------
    def preview_one_page(self):
        """Show:
           A) full-map color preview (zoomed out)
           B) page-style preview (Letter page with margins/border)
        """
        if self.grid_arr is None:
            self.root.title("Generate preview first.")
            return

        arr = self.grid_arr
        cell_px = self.cell_px.get()

        big = render_pattern_image(arr, cell_px, numbered=True, symbols=False)

        # Full-map preview (zoomed out)
        max_w, max_h = 700, 500
        scale = min(max_w / big.width, max_h / big.height, 1.0)
        full_preview = big.resize((int(big.width * scale), int(big.height * scale)), Image.NEAREST)

        # Page-style preview
        page_preview = build_page_preview_image(big, title="Pattern 1-page preview")

        win = ttk.Toplevel(self.root)
        win.title("1-Page Previews")

        ttk.Label(win, text="Full pattern map (zoomed out):").pack()
        full_canvas = ttk.Label(win)
        full_canvas.pack(pady=4)

        ttk.Label(win, text="Page-style preview (Letter layout):").pack()
        page_canvas = ttk.Label(win)
        page_canvas.pack(pady=4)

        self._preview_full_img = ImageTk.PhotoImage(full_preview)
        self._preview_page_img = ImageTk.PhotoImage(page_preview)

        full_canvas.configure(image=self._preview_full_img)
        page_canvas.configure(image=self._preview_page_img)

# ============================================================
#  MAIN
# ============================================================

if __name__ == "__main__":
    load_dmc_palette("dmc_palette_full.csv")  # safe if missing

    root = Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    App(root)
    root.minsize(1000, 600)
    root.mainloop()
