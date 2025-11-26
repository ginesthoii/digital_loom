import os
import csv
from collections import Counter

from tkinter import (
    Tk, StringVar, BooleanVar, IntVar, filedialog
)
from tkinter import ttk
import tkinter as tk

from PIL import Image, ImageTk, ImageDraw, ImageFont


# ============================================================
#  GLOBAL DMC PALETTE
# ============================================================

DMC_PALETTE = []  # list of dicts: number, name, type, r, g, b


def load_dmc_palette(csv_path="dmc_palette_full.csv"):
    """
    Load full DMC palette (all families) from CSV if present.

    CSV columns:
    number,name,r,g,b,type
    """
    global DMC_PALETTE
    DMC_PALETTE = []

    if not os.path.exists(csv_path):
        # No palette file is fine — app still runs, just no DMC mapping
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
                DMC_PALETTE.append(
                    {
                        "number": num,
                        "name": name,
                        "type": t,
                        "r": r,
                        "g": g,
                        "b": b,
                    }
                )
            except Exception:
                # Skip malformed rows silently
                continue


def nearest_dmc(rgb, allow_specialty=True):
    """
    Find nearest DMC color in loaded palette.

    If allow_specialty == False -> only 'regular' rows used.
    Returns (number, name, type) or ("", "", "") if no palette.
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
        d = dr * dr + dg * dg + db * db
        if d < bestd:
            bestd = d
            best = row

    if best is None:
        return ("", "", "")
    return (best["number"], best["name"], best["type"])


def rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)


# ============================================================
#  SYMBOL SET (DMC-STYLE)
# ============================================================

SYMBOLS = (
    ["●", "○", "■", "□", "▲", "△", "◆", "◇",
     "/", "\\", "x", "+", "-", "=", "*", "#",
     "◼", "◻", "◊", "◯", "¤", "▣", "▢"]
    + list("ABCDEFGHIJKLMNOQRSTUVWXYZ")
)


# ============================================================
#  IMAGE REDUCTION (image -> small palette grid)
# ============================================================

def quantize_image(img, out_w, out_h, colors):
    img_small = img.resize((out_w, out_h), Image.LANCZOS).convert("RGB")
    img_p = img_small.convert("P", palette=Image.ADAPTIVE, colors=colors)
    return img_p.convert("RGB")


# ============================================================
#  SAFE FILE DIALOGS (macOS-friendly)
# ============================================================

def safe_open_file_dialog(root):
    return filedialog.askopenfilename(
        parent=root,
        title="Select image",
        initialdir=os.path.expanduser("~"),
        filetypes=[
            ("PNG", "*.png"),
            ("JPG", "*.jpg"),
            ("JPEG", "*.jpeg"),
            ("WEBP", "*.webp"),
            ("BMP", "*.bmp"),
            ("All Files", "*"),
        ],
    )


def safe_save_base_dialog(root, initial_base):
    """
    Ask once where to save, then derive _color.pdf, _symbols.pdf, _legend.csv.
    """
    path = filedialog.asksaveasfilename(
        parent=root,
        title="Save Pattern",
        initialdir=os.path.expanduser("~"),
        defaultextension=".pdf",
        initialfile=f"{initial_base}_pattern.pdf",
        filetypes=[
            ("PDF", "*.pdf"),
            ("All Files", "*"),
        ],
    )
    if not path:
        return None

    base_root, _ = os.path.splitext(path)
    return base_root


# ============================================================
#  TEXT MEASUREMENT (Pillow 10+ safe)
# ============================================================

def _measure_text(draw, text, font):
    """
    Pillow 10+ dropped textsize(). Use textbbox with a fallback.
    """
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        return tw, th
    except Exception:
        try:
            return draw.textsize(text, font=font)
        except Exception:
            return (len(text) * 6, 10)


# ============================================================
#  RENDER STITCH GRID AS IMAGE (COLOR or SYMBOL)
# ============================================================

def render_pattern_image(arr, cell_px, numbered=True, symbols=False, palette_map=None):
    """
    arr: 2D list of RGB tuples
    symbols: if True, use palette_map[rgb] as symbol instead of color blocks
    """
    h = len(arr)
    w = len(arr[0])
    img = Image.new("RGB", (w * cell_px, h * cell_px), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    font = ImageFont.load_default()

    for y in range(h):
        for x in range(w):
            color = arr[y][x]
            x0 = x * cell_px
            y0 = y * cell_px
            x1 = x0 + cell_px
            y1 = y0 + cell_px

            if symbols:
                sym = palette_map[color]
                # White cell background + light grid
                draw.rectangle(
                    (x0, y0, x1, y1),
                    fill=(255, 255, 255),
                    outline=(200, 200, 200),
                )
                tw, th = _measure_text(draw, sym, font)
                draw.text(
                    (x0 + (cell_px - tw) // 2, y0 + (cell_px - th) // 2),
                    sym,
                    fill=(0, 0, 0),
                    font=font,
                )
            else:
                # solid color cells with dark grid
                draw.rectangle(
                    (x0, y0, x1, y1),
                    fill=color,
                    outline=(60, 60, 60),
                )

    if numbered:
        # thick grid line every 10 and coordinate labels
        for y in range(0, h, 10):
            ypix = y * cell_px
            draw.line((0, ypix, w * cell_px, ypix), fill=(0, 0, 0), width=2)
            label = str(y)
            tw, th = _measure_text(draw, label, font)
            draw.rectangle((2, ypix + 2, 2 + tw + 2, ypix + 2 + th), fill=(0, 0, 0))
            draw.text((4, ypix + 2), label, fill=(255, 255, 255), font=font)

        for x in range(0, w, 10):
            xpix = x * cell_px
            draw.line((xpix, 0, xpix, h * cell_px), fill=(0, 0, 0), width=2)
            label = str(x)
            tw, th = _measure_text(draw, label, font)
            draw.rectangle((xpix + 2, 2, xpix + 2 + tw + 2, 2 + th), fill=(0, 0, 0))
            draw.text((xpix + 4, 2), label, fill=(255, 255, 255), font=font)

    return img


# ============================================================
#  SINGLE-PAGE PDF EXPORT (AUTO-FIT, NO CLIPPING)
# ============================================================

def export_single_page_pdf(big_img, filename, margin_inches=0.5, dpi=300):
    """
    Puts the entire pattern on a single Letter page:
    - auto orientation (portrait vs landscape)
    - 0.5" margins
    - scaled so nothing touches the page edge
    """
    bw, bh = big_img.size
    if bw <= 0 or bh <= 0:
        return

    aspect = bw / bh if bh != 0 else 1.0

    # US Letter, orientation based on image
    if aspect >= 1.0:
        page_w_in, page_h_in = 11.0, 8.5   # landscape
    else:
        page_w_in, page_h_in = 8.5, 11.0   # portrait

    page_w = int(page_w_in * dpi)
    page_h = int(page_h_in * dpi)

    margin_px = int(margin_inches * dpi)
    usable_w = page_w - 2 * margin_px
    usable_h = page_h - 2 * margin_px

    if usable_w <= 0 or usable_h <= 0:
        usable_w = page_w
        usable_h = page_h

    # Scale to fit inside usable area, then shrink slightly for safety
    raw_scale = min(usable_w / bw, usable_h / bh)
    scale = min(raw_scale, 1.0) * 0.95  # 95% to guarantee no edge clipping

    cw = int(bw * scale)
    ch = int(bh * scale)

    page = Image.new("RGB", (page_w, page_h), "white")
    ox = (page_w - cw) // 2
    oy = (page_h - ch) // 2

    resized = big_img.resize((cw, ch), Image.NEAREST)
    page.paste(resized, (ox, oy))

    draw = ImageDraw.Draw(page)
    font = ImageFont.load_default()

    # Border box around stitched area
    border_box = (ox, oy, ox + cw, oy + ch)
    draw.rectangle(border_box, outline=(0, 0, 0), width=2)

    # Page label
    text = "Page 1"
    tw, th = _measure_text(draw, text, font)
    draw.text((page_w // 2 - tw // 2, 20), text, fill=(0, 0, 0), font=font)

    page.save(filename)


def build_page_preview_image(big_img, title="Pattern Preview", margin_inches=0.5, dpi=150):
    """
    Build a single-page, letter-sized preview image with margins + border.
    Uses same auto-fit logic as the PDF, just at lower DPI for speed.
    """
    bw, bh = big_img.size
    if bw <= 0 or bh <= 0:
        return Image.new("RGB", (600, 400), "white")

    aspect = bw / bh if bh != 0 else 1.0

    if aspect >= 1.0:
        page_w_in, page_h_in = 11.0, 8.5   # landscape
    else:
        page_w_in, page_h_in = 8.5, 11.0   # portrait

    page_w = int(page_w_in * dpi)
    page_h = int(page_h_in * dpi)

    margin_px = int(margin_inches * dpi)
    usable_w = page_w - 2 * margin_px
    usable_h = page_h - 2 * margin_px

    raw_scale = min(usable_w / bw, usable_h / bh)
    scale = min(raw_scale, 1.0) * 0.95

    cw = int(bw * scale)
    ch = int(bh * scale)

    page = Image.new("RGB", (page_w, page_h), "white")
    ox = (page_w - cw) // 2
    oy = (page_h - ch) // 2

    resized = big_img.resize((cw, ch), Image.NEAREST)
    page.paste(resized, (ox, oy))

    draw = ImageDraw.Draw(page)
    font = ImageFont.load_default()

    # Border
    border_box = (ox, oy, ox + cw, oy + ch)
    draw.rectangle(border_box, outline=(0, 0, 0), width=2)

    # Title + page label
    if title:
        tw, th = _measure_text(draw, title, font)
        draw.text((margin_px, page_h - margin_px + 5), title, fill=(0, 0, 0), font=font)

    label = "Page 1"
    tw, th = _measure_text(draw, label, font)
    draw.text((page_w // 2 - tw // 2, 20), label, fill=(0, 0, 0), font=font)

    return page


# ============================================================
#  GUI APP
# ============================================================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Needlepoint Pattern Designer — DMC Pro Edition (macOS safe)")

        self.img_path = StringVar(value="")
        self.grid_max = IntVar(value=160)
        self.color_count = IntVar(value=40)
        self.cell_px = IntVar(value=18)
        self.keep_aspect = BooleanVar(value=True)
        self.include_dmc = BooleanVar(value=True)
        self.export_color = BooleanVar(value=True)
        self.export_symbols = BooleanVar(value=True)
        self.regular_only = BooleanVar(value=False)  # if True: only regular cotton DMC

        frm = ttk.Frame(root, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        # ---- File row ----
        row = ttk.Frame(frm)
        row.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(row, text="Image:").pack(side="left")
        ttk.Entry(row, textvariable=self.img_path, width=70).pack(side="left", padx=6)
        ttk.Button(row, text="Browse", command=self.browse).pack(side="left")

        # ---- Settings panel ----
        sfrm = ttk.Frame(frm)
        sfrm.grid(row=1, column=0, sticky="nw")

        ttk.Label(sfrm, text="Max Stitches (long side):").grid(row=0, column=0, sticky="w")
        ttk.Entry(sfrm, textvariable=self.grid_max, width=8).grid(row=0, column=1)

        ttk.Label(sfrm, text="Color Count:").grid(row=1, column=0, sticky="w")
        ttk.Entry(sfrm, textvariable=self.color_count, width=8).grid(row=1, column=1)

        ttk.Label(sfrm, text="Pixels per Stitch:").grid(row=2, column=0, sticky="w")
        ttk.Entry(sfrm, textvariable=self.cell_px, width=8).grid(row=2, column=1)

        ttk.Checkbutton(sfrm, text="Keep aspect ratio", variable=self.keep_aspect).grid(row=3, column=0, sticky="w")
        ttk.Checkbutton(sfrm, text="Include DMC mapping", variable=self.include_dmc).grid(row=4, column=0, sticky="w")
        ttk.Checkbutton(sfrm, text="Export Color Chart PDF", variable=self.export_color).grid(row=5, column=0, sticky="w")
        ttk.Checkbutton(sfrm, text="Export Symbol Chart PDF", variable=self.export_symbols).grid(row=6, column=0, sticky="w")
        ttk.Checkbutton(
            sfrm,
            text="Limit DMC mapping to regular cotton only",
            variable=self.regular_only,
        ).grid(row=7, column=0, sticky="w", columnspan=2)

        ttk.Button(sfrm, text="Generate Preview", command=self.generate_preview).grid(
            row=8, column=0, pady=10, sticky="w"
        )
        ttk.Button(sfrm, text="Export PDFs + CSV", command=self.export_all).grid(
            row=9, column=0, pady=6, sticky="w"
        )
        ttk.Button(sfrm, text="1-Page Preview", command=self.preview_one_page).grid(
            row=10, column=0, pady=6, sticky="w"
        )

        # ---- Preview area ----
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

        # ---- LEGEND CSV ----
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
                f.write(
                    f"{sym},{r},{g},{b},{hx},{d_code},{d_name},{d_type},{st}\n"
                )

        cell_px = self.cell_px.get()

        # ---- COLOR PDF (single page, auto-fit) ----
        if self.export_color.get():
            big = render_pattern_image(arr, cell_px, numbered=True, symbols=False)
            export_single_page_pdf(big, color_pdf_path, margin_inches=0.5, dpi=300)

        # ---- SYMBOL PDF (single page, auto-fit) ----
        if self.export_symbols.get():
            bigs = render_pattern_image(arr, cell_px, numbered=True, symbols=True, palette_map=palette_map)
            export_single_page_pdf(bigs, symbols_pdf_path, margin_inches=0.5, dpi=300)

        self.root.title(
            f"Exported: {os.path.basename(color_pdf_path)}, "
            f"{os.path.basename(symbols_pdf_path)}, {os.path.basename(csv_path)}"
        )

    # ---------------------------
    def preview_one_page(self):
        """
        Show:
          A) full-map color preview (zoomed out)
          B) page-style preview (Letter layout, auto-fit)
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
        page_preview = build_page_preview_image(big, title="Pattern 1-page preview", margin_inches=0.5, dpi=150)

        win = tk.Toplevel(self.root)
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
    # Load DMC palette if CSV is present; app still works without it.
    load_dmc_palette("dmc_palette_full.csv")

    root = Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    App(root)
    root.minsize(1000, 600)
    root.mainloop()
