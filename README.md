
<p align="center">
<img width="490" height="150" alt="Image" src="https://github.com/user-attachments/assets/248d6860-cb9a-44d0-af48-0eecb1f47cbe" alt="Banner""/>
  

</p>

<h1 align="center">Digital Loom: Image-to-Needlepoint Pattern Generator</h1>

<br>

> Digital Loom is a standalone desktop tool that converts any image into a full needlepoint, cross-stitch, or tapestry pattern.
It creates a clean, printable chart with DMC color matching, symbol maps, stitch gridlines, and optional PDF + CSV exports.

<br> 

This version includes:

- Auto-fit single-page PDF export with no clipping
- Color chart and symbol chart generation
- Accurate DMC color matching (regular or specialty sets)
- CSV legend with codes, hex values, stitch counts, and assigned symbols
- Adjustable grid size, color count, and pixel-per-stitch controls
- Live preview and printable one-page preview
- MacOS-safe file dialogs and Pillow-10-safe text rendering
- Toplevel window fix for multi-preview mode

Digital Loom is built for creators who want fast, high-quality needlepoint or cross-stitch patterns without relying on paid software or online converters. Everything runs locally on your computer.

<br> 

## Features
### Image Processing

Adaptive palette reduction with user-selected color count

Optional aspect-ratio preservation

Customizable stitch resolution (max stitches)

### DMC Mapping

Full DMC palette support (including metallics, variegated, and specialty threads)

Option to restrict matching to standard cotton only

Automatic closest-color selection using RGB distance

### Pattern Rendering

Color chart with gridlines every 10 stitches

Symbol chart using a DMC-style symbol library

Numbered axes for easier alignment

High-contrast grid system and print-safe colors

### Export Options

Single-page, auto-fit PDF export for color and symbol charts

CSV legend including:

- Symbol
- RGB values
- Hex code
- DMC code
- DMC name
- Thread type (regular/metallic/etc.)
- Stitch count

### UI
- Simple Tkinter GUI
- Real-time preview
- One-page PDF preview window
- Safe file dialogs for macOS, Windows, and Linux

<br> 

## Installation
```
Install Python 3.12+

Install Pillow:
pip install pillow

Place the script and your palette file in the same directory:
needlepoint_designer_plus.py
dmc_palette_full.csv

Run:
python3 needlepoint_designer_plus.py

Usage

Click Browse to load an image.

Adjust:

Max stitches

Color count

Pixels per stitch

DMC options

Click Generate Preview.

When ready, click Export PDFs + CSV.
This creates:

image_pattern_color.pdf

image_pattern_symbols.pdf

image_legend.csv

Optional: use 1-Page Preview to preview the printable page layout.

```

## File Structure
```
digital_loom/
  demo_images/
  google-sheets-apps-script/
  needlepoint_designer_plus.py
  dmc_color_palette.xlsx
  dmc_color_palette_full.csv
  dmc_palette_full.csv
  README.md
```

<br> <br>

>original image
<img width="1536" height="1024" alt="Image" src="https://github.com/user-attachments/assets/d51cd8c7-5d78-49f9-9696-3378c0502e21" />

>uploaded to GUI

<img width="1049" height="616" alt="Image" src="https://github.com/user-attachments/assets/74583884-90a2-46fb-8370-e850d37c8e29" />

<img width="869" height="575" alt="Image" src="https://github.com/user-attachments/assets/a3b0f0f6-f7dc-4325-a4e4-411189d28f66" />

<img width="1088" height="742" alt="Image" src="https://github.com/user-attachments/assets/a3b2a362-a401-48c3-a58c-c141d45c701d" />

<img width="372" height="577" alt="Image" src="https://github.com/user-attachments/assets/e26e2ca3-dfb3-45af-90bb-ebe9617d7355" />

## DMC Google Sheets & Apps Script

<img width="1145" height="710" alt="Image" src="https://github.com/user-attachments/assets/2e08c43e-d346-4251-8716-6e27be7afaa8" />

<img width="1125" height="643" alt="Image" src="https://github.com/user-attachments/assets/3bd50ff4-23df-4f28-87e2-58a20d562554" />
