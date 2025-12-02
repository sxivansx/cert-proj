import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

DATA_PATH = "data/recipients.xlsx"
TEMPLATE_PATH = "template/vib-2025.png"
OUTPUT_DIR = "output"

FONT_PATH = "fonts/Souvenir.ttf"
FONT_SIZE = 30

NAME_POS = (500, 750)
OTHER_POS = (950, 750)
TEAM_POS = (1400, 750)

TEXT_COLOR = "#684631"

COL_NAME = "Name"
COL_USN = "USN"
COL_TEAM = "TEAM NAME"

def make_output_directory():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

def compute_x(image_width, x_given):
    """
    Convert X coordinate. If x_given < 0 treat it as offset from right edge:
      final_x = image_width + x_given
    Otherwise use x_given as absolute from left edge.
    """
    return int(image_width + x_given) if x_given < 0 else int(x_given)

def load_font(path, size):
    """Try to load a TTF font, fall back to default PIL font on error."""
    try:
        font = ImageFont.truetype(path, size)
        print(f"[FONT] Loaded '{path}' at size {size}")
        return font
    except Exception as e:
        print(f"[FONT] WARNING: failed to load '{path}': {e}")
        print("[FONT] Falling back to default font.")
        return ImageFont.load_default()

def draw_centered(draw, text, center_x, y, font, fill):
    """
    Draw text centered horizontally at (center_x, y).
    Uses textbbox for accurate measurement.
    """
    if not text:
        return
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    x = int(center_x - (text_width / 2))
    draw.text((x, y), text, font=font, fill=fill)

def _normalize(col_name):
    """Normalize column name: keep only alphanumeric lower-case for matching."""
    if not isinstance(col_name, str):
        return ""
    return "".join(ch.lower() for ch in col_name if ch.isalnum())

def _find_column(col_name, columns):
    """
    Find actual column name from columns list/index that matches col_name
    ignoring case, spaces and non-alphanumeric characters.
    Returns the exact column name if found, else None.
    """
    target = _normalize(col_name)
    if not target:
        return None
    # exact normalized match
    for c in columns:
        if _normalize(c) == target:
            return c
    # fallback: target contained in candidate or vice-versa
    for c in columns:
        cand = _normalize(c)
        if target in cand or cand in target:
            return c
    return None

def get_cell_value(row, col_name):
    """
    Safely extract cell value from a pandas row.
    Uses fuzzy column matching so Excel header variations still work.
    Returns empty string for NaN/None so we don't print 'nan'.
    """
    real_col = _find_column(col_name, row.index)
    if real_col is None:
        # column not found; return empty string
        return ""
    val = row[real_col]
    if pd.isna(val):
        return ""
    return str(val).strip()

def generate_certificate(row):
    """
    Generate a single certificate image for a row (pandas Series).
    Uses COL_NAME, COL_USN, COL_TEAM to fetch values.
    """
    person_name = get_cell_value(row, COL_NAME)
    usn_value = get_cell_value(row, COL_USN)
    team_value = get_cell_value(row, COL_TEAM)

    img = Image.open(TEMPLATE_PATH).convert("RGB")
    draw = ImageDraw.Draw(img)

    width, height = img.size
    name_x = compute_x(width, NAME_POS[0])
    name_y = NAME_POS[1]
    other_x = compute_x(width, OTHER_POS[0])
    other_y = OTHER_POS[1]
    team_x = compute_x(width, TEAM_POS[0])
    team_y = TEAM_POS[1]

    font = load_font(FONT_PATH, FONT_SIZE)

    print(f"[ROW] name='{person_name}' | USN='{usn_value}' | team='{team_value}'")
    print(f"[POS] name_x={name_x}, name_y={name_y} | usn_x={other_x}, usn_y={other_y} | team_x={team_x}, team_y={team_y}")

    draw_centered(draw, person_name, name_x, name_y, font, fill=TEXT_COLOR)
    draw_centered(draw, usn_value, other_x, other_y, font, fill=TEXT_COLOR)
    draw_centered(draw, team_value, team_x, team_y, font, fill=TEXT_COLOR)

    safe_name = person_name if person_name else "no_name"
    safe_name = "_".join(safe_name.split())

    output_path = os.path.join(OUTPUT_DIR, f"certificate_{safe_name}.png")
    img.save(output_path, "PNG")
    print(f"[OK] Saved -> {output_path}\n")
    return output_path

def main():
    make_output_directory()

    if not os.path.exists(TEMPLATE_PATH):
        print(f"[ERROR] Template not found: {TEMPLATE_PATH}")
        return
    if not os.path.exists(DATA_PATH):
        print(f"[ERROR] Data file not found: {DATA_PATH}")
        return

    df = pd.read_excel(DATA_PATH)
    print(f"[DATA] Rows to process: {len(df)}")
    print(f"[DATA] Columns found: {list(df.columns)}")

    # optional: show resolved column mapping for clarity
    cols = list(df.columns)
    mapped_name = _find_column(COL_NAME, cols) or "<not found>"
    mapped_usn = _find_column(COL_USN, cols) or "<not found>"
    mapped_team = _find_column(COL_TEAM, cols) or "<not found>"
    print(f"[MAP] Name -> {mapped_name} | USN -> {mapped_usn} | TEAM -> {mapped_team}")

    for idx, row in df.iterrows():
        try:
            generate_certificate(row)
        except Exception as e:
            print(f"[ERROR] Row {idx} failed: {e}")

    print("\n🎉 All done. Check the output/ folder.")

if __name__ == "__main__":
    main()
