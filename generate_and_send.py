import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ---------- CONFIG ----------

TEMPLATE_PATH = "template/certificate.png"   # <--- change name if different
DATA_PATH = "data/recipients.xlsx"          # <--- change if your file name is diff

FONT_PATH_BOLD = "fonts/YourBoldFont.ttf"   # <--- put your actual font file name
FONT_PATH_REG = "fonts/YourRegularFont.ttf" # <--- or use same as bold if you want

OUTPUT_DIR = "output"

# Positions - you will tweak these numbers
NAME_X, NAME_Y = 1300, 900
TEAM_X, TEAM_Y = 1300, 1000
OTHER_X, OTHER_Y = 1300, 1100

NAME_FONT_SIZE = 80
TEAM_FONT_SIZE = 50
OTHER_FONT_SIZE = 50


# ---------- HELPERS ----------

def make_output_dir():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

def generate_certificate(row):
    """
    row: pandas Series with fields: Name, Team, Other
    returns: path to generated certificate file
    """
    name = str(row["Name"]).strip()
    team = str(row.get("Team", "")).strip()
    other = str(row.get("Other", "")).strip()

    img = Image.open(TEMPLATE_PATH).convert("RGB")
    draw = ImageDraw.Draw(img)

    name_font = ImageFont.truetype(FONT_PATH_BOLD, NAME_FONT_SIZE)
    team_font = ImageFont.truetype(FONT_PATH_REG, TEAM_FONT_SIZE)
    other_font = ImageFont.truetype(FONT_PATH_REG, OTHER_FONT_SIZE)

    # Helper to center text horizontally around a given X
    def draw_centered(text, center_x, y, font):
        if not text:
            return
        text_w, text_h = draw.textsize(text, font=font)
        x = center_x - text_w // 2
        draw.text((x, y), text, font=font, fill=(0, 0, 0))

    # Draw all fields
    draw_centered(name, NAME_X, NAME_Y, name_font)
    if team:
        draw_centered(team, TEAM_X, TEAM_Y, team_font)
    if other:
        draw_centered(other, OTHER_X, OTHER_Y, other_font)

    safe_name = name.replace(" ", "_")
    output_path = os.path.join(OUTPUT_DIR, f"certificate_{safe_name}.png")
    img.save(output_path, "PNG")

    return output_path


# ---------- MAIN ----------

def main():
    make_output_dir()

    # Read Excel
    df = pd.read_excel(DATA_PATH)  # Make sure columns: Name, Team, Other at least

    for idx, row in df.iterrows():
        name = str(row["Name"]).strip()
        print(f"Generating for: {name}")
        cert_path = generate_certificate(row)
        print(f" -> Saved: {cert_path}")

    print("Done. Check the output/ folder.")

if __name__ == "__main__":
    main()
