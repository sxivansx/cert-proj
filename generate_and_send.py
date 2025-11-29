import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.message import EmailMessage

OUTPUT_DIR = "output"
TEMPLATE_PATH = "template/vib-2025.png"

NAME_Y = 500
TEAM_Y = 580
OTHER_Y = 660

NAME_FONT_PATH = "fonts/SouvenirB.ttf"
TEAM_FONT_PATH = "fonts/Souvenir.ttf"
OTHER_FONT_PATH = "fonts/Souvenir.ttf"

NAME_FONT_SIZE = 40
TEAM_FONT_SIZE = 40
OTHER_FONT_SIZE = 40


def make_output_directory():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)


def generate_certificate(row):
    name = str(row["Name"]).strip()
    team = str(row.get("Team", "")).strip()
    other = str(row.get("Other", "")).strip()

    img = Image.open(TEMPLATE_PATH).convert("RGB")
    draw = ImageDraw.Draw(img)

    name_font = ImageFont.truetype(NAME_FONT_PATH, NAME_FONT_SIZE)
    team_font = ImageFont.truetype(TEAM_FONT_PATH, TEAM_FONT_SIZE)
    other_font = ImageFont.truetype(OTHER_FONT_PATH, OTHER_FONT_SIZE)

    img_width, img_height = img.size

    def draw_centered_text(text, y, font):
        if not text:
            return
        text_width, text_height = draw.textsize(text, font=font)
        x = (img_width - text_width) // 2
        draw.text((x, y), text, font=font, fill="black")

    draw_centered_text(name, NAME_Y, name_font)
    draw_centered_text(team, TEAM_Y, team_font)
    draw_centered_text(other, OTHER_Y, other_font)

    safe_name = name.replace(" ", "_")
    output_path = os.path.join(OUTPUT_DIR, f"certificate_{safe_name}.png")

    img.save(output_path, "PNG")

    return output_path
