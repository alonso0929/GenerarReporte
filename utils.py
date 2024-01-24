from docxtpl import InlineImage
from docx.shared import Inches
import os
from datetime import datetime


def save_image(image, path):
    if image:
        image.save(path)

def generate_image_paths(num_images):
    return [f'static/imagen{i}.png' for i in range(1, num_images + 1)]

def generate_inline_images(document, image_paths):
    return [InlineImage(document, path, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(path) else None for path in image_paths]

def generate_date():
    date_time = datetime.now()
    return date_time.strftime("%d-%m-%Y")

def generate_time():
    date_time = datetime.now()
    return date_time.strftime("%H:%M:%S")