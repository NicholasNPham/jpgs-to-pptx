"""
jpgs_to_pptx.py
----------------
Scans the folder this script (or its packaged .exe) is running from,
finds every .jpg/.jpeg image in that folder, and assembles them into
a single PowerPoint file, one image per slide. Each image is scaled
down only if it's larger than the slide (never scaled up), centered,
and placed on a black background, to mimic a plain copy/paste.

Requirements:
    pip install python-pptx pillow

Usage (running as a script):
    1. Place this file in the same folder as your .jpg/.jpeg images.
    2. Run: python jpgs_to_pptx.py
    3. Output: output.pptx saved into that same folder.

Usage (running as a packaged .exe):
    1. Drop the .exe into any folder containing .jpg/.jpeg images.
    2. Run the .exe.
    3. Output: output.pptx saved next to the .exe, in that same folder.
"""

# STANDARD LIBRARY IMPORTS
import sys

# THIRD-PARTY IMPORTS
from pathlib import Path
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor

# CONSTANTS
DPI = 96
EMU_PER_INCH = 914400
BLANK_SLIDE_INDEX = 6
BLACK_RGB = (0, 0, 0)
FILE_EXTENSIONS = [".jpg", ".jpeg"]
OUTPUT_FILENAME = "output.pptx"

# MAIN FUNCTION SETUP
if getattr(sys, "frozen", False):
    RUN_FOLDER = Path(sys.executable).parent
else:
    RUN_FOLDER = Path(__file__).parent

JPEG_FOLDER = RUN_FOLDER
OUTPUT_PPTX = RUN_FOLDER / OUTPUT_FILENAME


def get_sorted_jpegs(folder):
    """
    Returns a sorted list of every .jpg/.jpeg file in the given folder.

    Args:
        folder: Path (or path-like) to the folder to search.

    Returns:
        A sorted list of Path objects, one per matching image file.
    """
    local_dir = Path(folder)
    jpeg_jpg_list = []

    for jpeg_file in local_dir.iterdir():
        if jpeg_file.is_file() and jpeg_file.suffix.lower() in FILE_EXTENSIONS:
            jpeg_jpg_list.append(jpeg_file)

    return sorted(jpeg_jpg_list)


def add_image(prs, image_file, left, top, width, height):
    """
    Adds one blank slide to the presentation, fills its background
    black, and places the given image on it at the given position
    and size.

    Args:
        prs: The Presentation object to add the slide to.
        image_file: Path (as a string) to the image file to place.
        left: EMU distance from the slide's left edge to the image.
        top: EMU distance from the slide's top edge to the image.
        width: EMU width to draw the image at.
        height: EMU height to draw the image at.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[BLANK_SLIDE_INDEX])

    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(BLACK_RGB[0], BLACK_RGB[1], BLACK_RGB[2])

    slide.shapes.add_picture(image_file, left, top, width, height)


def scale_image_for_slide(image, prs):
    """
    Computes the width, height, and centered position (all in EMU)
    to draw the given image on the given presentation's slides,
    scaled down to fit if it's larger than the slide, and never
    scaled up if it's smaller.

    Args:
        image: An opened PIL Image object.
        prs: The Presentation object whose slide dimensions to fit to.

    Returns:
        A tuple of (width, height, left, top), all in EMU.
    """
    width, height = image.size

    width_file = width / DPI
    height_file = height / DPI

    emu_width = width_file * EMU_PER_INCH
    emu_height = height_file * EMU_PER_INCH

    width_scale = prs.slide_width / emu_width
    height_scale = prs.slide_height / emu_height

    scale = min(width_scale, height_scale, 1.0)
    file_width = emu_width * scale
    file_height = emu_height * scale

    left = (prs.slide_width - file_width) / 2
    top = (prs.slide_height - file_height) / 2

    return file_width, file_height, left, top


def main():
    """
    Entry point. Finds every JPEG in JPEG_FOLDER, scales and centers
    each one onto its own black-background slide, and saves the
    finished presentation to OUTPUT_PPTX. Prints progress as it goes
    and exits early with a message if no JPEGs are found.
    """
    prs = Presentation()

    jpg_list = get_sorted_jpegs(JPEG_FOLDER)

    if not jpg_list:
        print(f"No JPEG files found in {JPEG_FOLDER}")
        return

    print(f"Found {len(jpg_list)} image(s). Building Powerpoint")

    for i, jpeg_file_object in enumerate(jpg_list, start=1):
        print(f"  Adding slide {i}/{len(jpg_list)}: {jpeg_file_object.name}")
        jpeg_file_string = str(jpeg_file_object)
        img = Image.open(jpeg_file_string)
        w, h, l, t = scale_image_for_slide(img, prs)
        add_image(prs, jpeg_file_string, l, t, w, h)

    prs.save(OUTPUT_PPTX)
    print(f"\nDone. Saved to: {OUTPUT_PPTX}")


if __name__ == "__main__":
    main()