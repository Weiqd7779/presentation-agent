import sys
import os
from pathlib import Path

# Provide access to pptx_utils
SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

try:
    from pptx_utils import configure_utf8_output
except ImportError:
    def configure_utf8_output(): pass
import glob
import math
from PIL import Image, ImageDraw, ImageFont

def make_grids(output_dir, prefix):
    # Find all images matching the prefix
    pattern = os.path.join(output_dir, f"{prefix}-*.jpg")
    images = sorted(glob.glob(pattern))
    
    if not images:
        print(f"No images found for pattern: {pattern}")
        return

    # Configuration for grids
    cols = 3
    total_slides = len(images)
    rows = math.ceil(total_slides / cols)

    # Open first image to get dimensions
    with Image.open(images[0]) as img:
        img_w, img_h = img.width, img.height

    # Setup margins and spacing to match premium look
    # Using larger margins for text and spacing between cells
    padding_x = 60
    padding_top = 100
    padding_bottom = 60
    
    cell_w = img_w + (padding_x * 2)
    cell_h = img_h + padding_top + padding_bottom

    # Calculate overall grid size
    grid_w = cell_w * cols
    grid_h = cell_h * rows
    
    # Try to load a clean font
    try:
        # Standard fonts on Windows
        font = ImageFont.truetype("arial.ttf", 64)
    except IOError:
        try:
            font = ImageFont.truetype("segoeui.ttf", 64)
        except IOError:
            font = ImageFont.load_default()

    # Create one single large canvas for ALL slides
    grid_img = Image.new('RGB', (grid_w, grid_h), color=(245, 245, 245)) # Slightly off-white for premium feel
    # Actually, the user's image is pure white background
    grid_img = Image.new('RGB', (grid_w, grid_h), color=(255, 255, 255))
    draw = ImageDraw.Draw(grid_img)
    
    for index, img_path in enumerate(images):
        col = index % cols
        row = index // cols
        slide_num = index + 1
        
        # Calculate offsets
        x_start = col * cell_w
        y_start = row * cell_h
        
        # Center the image within the cell's horizontal space
        img_x = x_start + padding_x
        img_y = y_start + padding_top
        
        with Image.open(img_path) as img:
            # Draw a subtle border around the slide image
            border_color = (180, 180, 180)
            draw.rectangle([img_x - 2, img_y - 2, img_x + img_w + 1, img_y + img_h + 1], outline=border_color, width=2)
            grid_img.paste(img, (img_x, img_y))
            
        # Draw Label Text: "slide[X].xml"
        text = f"slide{slide_num}.xml"
        
        # Calculate text position for centering
        # Use textbbox if available (Pillow 9.2.0+)
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            text_w = bbox[2] - bbox[0]
        except AttributeError:
            # Fallback for older Pillow
            text_w, _ = draw.textsize(text, font=font)
            
        text_x = x_start + (cell_w - text_w) // 2
        text_y = y_start + (padding_top // 2) - 20 # Align slightly above the image
        
        draw.text((text_x, text_y), text, fill=(30, 30, 30), font=font)
            
    grid_filename = os.path.join(output_dir, f"{prefix}_overview.jpg")
    grid_img.save(grid_filename, quality=92, optimize=True)
    print(f"Generated comprehensive overview: {grid_filename}")

if __name__ == "__main__":
    configure_utf8_output()
    if len(sys.argv) < 3:
        print("Usage: make_grids.py <output_dir> <prefix>")
        sys.exit(1)
        
    out_dir = sys.argv[1]
    name_prefix = sys.argv[2]
    make_grids(out_dir, name_prefix)
