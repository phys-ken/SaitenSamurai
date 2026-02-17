
import sys
from pathlib import Path
import cv2
import numpy as np
from PIL import Image, ImageDraw

# Add main_src to path
sys.path.append(str(Path(__file__).parent.parent / "main_src"))

from mark_checker import crop_from_corrected_image

def draw_dashed_rect(draw, x, y, w, h, color='red', width=3, dash_len=5, gap_len=3):
    """Refactored from MarkCheckerGUI for standalone use"""
    for i in range(x, x + w, dash_len + gap_len):
        end = min(i + dash_len, x + w)
        draw.line([(i, y), (end, y)], fill=color, width=width)
    for i in range(x, x + w, dash_len + gap_len):
        end = min(i + dash_len, x + w)
        draw.line([(i, y + h), (end, y + h)], fill=color, width=width)
    for i in range(y, y + h, dash_len + gap_len):
        end = min(i + dash_len, y + h)
        draw.line([(x, i), (x, end)], fill=color, width=width)
    for i in range(y, y + h, dash_len + gap_len):
        end = min(i + dash_len, y + h)
        draw.line([(x + w, i), (x + w, end)], fill=color, width=width)

def main():
    # 1. Simulate Corrected Image (White background)
    h, w = 842, 595 # A4 approx at 72dpi? No, MARK2_BASE is 595x842
    base_img = np.full((h, w, 3), 255, dtype=np.uint8)
    
    # 2. Draw some "mark" boxes (Blue) simulating OMR result
    # Question 1, Choice 1 (Correct)
    cv2.rectangle(base_img, (100, 200), (140, 240), (255, 0, 0), 2) # Blue box
    # Question 1, Choice 2 (Incorrect)
    cv2.rectangle(base_img, (160, 200), (200, 240), (255, 0, 0), 2)
    
    # 3. Simulate Crop
    # Question area is around 100,200 to 200,240
    # Let's say we crop from (80, 180) with size (160, 100)
    bbox = (80, 180, 160, 100)
    
    # Use real function
    pil_img, crop_info = crop_from_corrected_image(base_img, bbox, scale_factor=2.0)
    
    # 4. Calculate Draw Coordinates for Correct Answer (Choice 1)
    # Choice 1 is at (100, 200, 40, 40) in base coords
    cx, cy, cw, ch = 100, 200, 40, 40
    
    crop_x = crop_info['crop_x']
    crop_y = crop_info['crop_y']
    scale_x = crop_info['scale_x']
    scale_y = crop_info['scale_y']
    # res_scale is 1.0 here
    
    draw_x = (cx - crop_x) * scale_x
    draw_y = (cy - crop_y) * scale_y
    draw_w = cw * scale_x
    draw_h = ch * scale_y
    
    # 5. Draw Dashed Red Box
    draw = ImageDraw.Draw(pil_img)
    margin = 3
    draw_dashed_rect(
        draw, 
        int(draw_x - margin), int(draw_y - margin), 
        int(draw_w + margin*2), int(draw_h + margin*2),
        color='red', width=2
    )
    
    # 6. Save
    output_path = Path("tests/visual_verification_result.png")
    pil_img.save(output_path)
    print(f"Saved visual verification to {output_path}")
    
    # Generate HTML report
    html_content = f"""
    <html>
    <body>
    <h1>Visual Verification of Improvements</h1>
    <h2>Improvement 2: Correct Answer Frame in Mark Check</h2>
    <p>The image below shows a simulated Mark Check view.</p>
    <p>Blue box: Detection area (simulated)</p>
    <p>Red dashed box: Correct answer frame (implemented)</p>
    <img src="{output_path.name}" alt="Verification Results" style="border:1px solid #ccc;">
    </body>
    </html>
    """
    with open("tests/visual_report.html", "w") as f:
        f.write(html_content)
    print("Saved HTML report to tests/visual_report.html")

if __name__ == "__main__":
    main()
