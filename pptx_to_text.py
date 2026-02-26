from pptx import Presentation
from pathlib import Path

# --------------------------------------------------
# User inputs
# --------------------------------------------------
pptx_path = Path("C:/DRAFUSE/ICMP/Iraq/Training/Trimble DGPS/Training/powerpoints/Day 1 Intro and CRS.pptx")
output_txt = Path("C:/DRAFUSE/ICMP/Iraq/Training/Trimble DGPS/Training/powerpoints/Day 1 Intro and CRS.txt")

# --------------------------------------------------
# Load presentation
# --------------------------------------------------
prs = Presentation(pptx_path)

# --------------------------------------------------
# Extract text
# --------------------------------------------------
lines = []

for slide_number, slide in enumerate(prs.slides, start=1):
    lines.append(f"--- Slide {slide_number} ---")

    for shape in slide.shapes:
        if hasattr(shape, "text"):
            text = shape.text.strip()
            if text:
                lines.append(text)

    lines.append("")  # blank line between slides

# --------------------------------------------------
# Write to text file
# --------------------------------------------------
with output_txt.open("w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"Text exported successfully to: {output_txt}")
