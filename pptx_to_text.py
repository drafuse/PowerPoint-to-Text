from pptx import Presentation
from pathlib import Path
import argparse

parser = argparse.ArgumentParser(description="Extract text from a .pptx file.")
parser.add_argument("pptx_path", type=Path, help="Path to the input .pptx file")
parser.add_argument(
    "-o",
    "--output",
    dest="output_txt",
    type=Path,
    help="Path to output .txt file (default: same name as input)",
)
args = parser.parse_args()

pptx_path = args.pptx_path
output_txt = args.output_txt or pptx_path.with_suffix(".txt")

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
