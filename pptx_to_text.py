import argparse
from pathlib import Path

from pptx import Presentation


def extract_pptx_text(pptx_path: Path, output_txt: Path) -> None:
    prs = Presentation(pptx_path)
    lines = []

    for slide_number, slide in enumerate(prs.slides, start=1):
        lines.append(f"--- Slide {slide_number} ---")

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text = shape.text.strip()
                if text:
                    lines.append(text)

        lines.append("")

    output_txt.write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract text from a .pptx file.")
    parser.add_argument("pptx_path", type=Path, help="Path to the input .pptx file")
    parser.add_argument(
        "-o",
        "--output",
        dest="output_txt",
        type=Path,
        help="Path to output .txt file (default: same name as input)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    pptx_path = args.pptx_path.expanduser().resolve()
    if not pptx_path.exists():
        print(f"Input file not found: {pptx_path}")
        return 1
    if pptx_path.suffix.lower() != ".pptx":
        print(f"Input must be a .pptx file: {pptx_path}")
        return 1

    output_txt = (args.output_txt or pptx_path.with_suffix(".txt")).expanduser().resolve()
    output_txt.parent.mkdir(parents=True, exist_ok=True)

    extract_pptx_text(pptx_path, output_txt)
    print(f"Text exported successfully to: {output_txt}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
