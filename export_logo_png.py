"""
Export ISWS logo SVG to PNG using a headless browser approach via subprocess.
No native C libraries needed — uses the built-in macOS screencapture or
falls back to creating an HTML wrapper and using /usr/bin/qlmanage.
"""
import subprocess
import sys
import os
from pathlib import Path

def export_svg_to_png_via_qlmanage(svg_path, out_path, size):
    """Use macOS Quick Look to render SVG to PNG."""
    tmp = out_path + ".tmp"
    subprocess.run(
        ["qlmanage", "-t", "-s", str(size), "-o", str(Path(tmp).parent), str(svg_path)],
        capture_output=True
    )
    # qlmanage appends .png to the filename
    ql_output = Path(str(svg_path) + ".png")
    parent_ql = Path(tmp).parent / (Path(svg_path).name + ".png")
    for candidate in [ql_output, parent_ql]:
        if candidate.exists():
            candidate.rename(out_path)
            return True
    return False


def export_with_pillow_svg(svg_path, out_dir, sizes):
    """Try using Pillow with svg support or fall back to qlmanage."""
    svg_file = Path(svg_path)
    stem = svg_file.stem

    for label, size in sizes.items():
        out_path = Path(out_dir) / f"{stem}-{size}px.png"
        success = export_svg_to_png_via_qlmanage(str(svg_file), str(out_path), size)
        if success:
            print(f"  {out_path}")
        else:
            print(f"  [skip] Could not export {size}px — trying sips...")
            # fallback: sips can convert some SVGs
            subprocess.run(
                ["sips", "-s", "format", "png", "-z", str(size), str(size),
                 str(svg_file), "--out", str(out_path)],
                capture_output=True
            )
            if out_path.exists():
                print(f"  {out_path}")
            else:
                print(f"  [fail] {size}px export failed")


SIZES = {
    "favicon": 32,
    "small": 64,
    "medium": 200,
    "large": 512,
    "xlarge": 1024,
}

if __name__ == "__main__":
    base = Path(__file__).parent
    logo_dir = base / "logo-exports"
    logo_dir.mkdir(exist_ok=True)

    for svg in ["isws-logo-mark.svg", "isws-logo-mark-light.svg"]:
        svg_path = base / svg
        if svg_path.exists():
            print(f"\nExporting {svg}:")
            export_with_pillow_svg(str(svg_path), str(logo_dir), SIZES)

    print(f"\nAll PNGs saved to: {logo_dir}")
