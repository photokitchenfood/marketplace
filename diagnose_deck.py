#!/usr/bin/env python3
"""
Read-only diagnostic dump of a Recipe & Captions Deck .pptx.

For each sampled slide, prints every shape's type, position/size
(inches, from slide top-left), and text content, plus fill color
for non-text/autoshapes (useful for spotting the campaign color banner).

Does not modify the deck. Run this first on a handful of slides to
confirm exactly where Layout Name, category label, caption, and
recipe text live before the real extractor is written.

Usage:
    python3 diagnose_deck.py "<path to .pptx>" [--slides 1,2,3,4,5,6] [--all]
"""

import argparse
import sys
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE


def emu_to_in(emu):
    if emu is None:
        return None
    return round(Emu(emu).inches, 2)


def describe_fill(shape):
    """Best-effort fill color description. Returns None if not applicable/solid."""
    try:
        fill = shape.fill
    except Exception:
        return None
    try:
        fill_type = fill.type
    except Exception:
        return None
    if fill_type is None:
        return None
    try:
        if str(fill_type) == "MSO_FILL_TYPE.SOLID (1)" or fill_type == 1:
            color = fill.fore_color
            try:
                if color.type is not None and str(color.type).startswith("MSO_THEME_COLOR"):
                    return f"solid theme_color={color.theme_color} brightness={color.brightness}"
            except Exception:
                pass
            try:
                return f"solid rgb=#{color.rgb}"
            except Exception:
                return "solid (rgb unavailable)"
        return f"fill_type={fill_type}"
    except Exception:
        return None


def describe_shape(shape, indent=""):
    lines = []
    left = emu_to_in(getattr(shape, "left", None))
    top = emu_to_in(getattr(shape, "top", None))
    width = emu_to_in(getattr(shape, "width", None))
    height = emu_to_in(getattr(shape, "height", None))

    shape_type = shape.shape_type
    name = shape.name

    header = f"{indent}- shape_id={shape.shape_id} name={name!r} type={shape_type}"
    header += f" pos=(L={left}, T={top}, W={width}, H={height}) in"
    lines.append(header)

    if getattr(shape, "is_placeholder", False):
        try:
            pf = shape.placeholder_format
            header2 = f"{indent}    placeholder: idx={pf.idx} type={pf.type}"
            lines.append(header2)
        except Exception as e:
            lines.append(f"{indent}    placeholder: <error: {e}>")

    # detect an image inside what python-pptx reports as a plain placeholder
    # (picture placeholders filled via insert_picture keep a text_frame but
    # carry image data accessible through .image)
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        try:
            img = shape.image
            lines.append(
                f"{indent}    [placeholder holds image] ext={img.ext} size={img.size} content_type={img.content_type}"
            )
        except Exception:
            pass

    if shape.has_text_frame:
        text = shape.text_frame.text
        if text.strip():
            snippet = text if len(text) <= 300 else text[:300] + "...[truncated]"
            snippet = snippet.replace("\n", "\\n")
            lines.append(f"{indent}    text: {snippet!r}")
        # per-paragraph font info can matter for distinguishing title vs body
        for pi, para in enumerate(shape.text_frame.paragraphs):
            runs_info = []
            for run in para.runs:
                sz = run.font.size
                sz_pt = sz.pt if sz else None
                runs_info.append(
                    f"(bold={run.font.bold}, size={sz_pt}pt, color={_run_color(run)})"
                )
            if runs_info and any(r.strip() for r in [para.text]):
                lines.append(f"{indent}    para[{pi}] text={para.text!r} runs={runs_info}")

    fill_desc = describe_fill(shape)
    if fill_desc:
        lines.append(f"{indent}    fill: {fill_desc}")

    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        try:
            image = shape.image
            lines.append(
                f"{indent}    image: ext={image.ext} size={image.size} content_type={image.content_type}"
            )
        except Exception as e:
            lines.append(f"{indent}    image: <error reading image: {e}>")

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        lines.append(f"{indent}    group contains {len(shape.shapes)} shapes:")
        for sub in shape.shapes:
            lines.extend(describe_shape(sub, indent + "    "))

    if shape.has_table:
        lines.append(f"{indent}    [table] rows={len(shape.table.rows)} cols={len(shape.table.columns)}")

    return lines


def _run_color(run):
    try:
        color = run.font.color
        if color.type is None:
            return None
        try:
            return f"#{color.rgb}"
        except Exception:
            return f"theme:{color.theme_color}"
    except Exception:
        return None


def dump_slide(prs, idx):
    slide = prs.slides[idx]
    print(f"\n{'=' * 80}")
    print(f"SLIDE {idx + 1}  (layout: {slide.slide_layout.name!r})")
    print(f"{'=' * 80}")
    if not slide.shapes:
        print("  <no shapes>")
        return
    for shape in slide.shapes:
        for line in describe_shape(shape):
            print(line)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx_path", help="Path to the .pptx deck")
    parser.add_argument(
        "--slides",
        help="Comma-separated 1-based slide numbers to dump, e.g. 1,2,3,4,5,6",
        default=None,
    )
    parser.add_argument(
        "--all", action="store_true", help="Dump every slide (can be long)"
    )
    args = parser.parse_args()

    prs = Presentation(args.pptx_path)
    total = len(prs.slides)
    print(f"Deck: {args.pptx_path}")
    print(f"Total slides: {total}")
    print(f"Slide size: {emu_to_in(prs.slide_width)} x {emu_to_in(prs.slide_height)} in")

    if args.all:
        indices = range(total)
    elif args.slides:
        indices = [int(s.strip()) - 1 for s in args.slides.split(",")]
    else:
        # default sample: first 6 slides, should cover at least one
        # RECIPE pair (2 slides) plus whatever single-slide item follows
        indices = range(min(6, total))

    for idx in indices:
        if idx < 0 or idx >= total:
            print(f"\n[skip] slide {idx + 1} out of range (deck has {total} slides)")
            continue
        dump_slide(prs, idx)


if __name__ == "__main__":
    sys.exit(main())
