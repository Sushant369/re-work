import argparse
import logging

from .headings import HeadingConfig
from .pipeline import run


def main() -> None:
    ap = argparse.ArgumentParser(prog="python -m pptx_pipeline", description="Structure-aware PPTX extraction")
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", default="out")
    ap.add_argument("--notes", action="store_true", help="include speaker notes")
    ap.add_argument("--max-level", type=int, default=3)
    ap.add_argument("--min-support", type=int, default=None, help="slides a style must recur on to be a heading")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING)

    cfg = HeadingConfig(max_level=args.max_level, min_support=args.min_support)
    deck = run(args.pptx, args.out, config=cfg, include_notes=args.notes)
    ambiguous = sum(e.ambiguous for e in deck.elements)
    print(f"{deck.slide_count} slides, {len(deck.elements)} elements, {ambiguous} ambiguous -> {args.out}/")


if __name__ == "__main__":
    main()
