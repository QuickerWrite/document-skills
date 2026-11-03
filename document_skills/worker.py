"""One-shot worker: binary OOXML on stdin, PDF on stdout; no DB credentials."""
import argparse
import sys
from .render import _libreoffice_pdf, MAX_BYTES

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--format", choices=["docx", "pptx", "xlsx"], required=True)
    args = parser.parse_args()
    try:
        result = _libreoffice_pdf(sys.stdin.buffer.read(MAX_BYTES + 1), "." + args.format)
        sys.stdout.buffer.write(result)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)

if __name__ == "__main__":
    main()
