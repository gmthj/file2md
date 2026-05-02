import argparse
import os
import sys

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from .converter import convert, is_supported


def collect_files(paths):
    files = []
    for path in paths:
        if os.path.isfile(path):
            files.append(path)
        elif os.path.isdir(path):
            for root, _dirs, filenames in os.walk(path):
                for filename in sorted(filenames):
                    filepath = os.path.join(root, filename)
                    if is_supported(filepath):
                        files.append(filepath)
        else:
            print(f"Error: '{path}' not found", file=sys.stderr)
            sys.exit(1)
    return files


def main():
    parser = argparse.ArgumentParser(
        description="Convert files to lightweight Markdown"
    )
    parser.add_argument(
        "paths",
        nargs="+",
        help="Files or directories to convert",
    )
    parser.add_argument(
        "-o", "--output",
        help="Output file path (only valid for single-file input)",
    )
    parser.add_argument(
        "--stdout",
        action="store_true",
        help="Print output to stdout instead of writing files",
    )
    args = parser.parse_args()

    if args.output and len(args.paths) > 1:
        print("Error: -o/--output can only be used with a single file", file=sys.stderr)
        sys.exit(1)

    files = collect_files(args.paths)

    for filepath in files:
        try:
            md = convert(filepath)
        except Exception as e:
            print(f"Error converting {filepath}: {e}", file=sys.stderr)
            sys.exit(1)

        if args.stdout:
            header = f"\n{'='*60}\n## {filepath}\n{'='*60}\n" if len(files) > 1 else ""
            print(f"{header}{md}")
        elif args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(md)
        else:
            out_path = os.path.splitext(filepath)[0] + ".md"
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(md)
            print(f"Converted: {filepath} -> {out_path}")


if __name__ == "__main__":
    main()
