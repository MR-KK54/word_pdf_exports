"""
Word Page Exporter Pro - Command Line Interface (CLI)
Provides command-line capabilities for batch document page extraction.
"""

import os
import sys
import glob
import argparse
from typing import List

from word_exporter_pro.core.com_engine import DocumentInspector
from word_exporter_pro.core.batch_processor import BatchProcessor, ExportJobConfig
from word_exporter_pro.utils.logger import get_logger

logger = get_logger()


def main():
    parser = argparse.ArgumentParser(
        description="Microsoft Word Page Exporter Pro - High-Fidelity Page Range Extraction Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Inspect document page count & metadata
  python run_cli.py -i "Report.docx" --inspect

  # Export pages 1 to 5 to PDF
  python run_cli.py -i "Report.docx" -r "1-5" -o "./output" -f pdf

  # Export multiple ranges (1-3, page 7, 10 to end)
  python run_cli.py -i "Report.docx" -r "1-3, 7, 10-end" -o "./output"

  # Split every single page into individual DOCX files
  python run_cli.py -i "Report.docx" -r "all-individual" -o "./split_pages"

  # Batch export all Word documents in a directory
  python run_cli.py -b "./documents/*.docx" -r "1-2" -o "./summaries" -f docx
"""
    )

    parser.add_argument("-i", "--input", nargs="+", help="Input Word document file path(s)")
    parser.add_argument("-b", "--batch", help="Batch glob pattern or directory path (e.g. './docs/*.docx')")
    parser.add_argument("-r", "--range", default="1-end", help="Page range specification (e.g., '1-3', '1, 5, 8-10', 'even', 'odd', 'all-individual'). Default: '1-end'")
    parser.add_argument("-o", "--output-dir", default="./exported_pages", help="Output directory path. Default: './exported_pages'")
    parser.add_argument("-f", "--format", choices=["same", "docx", "pdf", "doc", "rtf", "docm"], default="same", help="Export file format ('same' matches source file type). Default: 'same'")

    parser.add_argument("-n", "--naming", default="{original_name}_pages_{start_page}-{end_page}", help="Output filename template. Placeholders: {original_name}, {range}, {start_page}, {end_page}, {timestamp}")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite output files if they already exist")
    parser.add_argument("--mode", choices=["trimming", "selection"], default="trimming", help="Extraction engine mode. Default: 'trimming'")
    parser.add_argument("--inspect", action="store_true", help="Inspect input document statistics and exit")
    parser.add_argument("--visible", action="store_true", help="Run Microsoft Word visibly (useful for debugging)")

    args = parser.parse_args()

    # Collect source files
    source_files: List[str] = []
    if args.input:
        for f in args.input:
            if os.path.exists(f):
                source_files.append(os.path.abspath(f))
            else:
                logger.error(f"Input file not found: {f}")

    if args.batch:
        if os.path.isdir(args.batch):
            search_pattern = os.path.join(args.batch, "*.[dD][oO][cC]*")
            matched = glob.glob(search_pattern)
            source_files.extend([os.path.abspath(m) for m in matched if not os.path.basename(m).startswith("~$")])
        else:
            matched = glob.glob(args.batch)
            source_files.extend([os.path.abspath(m) for m in matched if not os.path.basename(m).startswith("~$")])

    source_files = sorted(list(set(source_files)))

    if not source_files:
        logger.error("No valid input files specified. Use -i or -b options.")
        parser.print_help()
        sys.exit(1)

    # Inspect Mode
    if args.inspect:
        print("\n" + "="*60)
        print(" DOCUMENT INSPECTION REPORT")
        print("="*60)
        for f in source_files:
            try:
                info = DocumentInspector.get_info(f, visible=args.visible)
                print(f"\nFile: {info['filename']}")
                print(f"Path: {info['path']}")
                print(f"Size: {info['size_bytes'] / 1024:.1f} KB")
                print(f"Pages: {info['page_count']}")
                print(f"Sections: {info['section_count']}")
                print(f"Format: {info['format'].upper()}")
                if info['title']:
                    print(f"Title: {info['title']}")
                if info['author']:
                    print(f"Author: {info['author']}")
            except Exception as e:
                print(f"Error inspecting '{f}': {e}")
        print("="*60 + "\n")
        sys.exit(0)

    # Export Mode
    config = ExportJobConfig(
        source_files=source_files,
        range_expression=args.range,
        output_dir=os.path.abspath(args.output_dir),
        export_format=args.format,
        naming_pattern=args.naming,
        overwrite=args.overwrite,
        engine_mode=args.mode,
        visible=args.visible
    )

    processor = BatchProcessor(config)

    def on_progress(completed: int, total: int, filename: str, status: str):
        pct = (completed / total) * 100 if total > 0 else 0
        print(f"\rProgress: [{completed}/{total}] ({pct:.1f}%) - {status}", end="", flush=True)

    def on_finished(success: int, fail: int, errors: List[str]):
        print("\n\n" + "="*60)
        print(f" EXPORT COMPLETED: {success} Succeeded, {fail} Failed")
        print("="*60)
        if errors:
            print("\nErrors encountered:")
            for err in errors:
                print(f" - {err}")
        print()

    processor._run_job(on_progress, on_finished)


if __name__ == "__main__":
    main()
