r"""Helper script to create a ZIP archive from the portable build folder.

Usage: python scripts/zip_portable.py <source_dir> <zip_path>
Example: python scripts/zip_portable.py dist\WordTextReplacer dist\WordTextReplacer_portable.zip
"""

import os
import sys
import zipfile


def create_zip(source_dir: str, zip_path: str) -> None:
    parent = os.path.dirname(source_dir)
    total_files = sum(len(files) for _, _, files in os.walk(source_dir))
    count = 0

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for dirpath, _, filenames in os.walk(source_dir):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                arcname = os.path.relpath(filepath, parent)
                zf.write(filepath, arcname)
                count += 1
                if count % 50 == 0:
                    print(f"      Packed {count}/{total_files} files...")

    size_mb = os.path.getsize(zip_path) / 1024 / 1024
    print(f"      Packed {count}/{total_files} files.")
    print(f"      Archive size: {size_mb:.1f} MB")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python scripts/zip_portable.py <source_dir> <zip_path>")
        sys.exit(1)

    src = sys.argv[1]
    out = sys.argv[2]

    if not os.path.isdir(src):
        print(f"Error: Source directory does not exist: {src}")
        sys.exit(1)

    create_zip(src, out)
