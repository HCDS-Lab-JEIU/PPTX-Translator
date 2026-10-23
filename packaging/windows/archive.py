"""Archive a built folder without PowerShell Compress-Archive exclusive reads."""
from pathlib import Path
import sys
from zipfile import ZipFile, ZIP_DEFLATED


def main():
    bundle, destination = (Path(value).resolve() for value in sys.argv[1:])
    if not (bundle / "PPTTranslator.exe").is_file():
        raise FileNotFoundError("The distribution executable is missing.")
    temporary = destination.with_suffix(".zip.partial")
    try:
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as archive:
            for path in sorted(bundle.rglob("*")):
                if path.is_file():
                    archive.write(path, path.relative_to(bundle.parent))
        temporary.replace(destination)
    finally:
        temporary.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
