"""Check wheel/sdist contents before sharing; never extracts or executes archives."""

from __future__ import annotations

import argparse
import json
import tarfile
import zipfile
from pathlib import Path, PurePosixPath

FORBIDDEN_PARTS = {
    ".research",
    "artifacts",
    ".venv",
    ".uv-cache",
    ".git",
    ".codex",
    "__pycache__",
    ".pytest_cache",
    ".ruff_cache",
    ".env",
}
FORBIDDEN_SUFFIXES = {".exe", ".dll", ".dotm", ".ppam", ".afx", ".bas", ".pyc"}
SOURCE_ROOTS = {
    "src",
    "tests",
    "scripts",
    "config",
    "docs",
    ".github",
    "README.md",
    "pyproject.toml",
    "uv.lock",
    ".gitignore",
    ".gitattributes",
    "LICENSE",
    "LICENSE.txt",
    "NOTICE",
    "PKG-INFO",
}


def check_archive(path: Path) -> dict[str, object]:
    if path.suffix == ".whl":
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
        wheel = True
    elif path.name.endswith(".tar.gz"):
        with tarfile.open(path, "r:gz") as archive:
            names = []
            for member in archive.getmembers():
                if member.issym() or member.islnk():
                    raise ValueError(f"Archive contains a link: {member.name}")
                if member.isfile():
                    names.append(member.name)
        wheel = False
    else:
        raise ValueError(f"Unsupported archive: {path.name}")
    for name in names:
        p = PurePosixPath(name)
        if p.is_absolute() or ".." in p.parts:
            raise ValueError(f"Unsafe archive path: {name}")
        if any(part in FORBIDDEN_PARTS or part.startswith(".env.") for part in p.parts):
            raise ValueError(f"Local/private directory in archive: {name}")
        if p.suffix.lower() in FORBIDDEN_SUFFIXES:
            raise ValueError(f"Vendor/runtime file in archive: {name}")
        if wheel:
            if p.parts[0] != "axmath_mcp" and not p.parts[0].endswith(".dist-info"):
                raise ValueError(f"Unexpected wheel content: {name}")
        elif len(p.parts) < 2 or p.parts[1] not in SOURCE_ROOTS:
            raise ValueError(f"Unexpected source content: {name}")
    return {"archive": str(path), "files": len(names), "checked": True}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "directory", nargs="?", type=Path, default=Path(__file__).resolve().parents[1] / "dist"
    )
    args = parser.parse_args()
    files = sorted([*args.directory.glob("*.whl"), *args.directory.glob("*.tar.gz")])
    if not any(p.suffix == ".whl" for p in files) or not any(
        p.name.endswith(".tar.gz") for p in files
    ):
        raise SystemExit("Run uv build first: both wheel and source archive are required.")
    print(json.dumps([check_archive(p) for p in files], indent=2))


if __name__ == "__main__":
    main()
