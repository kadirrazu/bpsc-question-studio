"""Build the slim Windows distribution from official Python archives and wheels."""

import argparse
import shutil
from pathlib import Path
from zipfile import ZipFile

SOURCE = Path(__file__).resolve().parents[1]
QT_DLLS = {"Qt6Core.dll", "Qt6Gui.dll", "Qt6Widgets.dll", "Qt6Svg.dll", "pyside6.abi3.dll"}
QT_BINDINGS = {"QtCore.pyd", "QtGui.pyd", "QtWidgets.pyd"}
QT_PLUGINS = {
    "plugins/platforms/qwindows.dll",
    "plugins/iconengines/qsvgicon.dll",
    "plugins/imageformats/qsvg.dll",
    "plugins/imageformats/qico.dll",
    "plugins/imageformats/qjpeg.dll",
}


def include_wheel_file(name):
    if not name.startswith("PySide6/"):
        return True
    relative = name.removeprefix("PySide6/")
    item = Path(relative)
    if relative.startswith("support/"):
        return True
    if relative in QT_BINDINGS or relative in QT_DLLS or relative in QT_PLUGINS:
        return True
    if len(item.parts) == 1:
        return item.suffix.lower() not in (".dll", ".pyd", ".lib", ".exe", ".pyi") or (
            item.suffix.lower() == ".dll" and (
                item.name.startswith(("msvcp", "vcruntime", "d3d", "icu"))
                or item.name in {"concrt140.dll", "opengl32sw.dll"}
            )
        )
    return False


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--python-zip", type=Path, required=True)
    parser.add_argument("--wheels", type=Path, required=True)
    parser.add_argument("--launcher", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    target = args.output
    target.mkdir(parents=True, exist_ok=False)
    for filename in ("README.md", "Run-Diagnostics.bat", "main.py"):
        shutil.copy2(SOURCE / filename, target)
    for folder in ("docs", "examples", "mcq", "assets"):
        shutil.copytree(SOURCE / folder, target / folder, ignore=shutil.ignore_patterns("__pycache__"))
    default_templates = target / "Config" / "Template"
    default_templates.mkdir(parents=True)
    for kind in ("HEADER", "FOOTER"):
        filename = f"DEFAULT_{kind}.docx"
        shutil.copy2(SOURCE / "Config" / "Template" / filename, default_templates)
    shutil.copy2(args.launcher, target / "BPSC Question Studio.exe")
    runtime = target / "Runtime"
    modules = runtime / "packages"
    with ZipFile(args.python_zip) as archive:
        archive.extractall(runtime)
    for wheel in sorted(args.wheels.glob("*.whl")):
        with ZipFile(wheel) as archive:
            archive.extractall(modules, [name for name in archive.namelist() if include_wheel_file(name)])
    path_file = next(runtime.glob("python*._pth"))
    stdlib = next(runtime.glob("python*.zip"))
    path_file.write_text(f"{stdlib.name}\n.\n..\npackages\nimport site\n", encoding="utf-8")


if __name__ == "__main__":
    main()
