# BPSC Question Studio

BPSC Question Studio is an offline desktop app for the Bangladesh Public Service Commission. It runs on 64-bit Windows.

- **Set Generation** creates question sets, a mapping, and print-ready DOCX or PDF files.
- **Center File Wizard** adds the columns needed for printing to a copy of your center workbook.
- **Room Print Packs** uses the workbook and question-set PDFs to make a print pack for each room.

See the [User Guide](docs/guides/USER-GUIDE.md) for step-by-step instructions.

## What you need to build it

- Windows 10 or 11 (64-bit) and [Python 3.12 for Windows](https://www.python.org/downloads/windows/). Install the 64-bit version and make sure the Python Launcher is available.
- An internet connection for the first build. The build script downloads the compiler, the Python runtime, and the required packages.
- [LibreOffice](https://www.libreoffice.org/download/) and the fonts used in your question papers to export question sets as PDFs. If the app cannot find LibreOffice, open **Common Settings → Renderer** and select `program\soffice.com` from its installation folder.

Open Command Prompt and check that the launcher can find Python 3.12.

```bat
py -3.12 --version
```

## Build the Windows portable app

1. Extract the complete Source ZIP. Open Command Prompt in the folder that contains `Build-Portable.bat`.
2. Run the build script. It compiles the launcher and puts the app and its runtime in `dist\BPSC-Question-Studio`.

   ```bat
   Build-Portable.bat
   ```

3. Start the app from the same Command Prompt.

   ```bat
   "dist\BPSC-Question-Studio\BPSC Question Studio.exe"
   ```

To use the app on another Windows PC, copy or zip the whole `dist\BPSC-Question-Studio` folder. Keep the files together. The folder includes Python, so you do not need to install Python on the other PC.

Before rebuilding, rename or move the existing `dist\BPSC-Question-Studio` folder. The build script creates a new one.

## Run from source

From the project folder, run this in Command Prompt.

```bat
Run-Source.bat
```

See the [Windows Build Guide](docs/guides/BUILD-WINDOWS.md) for manual build steps, tests, and troubleshooting. For changes in each version, see the [release history](docs/version-history/CHANGELOG.md).
