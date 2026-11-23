# BPSC Question Studio Release History

The latest release is at the top. See the [guides](../guides/) for help using or building the app.

## v0.5.1

- The sign-in form fills in the default email. You can change it before signing in.
- A progress bar shows the current step and elapsed time while question sets are being made. LibreOffice starts with a fresh profile for each conversion batch. If it times out, the message shows the renderer path.
- The portable app uses the BPSC Question Studio name on Windows.
- Renamed the set footer label to **Unique Alphanumeric Code for Each Set**. Room codes now default to `{code}`, with settings for font size, color, and opacity.
- Gave the Center File Wizard a new layout with separate panels for the source workbook, generation rules, file-name fields, and live preview.

## v0.5.0

- The default QR width is 0.75 inches. You can change it in settings.
- You can turn the code in the set footer on or off and set its font, size, color, and opacity.
- Added the Center File Wizard. It previews question counts, room numbers in sequence for each center, file names, print instructions, and optional room codes. It saves the new columns in a copy of the workbook.
- Room Print Packs can print room codes with a custom label. You can choose the position and the pages they appear on.

## v0.4.4

- Each run saves the room PDFs to a new `YYYY-MM-DD-HHMMSS_CENTER_WISE_GENERATED_FILES` folder. The name uses the date and time in Bangladesh.

## v0.4.3

- The cover and end pages use light gray shading for their headings and the left column of each table.
- Each page has its own option to print the heading, table labels, and values in uppercase.
- The footer timestamp starts with `Generated on:`. The end page also has a new default note.

## v0.4.2

- Room Print Packs uses black-and-white headings and table borders on the cover and end pages.
- Both pages can show a centered timestamp and software credit. You can change the footer font size and opacity.
- The console reports one result per room and shows the total time at the end.
- The source ZIP includes a README, separate guides, and a release history.

## v0.4.1

- Room Print Packs can pick up available set PDFs from the open exam and remember your last settings.
- The cover and end pages have separate tabs. The tables are easier to read, and the cover has a footer.
- Improved input spacing, checkboxes, scrolling, and the console.

## v0.4.0

- Added Room Print Packs. It makes print files for each center and room using an Excel room list and the question-set PDFs.
- You can add cover and end pages, insert blank pages for double-sided printing, follow progress in the console, and regenerate PDFs.
- The app writes `Done` to Excel for each PDF it saves. It keeps formulas and their saved results.

## v0.3.6

- Added **Choose Font** so you can pick a font installed on your computer.
- You can choose Bangla or English digits for set numbers separately in the runner and in templates and QR codes.

## v0.3.5

- Rotated set numbers and code names have separate paragraph and table-cell alignment settings.
- Improved how the app picks and saves fonts for Generated Text, Runner, and Footer Text.

## v0.3.4

- QR insets and gaps accept negative values. You can rotate the set number and code name.
- Common Settings buttons fit into two rows, font sizes change in 0.5 pt steps, and exam folders are organized by year.

## v0.3.3

- Improved QR placement, runner and footer text settings, default templates, and the way the app shows exam details.

## v0.3.2

- Improved QR alignment, how post grades appear, and how generation status is restored when you reopen an exam.
