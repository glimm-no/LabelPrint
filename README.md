# LabelPrint

LabelPrint is a Windows command-line utility for printing PDF files directly to TSPL-compatible label printers (such as Rollo, Munbyn, Zebra, and generic thermal printers). It supports printing via standard Windows printer queues or directly over serial COM ports (useful for Bluetooth SPP printers).

## Features

- **Direct PDF Printing:** Converts PDF pages to monochrome TSPL raster graphics and sends them directly to the printer.
- **Watch Folder Mode:** Automatically monitors a directory for new PDF files and prints them as they arrive.
- **Virtual Printer Integration:** Installs a silent "Label Printer" in Windows. You can print to this virtual printer from any application (browser, Word, etc.), and it will automatically route the job to your physical label printer.
- **Startup Task:** Easily register the watch mode to start automatically when you log into Windows.
- **Bluetooth Support:** Send raw TSPL commands directly to a COM port for Bluetooth thermal printers.

## Usage

### 1. Find your printer
List all installed Windows printers and available COM ports:
```cmd
LabelPrint --list-printers
```

### 2. Print a single PDF
Print a 100x150mm (4x6 inch) label to a Windows printer:
```cmd
LabelPrint --pdf "C:\path\to\label.pdf" --printer "Rollo Printer" --width-mm 100 --height-mm 150
```

Print to a Bluetooth printer on COM4:
```cmd
LabelPrint --pdf "C:\path\to\label.pdf" --port COM4 --baud 9600 --width-mm 100 --height-mm 150
```

### 3. Watch Folder Mode
Monitor a folder and print any PDFs dropped into it automatically:
```cmd
LabelPrint --watch --watch-folder "C:\LabelQueue" --printer "Rollo Printer"
```

### 4. Install Virtual Printer (Requires Admin)
Creates a virtual Windows printer named "Label Printer". Any document printed to this virtual printer is saved as a PDF to the watch folder and immediately printed to your physical printer.
```cmd
LabelPrint --install --watch-folder "C:\LabelQueue" --printer "Rollo Printer"
```
*Note: You must run this command as Administrator. Once installed, you can run the watcher normally or register it as a startup task.*

### 5. Run on Startup
Register the watcher to start automatically when you log in:
```cmd
LabelPrint --register-startup --watch-folder "C:\LabelQueue" --printer "Rollo Printer"
```

## Command-Line Options

### Transport (one required for print/watch/install)
* `--printer <name>`: Windows printer name.
* `--port <COMn>`: COM port for Bluetooth SPP printers (e.g., COM4).
* `--baud <n>`: Baud rate for `--port` (default 9600).

### Watch / Install Options
* `--watch-folder <dir>`: Folder to watch for new PDF files.
* `--virtual-printer-name <name>`: Name of the virtual printer (default: 'Label Printer').

### Label Options
* `--dpi <n>`: Render DPI (defaults to printer capability or 203).
* `--width-mm <n>`: Label width in mm (default 100).
* `--height-mm <n>`: Label height in mm (default 150).
* `--page <n>`: 1-based page index (default 1).
* `--threshold <n>`: 0-255 threshold for black/white conversion (default 180).
* `--gap-mm <n>`: Gap in mm (defaults to printer setting).
* `--gap-offset-mm <n>`: Gap offset in mm (defaults to printer setting).
* `--offset-mm <n>`: Vertical label offset in mm (defaults to printer setting).
* `--speed <n>`: Print speed (TSPL SPEED).
* `--density <n>`: Print density (TSPL DENSITY).
* `--direction <n>`: Print direction 0 or 1 (defaults to printer setting).
* `--tear`: Enable tear mode (default).
* `--no-tear`: Disable tear mode.
* `--peel`: Enable peel mode.
* `--feed <n>`: Feed n mm after print.

## Requirements
* Windows 10 or later.
* .NET 8.0 Runtime.
* For the `--install` feature, the "Microsoft Print to PDF" driver must be enabled in Windows Features.
