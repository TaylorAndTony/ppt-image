# Make PPT support image output

## How it works

For windows, open registry and navigate to `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options`, the 16.0 varies depending on Microsoft Office versions.

Create a 32-bit DWORD value named `ExportBitmapResolution` then set its value to decimal `300`. This 300 is image output dpi.

Go to your PPT and you'll find image format at "Save as" window.

## Usage

Clone this repo and start `hackPPT.py`, follow the instructions and enjoy :)

## Python Requirements

- `rich`, for colorful printing

install: `pip install rich`
