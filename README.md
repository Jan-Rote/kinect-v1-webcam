# kinect-v1-webcam

Use your **Xbox 360 Kinect (v1)** as a webcam on Windows — stream the color camera to Discord, Teams, Zoom, or any app that supports virtual cameras.

![Python](https://img.shields.io/badge/python-3.8%2B%20(64--bit)-blue)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

> Tested on Windows 11 x64 with Kinect SDK 1.8 and Python 3.11.

## ⚠️ Warning
Okay, This whole small program and readme HAS been made with AI. (Claude Sonnet 4.6) I respect developers and their human code, posting this here I am just trying to help anyone else having the same problem as me, Having all of the guides for "Use Kinect for xbox 360 as a webcam!" Not working in the current age. All of the code works (For me), but if you want to improve it, go for it! Okay, now you can go on and read the actually quite great AI Slop below.

---

## How it works

The script reads the Kinect color camera directly through `Kinect10.dll` (Kinect SDK 1.8) using ctypes, then pushes frames to **OBS Virtual Camera** via `pyvirtualcam`. Any app that supports camera input will see it as a normal webcam.

```
Kinect v1 (USB)
    └── Kinect10.dll  (Kinect SDK 1.8)
        └── kinect_webcam.py
            └── pyvirtualcam
                └── OBS Virtual Camera
                    └── Discord / Teams / Zoom / OBS
```

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows 10/11 x64 | 32-bit not supported |
| [Kinect SDK 1.8](https://www.microsoft.com/en-us/download/details.aspx?id=40278) | Install **before** connecting the Kinect |
| Python 3.8+ **64-bit** | 32-bit Python will not work |
| [OBS Studio 28+](https://obsproject.com/) | Provides the Virtual Camera driver |
| `numpy` | `pip install numpy` |
| `pyvirtualcam` | `pip install pyvirtualcam` |
| `opencv-python` | Optional — local preview only: `pip install opencv-python` |

---

## Installation

**1. Install Kinect SDK 1.8**
Download and install from [microsoft.com](https://www.microsoft.com/en-us/download/details.aspx?id=40278). Do this **before** plugging in the Kinect.

**2. Install OBS Studio**
Download from [obsproject.com](https://obsproject.com/). Launch it once and click **Start Virtual Camera** in the bottom-right controls panel. The virtual camera driver will remain available even after closing OBS.

**3. Install Python dependencies**
```bash
pip install numpy pyvirtualcam opencv-python
```

**4. Clone this repo**
```bash
git clone https://github.com/your-username/kinect-v1-webcam.git
cd kinect-v1-webcam
```

---

## Usage

```bash
# Stream to OBS Virtual Camera
python kinect_webcam.py

# Stream + show a local preview window
python kinect_webcam.py --preview

# Local preview only, no OBS needed (good for testing)
python kinect_webcam.py --no-vcam

# Mirror the image horizontally
python kinect_webcam.py --flip

# Lower FPS to reduce CPU usage
python kinect_webcam.py --fps 15
```

**In Discord:** Settings → Voice & Video → Camera → select **OBS Virtual Camera**

---


## Files

| File | Description |
|---|---|
| `kinect_webcam.py` | Main script — streams Kinect color camera as a virtual webcam |

---

## Known limitations

- **Windows only** — Kinect SDK 1.8 is Windows-exclusive
- **640×480** color resolution (Kinect v1 hardware limit at 30 fps)
- Requires OBS Studio for virtual camera output
- Only one application can use the Kinect color stream at a time

---

## Contributing

Issues and pull requests are welcome. If it works (or doesn't) on your setup, feel free to open an issue with your Windows version, Python version, and the exact error message.

---

## License

MIT
