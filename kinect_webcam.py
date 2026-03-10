"""
Kinect v1 (Xbox 360) — Virtual Webcam
======================================
Reads the Kinect color camera and streams it as a virtual webcam.
In Discord / Teams / Zoom select: "OBS Virtual Camera"

Requirements:
    - Windows 10/11 x64
    - Kinect SDK 1.8: https://www.microsoft.com/en-us/download/details.aspx?id=40278
    - Python 3.8+ (64-bit)
    - pip install numpy pyvirtualcam
    - OBS Studio 28+: https://obsproject.com/
    - opencv-python (optional, for preview): pip install opencv-python

Usage:
    python kinect_webcam.py               # stream to OBS Virtual Camera
    python kinect_webcam.py --preview     # stream + show local preview window
    python kinect_webcam.py --no-vcam     # local preview only, no OBS needed
    python kinect_webcam.py --flip        # mirror image horizontally
    python kinect_webcam.py --fps 15      # lower FPS (saves CPU)
"""

import sys
import ctypes
import time
import argparse
import struct
import numpy as np

# ──────────────────────────────────────────────────────────────
#  SYSTEM CHECKS
# ──────────────────────────────────────────────────────────────
if sys.platform != "win32":
    print("ERROR: This script only runs on Windows.")
    sys.exit(1)

if struct.calcsize("P") != 8:
    print("ERROR: 64-bit Python is required.")
    print("  Download: https://www.python.org/downloads/")
    sys.exit(1)

# ──────────────────────────────────────────────────────────────
#  ARGUMENTS
# ──────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(
    description="Kinect v1 Xbox 360 -> Virtual Webcam (OBS)"
)
parser.add_argument("--preview", action="store_true",
                    help="Show a local preview window (requires opencv-python)")
parser.add_argument("--no-vcam", action="store_true",
                    help="Local preview only, skip OBS Virtual Camera")
parser.add_argument("--flip",    action="store_true",
                    help="Mirror image horizontally")
parser.add_argument("--fps",     type=int, default=30,
                    help="Target FPS (default: 30)")
args = parser.parse_args()

USE_VCAM = not args.no_vcam
PREVIEW  = args.preview or args.no_vcam

# ──────────────────────────────────────────────────────────────
#  DEPENDENCIES
# ──────────────────────────────────────────────────────────────
if USE_VCAM:
    try:
        import pyvirtualcam
    except ImportError:
        print("ERROR: pyvirtualcam is not installed.")
        print("  pip install pyvirtualcam")
        print("  Also install OBS Studio 28+: https://obsproject.com/")
        print("")
        print("Tip: use --no-vcam to test the camera without OBS.")
        sys.exit(1)

HAS_CV2 = False
if PREVIEW:
    try:
        import cv2
        HAS_CV2 = True
    except ImportError:
        print("INFO: opencv-python not found — preview unavailable.")
        print("  pip install opencv-python")
        PREVIEW = False

# ──────────────────────────────────────────────────────────────
#  NUI CONSTANTS
#  Values determined experimentally for the Kinect10.dll flat API.
#  These differ from the COM/vtable values in the official SDK docs.
# ──────────────────────────────────────────────────────────────
NUI_INITIALIZE_FLAG_USES_COLOR = 0x00000002
NUI_IMAGE_TYPE_COLOR           = 1
NUI_IMAGE_RESOLUTION_640x480   = 2

CAM_W = 640
CAM_H = 480

# INuiFrameTexture vtable indices (Kinect SDK 1.8, x64)
# IUnknown:         [0] QueryInterface  [1] AddRef  [2] Release
# INuiFrameTexture: [3] BufferLen       [4] Pitch   [5] LockRect  [6] UnlockRect
VTBL_LOCK_RECT   = 5
VTBL_UNLOCK_RECT = 6

HRESULT = ctypes.c_long

# ──────────────────────────────────────────────────────────────
#  STRUCTURES
# ──────────────────────────────────────────────────────────────
class NUI_LOCKED_RECT(ctypes.Structure):
    _fields_ = [
        ("Pitch", ctypes.c_int),
        ("size",  ctypes.c_int),
        ("pBits", ctypes.c_void_p),
    ]

# NUI_IMAGE_FRAME memory layout (flat API, x64, SDK 1.8):
#   offset  0: liTimeStamp     (int64,  8 bytes)
#   offset  8: dwFrameNumber   (uint32, 4 bytes)
#   offset 12: eImageType      (uint32, 4 bytes)
#   offset 16: eResolution     (uint32, 4 bytes)
#   offset 20: _pad            (uint32, 4 bytes)  -- alignment padding
#   offset 24: pFrameTexture   (ptr,    8 bytes)  -- INuiFrameTexture*
#   offset 32: dwFrameFlags    (uint32, 4 bytes)
#   offset 36: ViewArea        (float*4,16 bytes)
class NUI_IMAGE_FRAME(ctypes.Structure):
    _fields_ = [
        ("liTimeStamp",   ctypes.c_longlong),
        ("dwFrameNumber", ctypes.c_ulong),
        ("eImageType",    ctypes.c_uint),
        ("eResolution",   ctypes.c_uint),
        ("_pad",          ctypes.c_uint),
        ("pFrameTexture", ctypes.c_void_p),
        ("dwFrameFlags",  ctypes.c_ulong),
        ("ViewArea",      ctypes.c_float * 4),
    ]

def _verify_frame_layout():
    """Assert that pFrameTexture is at the expected memory offset at runtime."""
    offset = NUI_IMAGE_FRAME.pFrameTexture.offset
    assert offset == 24, (
        f"NUI_IMAGE_FRAME layout mismatch: "
        f"pFrameTexture is at offset {offset}, expected 24. "
        f"Please open a GitHub issue."
    )

# ──────────────────────────────────────────────────────────────
#  DLL SETUP
# ──────────────────────────────────────────────────────────────
def setup_dll():
    try:
        dll = ctypes.WinDLL("Kinect10")
    except OSError:
        print("ERROR: Could not load Kinect10.dll")
        print("  Install Kinect SDK 1.8:")
        print("  https://www.microsoft.com/en-us/download/details.aspx?id=40278")
        sys.exit(1)

    dll.NuiInitialize.restype  = HRESULT
    dll.NuiInitialize.argtypes = [ctypes.c_ulong]

    dll.NuiShutdown.restype  = None
    dll.NuiShutdown.argtypes = []

    dll.NuiImageStreamOpen.restype  = HRESULT
    dll.NuiImageStreamOpen.argtypes = [
        ctypes.c_uint,                    # eImageType
        ctypes.c_uint,                    # eResolution
        ctypes.c_ulong,                   # dwImageFrameFlags
        ctypes.c_ulong,                   # dwFrameLimit
        ctypes.c_void_p,                  # hNextFrameEvent (NULL)
        ctypes.POINTER(ctypes.c_void_p),  # phStreamHandle [out]
    ]

    dll.NuiImageStreamGetNextFrame.restype  = HRESULT
    dll.NuiImageStreamGetNextFrame.argtypes = [
        ctypes.c_void_p,                  # hStream
        ctypes.c_ulong,                   # dwMillisecondsToWait
        ctypes.POINTER(ctypes.c_void_p),  # ppcImageFrame [out] (pointer-to-pointer)
    ]

    dll.NuiImageStreamReleaseFrame.restype  = HRESULT
    dll.NuiImageStreamReleaseFrame.argtypes = [
        ctypes.c_void_p,                  # hStream
        ctypes.c_void_p,                  # pImageFrame
    ]

    return dll

# ──────────────────────────────────────────────────────────────
#  FRAME READING
# ──────────────────────────────────────────────────────────────
def _get_vtfn(obj_ptr, index, restype, argtypes):
    """Resolve a COM vtable method pointer."""
    vtable = ctypes.cast(obj_ptr, ctypes.POINTER(ctypes.c_void_p))[0]
    fn_ptr = ctypes.cast(vtable,  ctypes.POINTER(ctypes.c_void_p))[index]
    return ctypes.cast(fn_ptr, ctypes.CFUNCTYPE(restype, ctypes.c_void_p, *argtypes))

def read_texture_rgb(texture_ptr, w, h):
    """
    Read pixels from an INuiFrameTexture via LockRect / UnlockRect.
    Returns a numpy array (H, W, 3) in RGB order, or None on failure.
    """
    locked = NUI_LOCKED_RECT()

    lock_fn = _get_vtfn(texture_ptr, VTBL_LOCK_RECT, HRESULT, [
        ctypes.c_uint,                    # Level (always 0)
        ctypes.POINTER(NUI_LOCKED_RECT),  # pLockedRect [out]
        ctypes.c_void_p,                  # pRect (NULL = entire buffer)
        ctypes.c_ulong,                   # Flags (always 0)
    ])

    hr = lock_fn(texture_ptr, 0, ctypes.byref(locked), None, 0)
    if hr != 0 or not locked.pBits or locked.Pitch <= 0:
        return None

    try:
        pitch    = locked.Pitch           # bytes per row (may be > w*4)
        buf_size = pitch * h
        raw = (ctypes.c_ubyte * buf_size).from_address(locked.pBits)
        arr = np.frombuffer(raw, dtype=np.uint8).copy()
        arr = arr.reshape((h, pitch))
        arr = arr[:, :w * 4].reshape((h, w, 4))  # drop row padding
        return arr[:, :, [2, 1, 0]].copy()         # BGRX -> RGB
    finally:
        unlock_fn = _get_vtfn(texture_ptr, VTBL_UNLOCK_RECT, HRESULT, [])
        unlock_fn(texture_ptr)

# ──────────────────────────────────────────────────────────────
#  MAIN LOOP
# ──────────────────────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  Kinect v1 Xbox 360 — Virtual Webcam")
    print("=" * 55 + "\n")

    _verify_frame_layout()
    dll = setup_dll()

    # Initialize Kinect
    hr = dll.NuiInitialize(NUI_INITIALIZE_FLAG_USES_COLOR)
    if hr != 0:
        u = hr & 0xFFFFFFFF
        print(f"ERROR: NuiInitialize failed: 0x{u:08X}")
        if u == 0x80070015:
            print("  The camera driver is not responding (NOT_READY).")
            print("  Check Device Manager: 'Kinect for Windows Camera' should show OK,")
            print("  not as Unknown or with a yellow warning icon.")
            print("  See README for driver fix instructions.")
        sys.exit(1)
    print("✓ NuiInitialize OK")

    # Open color stream
    h_stream = ctypes.c_void_p(None)
    hr = dll.NuiImageStreamOpen(
        NUI_IMAGE_TYPE_COLOR,
        NUI_IMAGE_RESOLUTION_640x480,
        0,    # dwImageFrameFlags
        2,    # dwFrameLimit (2-frame buffer)
        None, # hNextFrameEvent
        ctypes.byref(h_stream)
    )
    if hr != 0 or not h_stream.value:
        print(f"ERROR: NuiImageStreamOpen failed: 0x{hr & 0xFFFFFFFF:08X}")
        dll.NuiShutdown()
        sys.exit(1)
    print(f"✓ Color stream opened ({CAM_W}x{CAM_H})\n")

    if USE_VCAM:
        print("Starting OBS Virtual Camera...")
        print("-> In Discord / Teams select: 'OBS Virtual Camera'\n")
    if PREVIEW and HAS_CV2:
        print("Preview: press Q to close the window\n")
    print("Press Ctrl+C to stop.\n")

    black  = np.zeros((CAM_H, CAM_W, 3), dtype=np.uint8)
    ok_n   = 0
    fail_n = 0
    t_last = time.time()

    def process_frame(send_frame):
        """Grab one frame, process it, and send it. Returns False if the loop should stop."""
        nonlocal ok_n, fail_n, t_last

        frame_ptr = ctypes.c_void_p(None)
        hr = dll.NuiImageStreamGetNextFrame(h_stream, 100, ctypes.byref(frame_ptr))

        img = None
        if hr == 0 and frame_ptr.value:
            frame = NUI_IMAGE_FRAME.from_address(frame_ptr.value)
            if frame.pFrameTexture:
                try:
                    img = read_texture_rgb(frame.pFrameTexture, CAM_W, CAM_H)
                except Exception as e:
                    print(f"Texture read error: {e}")
            dll.NuiImageStreamReleaseFrame(h_stream, frame_ptr.value)

        if img is not None:
            if args.flip:
                img = img[:, ::-1, :]
            send_frame(img)
            ok_n += 1

            if PREVIEW and HAS_CV2:
                cv2.imshow("Kinect — press Q to quit", img[:, :, ::-1])
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    return False
        else:
            send_frame(black)
            fail_n += 1

        # Print a warning every 10 seconds if frame rate is poor
        now = time.time()
        if now - t_last >= 10.0:
            total = ok_n + fail_n
            pct   = int(100 * ok_n / total) if total > 0 else 0
            if pct < 50:
                print(f"  [warning] Low frame rate: {ok_n} OK / {fail_n} failed ({pct}%)")
            ok_n = fail_n = 0
            t_last = now

        return True

    try:
        if USE_VCAM:
            with pyvirtualcam.Camera(
                width=CAM_W, height=CAM_H, fps=args.fps,
                fmt=pyvirtualcam.PixelFormat.RGB
            ) as cam:
                print(f"Virtual camera active: {cam.device}")
                print("Ready!\n")
                while True:
                    if not process_frame(cam.send):
                        break
                    cam.sleep_until_next_frame()
        else:
            frame_time = 1.0 / args.fps
            while True:
                t0 = time.time()
                if not process_frame(lambda f: None):
                    break
                remaining = frame_time - (time.time() - t0)
                if remaining > 0:
                    time.sleep(remaining)

    except KeyboardInterrupt:
        print("\nStopped.")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if PREVIEW and HAS_CV2:
            cv2.destroyAllWindows()
        dll.NuiShutdown()
        print("Kinect closed.")


if __name__ == "__main__":
    main()
