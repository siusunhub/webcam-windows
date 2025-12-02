# Webcam Viewer (WebcamWindows)

A tiny Windows tray application that shows any webcam in a clean, borderless window.  
Originally built to use the **Blackmagic ATEM Mini Pro** USB-C output (set to Multiview) as a simple, always-visible webcam window, without all the extra UI and clutter of typical webcam/streaming software.

The app is written in C# using Windows Forms, OpenCvSharp and DirectShow.


## Why?

When using an **ATEM Mini Pro** as a USB-C webcam device in **Multiview** mode, Windows just sees it as “a webcam”. Most webcam programs:

- Add a lot of UI and controls around the video  
- Don’t handle aspect ratio nicely in a resizable window  
- Don’t offer an easy tray-based workflow with quick device / resolution switching  

So this project is a minimal viewer that:

- Picks the ATEM (or any webcam) as a standard camera device  
- Shows the feed in a clean, resizable window  
- Can run hidden in the tray and come back with one click  


## Features

- **Simple video window**
  - Single video panel that fills the window
  - Black background, no extra chrome or controls over the image

- **Multi-camera support**
  - Enumerates all DirectShow video devices (e.g. ATEM, USB webcams, capture cards)
  - Switch cameras from the tray context menu (`Select Webcam`)
  - Failed devices are marked as `[ERROR]` and can be refreshed

- **Resolution management**
  - Uses DirectShow to discover all supported video formats/resolutions
  - Automatically selects the **smallest** available resolution on startup (to keep CPU/bandwidth low)
  - Manual selection of resolution from the `Resolution` submenu
  - Separate “Refresh Resolutions” action if devices change at runtime

- **Aspect-ratio aware auto-resize**
  - When you resize the window, after a short delay the app auto-resizes so the **client area matches the camera aspect ratio** (as closely as possible)
  - Keeps the window inside the current screen’s working area
  - Tries to preserve the original top-left position

- **Tray icon & context menu**
  - Always-visible tray icon with context menu:
    - `Select Webcam`
    - `Resolution`
    - `Refresh Webcams`
    - `Enable / Disable Webcam`
    - `Full Screen`
    - `Always on Top`
    - `Minimize to Tray`
    - `Exit`

- **Window modes**
  - **Borderless mode**: double-click the video area to toggle bordered ↔ borderless
  - **Full screen mode**: use `Full Screen` in tray menu (on current monitor)
  - **Always on top**: optional, via tray menu

- **Minimize / close behavior**
  - On minimize, window hides and webcam can be disabled to free the device
  - Optional **“Minimize to Tray”**:
    - Clicking the window “X” hides to tray instead of exiting
    - Double-click tray icon to restore and re-enable webcam

- **Small touches**
  - Current time `[HH:mm:ss]` is shown in the window title bar
  - Title updates to include the active webcam name and “[DISABLED]” when the camera is off


## Tech stack

- **Language / UI**: C# + Windows Forms  
- **Video capture**: OpenCvSharp  
- **Device & format enumeration**: DirectShowLib  
- **Target OS**: Windows (10/11 recommended)


## Limitations / Notes

- **Single device at a time** – opens only one webcam concurrently.  
- **No audio** – video-only viewer, no audio capture.  
- **Not a virtual camera** – does not create a virtual webcam device; it just shows the real device in a window.  
- **Device in use** – while this app has the webcam open, other apps may not be able to use the same device.


## License

Add your preferred license here (MIT, GPL, etc.).
