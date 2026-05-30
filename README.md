# Conversion Studio

An open-source, high-fidelity universal media and document conversion desktop application built with Python 3.13 and PyQt6. Designed specifically for content creators, video editors, and power users who require **100% original, lossless quality** transformations completely offline. No file compression, no cloud tracking, no data leaks.

---

## 🚀 Download & Installation

For quick deployment on Windows, download the standard standalone setup installer compiled via Inno Setup:

### 📦 [Download Standalone Windows Installer (.exe)](https://github.com/Gauruuu/conversion-studio/releases/latest)

> **Installation Note:** Run the setup file, follow the wizard instructions, and launch the application directly from your desktop shortcut icon. No Python configuration or environment variable tweaks are needed for standard execution!

---

## ✨ Features & Processing Engines

### 📄 Document Deck
* **Word to PDF:** Retains original layouts, embedded formatting, and vector graphics using native OS hooks.
* **PPTX to PDF:** Direct high-fidelity presentation rendering.
* **PDF to Word:** Locally parses PDF structural layers and reconstructs paragraphs, tables, and typography into fully editable `.docx` master files.

### 🖼️ Image Deck
* **Image to ICO (App Icon):** Converts source graphics into a multi-layered Windows desktop icon system (automatically embedding `16px` through `256px` sizes inside a single file).
* **Mass Images to PDF:** Compiles batches of separate images into a single, unified PDF document without quality degradation.
* **HEIC to JPG:** Decodes mobile Apple camera streams straight to native JPEGs at full color-depth scaling, applying maximum quality flags and disabling chroma subsampling.
* **WebP/PNG/JPG Inter-conversions:** Lossless conversion profiles that fully maintain structural alpha transparency maps.

### 🎬 Video Studio (DaVinci Resolve Optimized)
* **MOV to MP4:** Fixes the infamous "Media Offline" error in DaVinci Resolve Free on Windows. Transcodes unsupported Apple video streams into high-bitrate (`-crf 14`), production-ready H.264 profiles with standard YUV420p pixel layouts and AAC high-fidelity audio.
* **MKV to MP4:** 100% lossless remuxing. Copies underlying bitstreams natively without re-encoding, taking less than 3 seconds for massive multi-gigabyte video files.
* **Video to GIF:** Converts video frames using a specialized two-pass Lanczos spatial filter for high-color depth, crisp, non-dithered animation sequences.

### 🎵 Audio Station
* **WAV to MP3:** Encodes master audio tracks directly into maximum constant bitrates (`320kbps`).
* **M4A to WAV & MP3 to WAV:** Unpacks compressed audio bitstreams natively back into fully uncompressed raw 16-bit Pulse Code Modulation linear structural tracks (`pcm_s16le`).

---

## 🛠️ Developer Setup & Modifying Source Code

Since this project is completely open source, you can clone this repository, tweak the source pipeline algorithms, or introduce custom processing features.

### Prerequisites
1. **Python 3.13+** installed on your workstation environment.
2. **FFmpeg** binaries available inside your operating system's PATH environment variables.

### Local Installation
Clone the repository and install the development dependencies:
```bash
git clone [https://github.com/gauruuu/conversion-studio.git](https://github.com/gauruuu/conversion-studio.git)
cd conversion-studio
pip install PyQt6 pillow pillow-heif comtypes pdf2docx pyinstaller
