# Sovereign AI Workbench (v2.0)
> **Air-Gapped, Privacy-First, On-Device Industrial AI Workstation**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-brightgreen.svg)](https://www.python.org/)
[![Platform: Windows | Linux](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-lightgrey.svg)]()
[![Air-Gapped: Zero Telemetry](https://img.shields.io/badge/Security-100%25%20Air--Gapped-success.svg)]()

---

## Executive Summary

**Sovereign AI Workbench** is an enterprise-grade, zero-telemetry desktop application and runtime engine designed for high-security industrial and petrochemical environments (e.g., refinery operations, pipeline inspection, PLC automation, and air-gapped critical infrastructure).

It delivers completely local, self-hosted LLM inference, dynamic model routing, automated hardware matching, document parsing (PDF, Word, Excel, Images), and interactive chat/reasoning workflows without relying on external cloud APIs or exposing sensitive telemetry.

---

## Key Capabilities

- **100% Air-Gapped & Zero Telemetry**: Complete local execution with strictly zero outbound network requests.
- **Dynamic Hardware-Matched AI Recommender**: Automatic GPU/VRAM/RAM profiling that scores and recommends optimal open-weight models (1.5B to 70B).
- **One-Click Model Hub**: Built-in discovery and local streaming downloader for 30+ open-source models (DeepSeek-R1, Qwen 2.5 Coder, Llama 3.2 Vision, Mistral, Gemma, Phi-4).
- **Intelligent Auto-Router**: Automatically analyzes incoming user prompts and delegates them to specialized models (Coding, Deep Reasoning, Vision/OCR, or Fast Chat).
- **Multi-Modal Document Parsing**: Ingests PDFs, Word (.docx), Excel spreadsheets, logs, and technical drawings for instantaneous grounded analysis.
- **Executive Report Exporter**: One-click generation of professional engineering memos and reports directly into formatted `.pdf` or `.docx`.
- **Modern Minimalist UI**: Clean, responsive interface featuring multiple high-contrast luxury themes, smooth animations, and dark/light modes.

---

## Repository Structure

```text
├── main.py                     # Primary desktop entrypoint & local HTTP server
├── sovereign_engine.py         # Hardware profiling, model orchestration & inference engine
├── app_ui.html                 # Modern SPA desktop frontend
├── workbench.py                # Standalone lightweight runner
├── Run_Sovereign_AI_Workbench.bat  # 1-click Windows launcher
├── run_linux.sh                # 1-click Linux launcher
├── Install_Sovereign_AI_Workbench.bat # Automated environment setup script
├── assets/                     # Application logos, branding, and icons
│   ├── app_icon.png
│   ├── icon.ico
│   └── icon.png
├── docs/                       # Technical architecture & project guides
│   ├── PROJECT_DOCUMENTATION.md
│   ├── PPT_SLIDE_EDITING_GUIDE.md
│   └── sovereign_ai_workbench_video_script.txt
└── scripts/                    # Build utilities, installers & packagers
    ├── build_release_zip.py
    ├── build_app_and_installer.py
    └── setup_installer.py
```

---

## Quick Start Guide

### Prerequisites
- Windows 10/11 or Linux (Ubuntu 20.04+)
- [Python 3.10+](https://www.python.org/downloads/)
- [Ollama Runtime](https://ollama.com/) installed locally

### Option 1: Standalone Portable Bundle (No Python Required)
1. Download the pre-built portable distribution archive **`SovereignAIWorkbench_v2.0_Portable.zip`** (available under [GitHub Releases](https://github.com/shabbir-beep/SovereignAIWorkbench/releases)).
2. Extract the `.zip` archive to any directory or USB drive.
3. Double-click **`SovereignAIWorkbench.exe`** to launch the self-contained workstation immediately.

---

### Option 2: Running from Source
```bash
# 1. Clone repository
git clone https://github.com/shabbir-beep/SovereignAIWorkbench.git
cd SovereignAIWorkbench

# 2. Launch sandboxed desktop application
python main.py
```
*The native sandboxed desktop window will launch directly in an isolated process.*

### Option 3: Running via Batch Launcher (Windows)
Double-click **`Run_Sovereign_AI_Workbench.bat`** to start the local engine and desktop application instantly.

---

## Tech Stack & Architecture

- **Frontend**: HTML5, Tailwind CSS, Heroicons SVG Library, Marked.js (Markdown parser)
- **Desktop Runtime**: PyWebView / Chromium Embedded Framework
- **Backend Server**: Python 3 standard library HTTP daemon
- **Inference Engine**: Local Ollama Server API / OpenAI-compatible local endpoints
- **Hardware Telemetry**: Windows WMI / Linux `/proc` hardware inspection modules

---

## Security & Compliance

| Security Pillar | Implementation |
|---|---|
| **Data Ingestion** | All files parsed in isolated local memory buffers |
| **Model Weights** | Encrypted/sandboxed local storage (`data/models/`) |
| **Network Traffic** | Bound strictly to `127.0.0.1` (Local Loopback only) |
| **Session State** | Encrypted local JSON storage (`data/sessions/`) |

---

## Documentation & Links

- Detailed Architecture & Design: [`docs/PROJECT_DOCUMENTATION.md`](docs/PROJECT_DOCUMENTATION.md)
- Presentation & Slide Guide: [`docs/PPT_SLIDE_EDITING_GUIDE.md`](docs/PPT_SLIDE_EDITING_GUIDE.md)
- YouTube Video Walkthrough Script: [`docs/sovereign_ai_workbench_video_script.txt`](docs/sovereign_ai_workbench_video_script.txt)

---

## License
Distributed under the **MIT License**. See `LICENSE` for more information.
