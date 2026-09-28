# Sovereign AI Workbench — Complete Project Documentation
**Project Code:** SIH26117 | **Organization:** Mangalore Refinery and Petrochemicals Limited (MRPL)  
**Version:** 2.0 (Portable Edition) | **Supported Platforms:** Windows 10/11 (64-bit), Linux

---

## 1. Executive Summary & Project Overview

The **Sovereign AI Workbench** is a fully self-contained, air-gapped, portable desktop AI workstation designed for industrial, petrochemical, and enterprise environments. Built specifically for Mangalore Refinery and Petrochemicals Limited (MRPL), it allows non-technical operators and engineers to run state-of-the-art Large Language Models (LLMs) and Vision-Language Models locally without requiring an active internet connection, external cloud API keys, or administrative installation privileges.

### Core Capabilities
- **100% Air-Gapped & Sovereign:** Zero telemetry, cloud leakage, or external data transmission. All inferences, weights, logs, and sessions remain strictly on the local machine.
- **Embedded Portable AI Engine:** Automatically manages, embeds, and runs a standalone local Ollama server instance with completely isolated model storage.
- **Hardware-Aware Model Recommendation Engine:** Inspects real-time system hardware (CPU cores, RAM, free disk space, Nvidia GPU VRAM) and automatically scores and recommends the best model that comfortably fits the device.
- **Integrated Model Hub:** One-click downloads with real-time progress bars and cancellation support across a curated catalog of 32 models (Coding, Reasoning/Math, SOP/General, Vision, Enterprise).
- **VRAM & Process Lifecycle Management:** Employs Windows Job Objects (`KILL_ON_JOB_CLOSE`), automated orphan process cleanup on launch, and dynamic VRAM model unloading (`keep_alive: 0`) when switching models or shutting down.
- **Multi-File Context Attachment:** Native extraction and parsing for industrial document formats (DOCX, XLSX/XLS, PDF, Code files, logs, and images).
- **Executive Document & Deliverable Generator:** One-click export of AI findings, SOP audit clearances, and engineering analysis into styled PDF, DOCX (Word), or TXT formats with automatic desktop dispatch.
- **Luxury Modern UI & Custom Theme Engine:** Polished Gemini/Claude-inspired interface featuring 5 curated themes (Midnight Obsidian, Alabaster Titanium, Deep Cyber Azure, Emerald Matrix, Amethyst Royale) plus an interactive Custom Color Palette builder.

---

## 2. System Architecture

The project is structured as a decoupled client-server desktop architecture packaged into a single portable binary.

```
┌─────────────────────────────────────────────────────────────────┐
│                    DESKTOP PRESENTATION LAYER                   │
│   • PyWebView (Native Chromium/Edge Window) or Default Browser   │
│   • Single Page Application (HTML5 / CSS3 / Vanilla JavaScript)  │
│   • 5 Luxury Themes + Custom Hex Color Palette Builder          │
└────────────────────────────────┬────────────────────────────────┘
                                 │ HTTP / REST & NDJSON Streams (Port 8085)
┌────────────────────────────────▼────────────────────────────────┐
│               MULTITHREADED BACKEND SERVER (main.py)            │
│   • ThreadedTCPServer (daemon threads, low-latency dispatch)    │
│   • REST API Routing, Static Asset & UI Serving                 │
│   • Windows Job Object Supervisor (KILL_ON_JOB_CLOSE)           │
│   • Signal Handlers (SIGINT, SIGTERM, SIGBREAK) & Teardown      │
└────────────────────────────────┬────────────────────────────────┘
                                 │ Python Internal Engine API
┌────────────────────────────────▼────────────────────────────────┐
│               CORE ENGINE LAYER (sovereign_engine.py)           │
│   • Hardware Profiling & VRAM Calculation (nvidia-smi, Win32)   │
│   • Dynamic Model Catalog & Compatibility Scoring               │
│   • Ollama Process Management (Isolated subprocess, env routing) │
│   • Model Download Streamer (NDJSON chunks with cancellation)    │
│   • VRAM Memory Reliever (Model unloading & memory freeing)     │
│   • Persistent Session Storage (JSON database)                  │
│   • Universal File Parser (DOCX, XLSX, PDF, Code, Images)       │
│   • Document Generation Engine (ReportLab PDF & python-docx)    │
└────────────────────────────────┬────────────────────────────────┘
                                 │ Loopback HTTP (Port 11434)
┌────────────────────────────────▼────────────────────────────────┐
│               LOCAL INFERENCE ENGINE (Embedded Ollama)          │
│   • Standalone Ollama Server Process                            │
│   • Local Storage: data/models/ (OLLAMA_MODELS)                │
│   • Local Binaries: data/ollama/                                │
│   • Model Execution on GPU (CUDA) / CPU (AVX2/AVX-512)          │
└─────────────────────────────────────────────────────────────────┘
```

---

## 3. Directory & Storage Structure

The application operates in a completely self-contained portable directory structure:

```
busy-lovelace/
├── SovereignAIWorkbench.exe          # Compiled single-file portable Windows executable
├── app_ui.html                       # Frontend SPA (HTML/CSS/JS, 5 themes, model hub)
├── main.py                           # Multithreaded HTTP server & desktop launcher
├── sovereign_engine.py               # Core inference, hardware detection, & engine logic
├── build_release_zip.py              # Packaging & release zip distribution builder
├── build_app_and_installer.py        # Alternative installer compilation script
├── setup_installer.py                # Standalone setup wizard script
├── Run_Sovereign_AI_Workbench.bat    # Windows CMD launcher
├── Install_Sovereign_AI_Workbench.bat# Windows Installer wizard launcher
├── run_linux.sh                      # Linux shell launcher
├── icon.ico / icon.png / favicon.ico # Application branding and icons
├── deliverables/                     # Output folder for generated inspection notes
│   └── Hydrocracker_Inspection_Approval_Note.docx
└── data/                             # Portable Data Directory (Persisted across runs)
    ├── settings/
    │   ├── server_config.json        # Backend connection profile & active model
    │   └── models_catalog.json       # User-extensible model specifications
    ├── models/                       # Ollama model blobs & manifests
    ├── sessions/
    │   └── sessions.json             # Persistent conversation histories
    ├── ollama/                       # Standalone embedded Ollama binaries
    ├── exports/                      # Exported PDF, DOCX, and TXT analysis reports
    └── logs/
        └── ollama.log                # Ollama daemon stdout/stderr log
```

---

## 4. Key Python Modules & Components

### 4.1 `sovereign_engine.py` (Core Engine)
The central intelligence and execution library for the workbench (~1650 lines).

#### Key Responsibilities & Functions:
- **Hardware Detection (`profile_hardware`):**
  - CPU Cores: `os.cpu_count()`.
  - System RAM: Win32 `GlobalMemoryStatusEx` struct via `ctypes` (`ullTotalPhys`, `ullAvailPhys`) or Linux `/proc/meminfo`.
  - Storage: `shutil.disk_usage(DATA_DIR)`.
  - GPU & VRAM: Runs hidden `nvidia-smi --query-gpu=gpu_name,memory.total,memory.free --format=csv,noheader,nounits` (3-second timeout).
- **Recommendation & Fit Scoring (`get_available_models_catalog`):**
  - Formula: $\text{Required Memory (GB)} \approx (\text{Parameters (B)} \times 0.65) + 0.8$
  - Assigns dynamic badges based on hardware capability:
    - `⭐ Recommended (Best Fit)` (Optimal score fitting inside available VRAM/RAM)
    - `🟢 Full GPU Acceleration` (VRAM $\ge$ Estimated Memory)
    - `⚡ Hybrid GPU + RAM` (VRAM + RAM $\ge$ Estimated Memory)
    - `🟡 CPU Mode (Runs on RAM)` (RAM $\ge$ Estimated Memory)
    - `⚠️ High RAM Needed` (Exceeds current system memory)
- **Ollama Lifecycle Management:**
  - `find_existing_ollama()`: Detects local running instance, bundled binary in `data/ollama/`, or system PATH.
  - `ensure_ollama_installed()`: Automatically downloads standalone Ollama archive from GitHub releases in 2MB chunks if missing.
  - `start_ollama_server()`: Spawns `ollama serve` with `OLLAMA_MODELS=data/models` using `CREATE_NO_WINDOW (0x08000000)` and routes logs to `data/logs/ollama.log`.
  - `stop_ollama_server()`: Gracefully unloads all active models, terminates the child process, and kills any remaining zombie processes.
- **Process Supervision & VRAM Safety:**
  - `create_job_object()`: Initializes a Windows Job Object configured with `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE (0x2000)`.
  - `assign_process_to_job(process)`: Attaches child process handles to the job object so OS forces child cleanup if the app exits or crashes.
  - `cleanup_orphan_ollama()`: Executes `taskkill /F /IM ollama.exe` (Windows) or `pkill -f 'ollama serve'` (Linux) on startup.
  - `unload_model(model_name)`: Posts `keep_alive: 0` to Ollama `/api/generate` to immediately free VRAM.
  - `switch_model(new_model)`: Automatically triggers `unload_model` on the old model when changing selections.
- **Model Pull Streaming & Cancellation:**
  - `stream_pull_model(model_name)`: Generator reading 8KB chunks from Ollama `/api/pull` yielding NDJSON lines.
  - `cancel_model_pull(model_name)`: Triggers a `threading.Event` to immediately halt download streams.
- **Universal File Parser (`parse_any_file`):**
  - Extracts text from `.docx` (python-docx), `.xlsx`/`.xls` (openpyxl), `.pdf`, `.py`, `.js`, `.ts`, `.html`, `.cpp`, `.c`, `.java`, `.json`, `.csv`, `.txt`, `.log`, `.md`, `.yaml`, `.sql`, and image metadata.
- **Inference & Multi-Turn Chat (`query_model_chat`):**
  - Injects MRPL industrial system prompt.
  - Merges file context directly into user turns.
  - Supports local/remote Ollama (`/api/chat`) and OpenAI-compatible endpoints (`/v1/chat/completions`).
- **Document Export (`export_response_document`):**
  - PDF: Generates styled ReportLab documents with custom header banners, metadata, and markdown translation (bold, italics, monospace code blocks).
  - DOCX: Generates structured Microsoft Word documents.

---

### 4.2 `main.py` (Desktop Launcher & Multithreaded Server)
Entrypoint that runs the backend HTTP server and spawns the desktop interface.

#### Key Endpoints:
- `GET /` & `/index.html`: Serves `app_ui.html`.
- `GET /api/hardware`: Returns hardware metrics and the best-fit recommended model.
- `GET /api/available_models`: Returns full 32-model catalog enriched with live compatibility badges.
- `GET /api/local_models`: Lists installed models (checks both Ollama API tags and physical disk manifests).
- `GET /api/ollama_status`: Real-time setup and runtime readiness status.
- `GET /api/pull_model?model=...`: Streaming NDJSON progress endpoint for model downloads.
- `GET /api/sessions`: Returns saved session history.
- `GET /api/download?file=...`: Streams exported PDF/DOCX/TXT files.
- `POST /api/chat` / `/api/execute`: Multi-turn conversational chat inference.
- `POST /api/export_doc`: Generates export files and returns JSON download URL.
- `POST /api/cancel_pull`: Cancels an active model download.
- `POST /api/unload_model`: Unloads active model from GPU memory.
- `POST /api/switch_model`: Unloads previous model and switches context.
- `POST /api/delete_model`: Deletes a downloaded model from local storage.
- `POST /api/server_config`: Saves backend server type, URL, and API keys.
- `POST /api/test_server`: Validates connection to local, LAN, or remote OpenAI backends.
- `DELETE /api/sessions?id=...`: Deletes a session from persistent storage.

---

### 4.3 `app_ui.html` (Single-Page Desktop Application)
A responsive, self-contained single-page application built with vanilla HTML5, CSS3, and JavaScript (~1700 lines).

#### UI Highlights:
- **Clean Luxury AI Interface:** Minimalist header with model selector, hardware badge, engine status pill, and session drawer.
- **5 Theme Presets + Custom Palette:**
  1. 🌙 **Midnight Obsidian** (Dark background with warm amber accents)
  2. ☀️ **Alabaster Titanium** (Clean light mode with titanium grays and amber highlights)
  3. 💻 **Deep Cyber Azure** (Dark slate with GitHub/VSCode cyan-blue accents)
  4. ⚡ **Emerald Matrix** (Dark charcoal with emerald green highlights)
  5. 👑 **Amethyst Royale** (Deep obsidian with royal purple/violet accents)
  - *Custom Palette Builder:* Live interactive color pickers for Canvas, Surface, Primary Accent, and Text colors.
- **Full-Featured Model Hub:**
  - Search filter + category pills (All, Coding, Reasoning, Fast, General, Vision, Large).
  - One-click downloads with live byte/percentage progress bars and Cancel button.
  - Delete model confirmation modal.
- **Multi-File Context Attachment Bar:** Drag-and-drop or select files (DOCX, Excel, PDF, Code, Logs), preview attachment chip, and clear before sending.
- **Response Controls:** Real-time generation stop button (`AbortController`), copy to clipboard, and instant document export (PDF, Word, Text).

---

## 5. Curated Model Catalog (32 Models)

| Category | Model Name | Parameters | Size (GB) | Est. RAM/VRAM | Primary Use Case |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Coding & Automation** | `qwen2.5-coder:1.5b` | 1.5B | 1.0 GB | ~1.8 GB | Ultra-fast code completion & script fixes |
| | `starcoder2:3b` | 3B | 1.7 GB | ~2.8 GB | Multi-language code generation (600+ langs) |
| | `starcoder2:7b` | 7B | 4.3 GB | ~5.4 GB | C++, Python, Rust & PLC ladder logic |
| | `qwen2.5-coder:7b` | 7B | 4.7 GB | ~5.4 GB | Python telemetry, automation, control logic |
| | `qwen2.5-coder:14b` | 14B | 9.0 GB | ~9.9 GB | Full-stack engineering & architecture |
| **Reasoning & Math** | `deepseek-r1:1.5b` | 1.5B | 1.1 GB | ~1.8 GB | Rapid chain-of-thought logic validation |
| | `deepseek-r1:7b` | 7B | 4.7 GB | ~5.4 GB | Thermodynamic audits & root-cause analysis |
| | `deepseek-r1:8b` | 8B | 4.9 GB | ~6.0 GB | Llama-3.1 distilled mathematical reasoning |
| | `deepseek-r1:14b` | 14B | 9.0 GB | ~9.9 GB | Complex engineering physics & calculations |
| **Fast & Lightweight** | `smollm2:135m` | 0.135B | 0.1 GB | ~0.9 GB | Sub-100MB micro model for edge devices |
| | `smollm2:360m` | 0.36B | 0.3 GB | ~1.0 GB | Instant response CPU model |
| | `qwen2.5:0.5b` | 0.5B | 0.4 GB | ~1.1 GB | Low-spec hardware & rapid summarization |
| | `smollm2:1.7b` | 1.7B | 1.0 GB | ~1.9 GB | Best-in-class on-device benchmark leader |
| | `qwen2.5:1.5b` | 1.5B | 1.0 GB | ~1.8 GB | Compact multilingual model (GT 710 compatible) |
| | `llama3.2:1b` | 1B | 1.3 GB | ~1.5 GB | Meta on-device conversational assistant |
| | `gemma2:2b` | 2B | 1.6 GB | ~2.1 GB | Google high-efficiency compact model |
| | `phi3.5:3.8b` | 3.8B | 2.2 GB | ~3.3 GB | Microsoft technical reasoning on low RAM |
| **General & SOP** | `granite3-dense:2b` | 2B | 1.5 GB | ~2.1 GB | IBM enterprise compliance & business data |
| | `qwen2.5:3b` | 3B | 1.9 GB | ~2.8 GB | Technical multilingual general assistant |
| | `llama3.2:3b` | 3B | 2.0 GB | ~2.8 GB | Meta balanced speed & intelligence |
| | `mistral:7b` | 7B | 4.1 GB | ~5.4 GB | SOP clearance notes & regulatory drafting |
| | `qwen2.5:7b` | 7B | 4.7 GB | ~5.4 GB | Comprehensive science, engineering & law |
| | `llama3.1:8b` | 8B | 4.7 GB | ~6.0 GB | Meta open flagship with 128k context support |
| | `granite3-dense:8b` | 8B | 4.9 GB | ~6.0 GB | IBM enterprise risk assessment & governance |
| | `gemma2:9b` | 9B | 5.4 GB | ~6.7 GB | Google high-precision data analytics model |
| **Vision & Multimodal** | `moondream:1.8b` | 1.8B | 1.0 GB | ~2.0 GB | Lightweight image & gauge diagram inspector |
| | `llava:7b` | 7B | 4.5 GB | ~5.4 GB | Visual inspection photos & equipment gauges |
| | `llama3.2-vision:11b`| 11B | 7.9 GB | ~8.0 GB | Scanned drawings, P&ID schematics & blueprints |
| **Large & Enterprise** | `phi4:14b` | 14B | 9.1 GB | ~9.9 GB | Microsoft synthetic training for physics/math |
| | `qwen2.5:14b` | 14B | 9.0 GB | ~9.9 GB | In-depth technical synthesis & research |
| | `codestral:22b` | 22B | 13.0 GB| ~15.1 GB| Mistral 80+ language coding with 32k context |
| | `llama3.3:70b` | 70B | 42.0 GB| ~46.3 GB| Top-tier workstation intelligence |

---

## 6. Build & Packaging Guide

### 6.1 Portable Executable & Zip Distribution
The project includes an automated build script: [build_release_zip.py](file:///c:/Users/Stech/Documents/antigravity/busy-lovelace/build_release_zip.py).

To compile the application:
```powershell
python build_release_zip.py
```

### 6.2 Build Process Steps:
1. **PyInstaller Compilation:**
   - Bundles `main.py` into a single standalone binary: `dist/SovereignAIWorkbench.exe`.
   - Embeds assets: `app_ui.html`, `sovereign_engine.py`, `icon.ico`, and `icon.png`.
   - Embeds multi-resolution icon (16px to 256px) for crisp desktop and taskbar display.
   - Strips heavy non-essential dependencies (`torch`, `scipy`, `numpy`, `matplotlib`, `cv2`, `tkinter`) to keep executable under 22 MB.
2. **Distribution Assembly:**
   - Creates `dist/SovereignAIWorkbench_Portable/` with pre-made `data/` directories (`models/`, `sessions/`, `settings/`, `ollama/`, `exports/`).
   - Copies clean `README.txt`.
3. **Zip Archive:**
   - Packages entire portable suite into `SovereignAIWorkbench_v2.0_Portable.zip` (~21 MB).

---

## 7. Operational Instructions

### 7.1 Running the App
- **Windows:** Double-click `SovereignAIWorkbench.exe` (or run `Run_Sovereign_AI_Workbench.bat`).
- **Linux:** Run `bash run_linux.sh`.

### 7.2 First Launch Experience
1. On startup, the app creates and validates the `data/` folder structure.
2. It detects your CPU, RAM, and GPU VRAM, automatically profiling your system.
3. The background thread checks if Ollama is running. If not, it starts the embedded engine.
4. Open the **Model Hub** (top-right icon or dropdown menu) to select and download your preferred model.
5. Once downloaded, start chatting, attach files, or export official reports.

---
*Developed for Project SIH26117 | Mangalore Refinery and Petrochemicals Limited (MRPL)*
