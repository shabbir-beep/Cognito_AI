# Sovereign AI Workbench (Project SIH26117 | MRPL)
> **Privacy-First, Air-Gapped, On-Device Multimodal Agentic AI Assistant for Industrial & Enterprise Use**

---

## 🌟 Project Overview
The **Sovereign AI Workbench** is a 100% self-hosted, air-gapped agentic workstation engineered for high-security industrial facilities (e.g., Mangalore Refinery and Petrochemicals Limited - MRPL). It eliminates external cloud AI dependencies and data leakage risks by processing sensitive engineering drawings, inspection logs, SOP manuals, and automation code entirely on local hardware.

---

## 🚀 Quick Start Guide (How to Use)

### Option 1: Standalone Single-File Browser UI (Easiest)
Simply double-click `index.html` in your web browser, or open it directly:
```bash
# Double click index.html or open via browser
start index.html   # On Windows
```
**Key Features in Browser UI:**
- 🖥️ **Live Hardware Profiler**: VRAM, RAM, and GPU status gauges.
- 🎯 **Industrial Scenario Presets**: Instant load for Hydrocracker inspection notes, PLC automation code, and P&ID drawing OCR.
- 🔀 **Task Classifier & Router**: Live visual routing between `Qwen2.5-Coder:7B`, `Llama3.2-Vision:11B`, `DeepSeek-R1:7B`, and `Mistral-7B`.
- 🧠 **Agent ReAct Planning Loop**: 6-step animated progress console.
- 🛡️ **Permission Interceptor Gate Modal**: Interactive approval prompt before executing sandboxed file writes.
- 📄 **Deliverable Exporter**: Download compiled Microsoft Word (`.docx`) approval notes or `.py` code modules.
- 🔒 **Air-Gap Network Audit Stream**: Live packet monitor verifying 0 outbound WAN traffic.

---

### Option 2: Python Engine & Local Server
Run the single-file Python engine (`workbench.py`):

#### A. Launch Web Server Dashboard
```bash
python workbench.py
```
Open **`http://localhost:8080`** in any browser.

#### B. Launch Interactive Terminal CLI Agent
```bash
python workbench.py --cli
```

---

## 🛠️ Complete Setup & Offline Deployment Guide

### 1. Pre-Flight Resource Downloads (Before Air-Gapping)
Download these required components on an internet-connected machine:

#### LLM Serving Runtime
- **Ollama Offline Installer**: [https://github.com/ollama/ollama/releases](https://github.com/ollama/ollama/releases)

#### Open-Weight GGUF Models
Pull models via Ollama CLI:
```bash
ollama pull qwen2.5-coder:7b
ollama pull llama3.2-vision:11b
ollama pull deepseek-r1:7b
ollama pull mistral:7b
```

#### Offline Python Wheels
```bash
pip download -d ./wheels python-docx psutil requests torch
```

---

### 2. Air-Gap Zero-Network Verification Protocol
To prove 100% sovereign air-gapped security during presentation audits:

#### System Firewall Enforcement (Windows PowerShell Admin)
```powershell
New-NetFirewallRule -DisplayName "AirGap-Block-Outbound" -Direction Outbound -Action Block -Protocol Any
```

#### Socket Netstat Verification
```bash
netstat -an | findstr 11434
# Confirms socket listening exclusively on 127.0.0.1 (Loopback)
```

---

## 📁 Deliverables & Structure
- `index.html` — Single-file interactive HTML/JS UI prototype.
- `workbench.py` — Single-file Python engine with hardware profiler, task router, permission gate, local RAG retriever, deliverable builder, and web server.
- `./deliverables/` — Output directory for generated `.docx` approval notes and `.py` code files.
- `README.md` — Complete master guide & operational playbook.
