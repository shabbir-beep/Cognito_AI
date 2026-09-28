# Smart India Hackathon 2026 — Slide-Wise Manual PPT Editing Guide
**Problem Statement ID:** `SIH26117` | **Organization:** Mangalore Refinery and Petrochemicals Limited (MRPL)  
**Project Title:** Sovereign On-Premise Agentic AI Workbench using Open-Weight Multimodal LLMs for Confidential Industrial Work  

---

## Overview
This document provides the exact text, diagrams, flowcharts, tables, and screenshot placement slots needed to manually edit and finalize the official 6-slide Smart India Hackathon PowerPoint presentation.

---

# 📌 SLIDE 1: Title Slide

### 1. Text Content to Fill In:
* **Top Heading:** `SMART INDIA HACKATHON 2026`
* **Sub-Heading:** `SOVEREIGN ON-PREMISE AGENTIC AI WORKBENCH`
* **Bullet Points (Left Side):**
  * `• Problem Statement ID: SIH26117`
  * `• Problem Statement Title: Sovereign On-Premise Agentic AI Workbench using Open-Weight Multimodal LLMs for Confidential Industrial Work`
  * `• Organization: Mangalore Refinery and Petrochemicals Limited (MRPL)`
  * `• Theme: Smart Automation`
  * `• PS Category: Software`
  * `• Team ID: [Your Registered Team ID]`
  * `• Team Name: [Your Registered Team Name]`
  * `• Tech Stack & Architecture: Python 3.10+ (Multithreaded Server), PyWebView SPA, Embedded Ollama Daemon (Port 11434), ReportLab PDF, python-docx, openpyxl`
  * `• Deployment: Zero-Install Single-File Portable Windows Executable (SovereignAIWorkbench.exe, ~21MB) with portable data/ directory`

---

### 2. Diagram / Visual Element (Right Side):
* **Replace the generic lightbulb with a "Hardware VRAM Fit Matrix" Table/Card**:

| Model Category & Parameters | Model Size | Est. Memory Formula ($\text{Params} \times 0.65 + 0.8$) | Target Hardware Offload |
| :--- | :---: | :---: | :--- |
| **Qwen2.5 (0.5B – 1.5B)** | 0.4 – 1.0 GB | 1.1 – 1.8 GB | 🟢 **Full GPU** on 2GB VRAM (Nvidia GT 710) |
| **Phi-3.5 / Llama-3.2 (3B)** | 1.9 – 2.2 GB | 2.8 – 3.3 GB | ⚡ **Hybrid GPU (2GB) + System RAM** |
| **Qwen2.5-Coder / R1 (7B)** | 4.7 GB | 5.4 GB | 🟢 **Full GPU** on 6GB VRAM (RTX 3060) |
| **DeepSeek-R1 / Phi-4 (14B)** | 9.0 – 9.1 GB | 9.9 GB | 🟢 **Full GPU** on 12GB VRAM (RTX 4070) |
| **Llama-3.3 (70B Q4)** | 42.0 GB | 46.3 GB | 🏢 **Multi-GPU AI Server / Workstation** |

---

### 3. 📸 Screenshot Slot for Slide 1:
* **Placement:** Bottom-right corner or center-right.
* **What to capture:** A clean screenshot of the **Sovereign AI Workbench Home Screen** showing the app banner, active model selector (`Qwen2.5-Coder` or `DeepSeek-R1`), and the hardware telemetry pill (`GPU: GT 710 / RTX | 100% Air-Gapped`).

---

---

# 📌 SLIDE 2: Problem, Solution & Core Pillars

### 1. Text Content to Fill In:
* **Top Header:** `SOVEREIGN AI WORKBENCH`
* **THE PROBLEM (Top Left):**
  > *"Sensitive refinery telemetry across Hydrocracker Unit-4, FCCU, and Crude Distillation units, P&IDs, relief valve inspection notes (PSV-102A), and internal control scripts cannot leave premises. Commercial cloud assistants (ChatGPT, Claude) create severe data leakage risks, while manual SOP cross-referencing against 200+ page manuals causes multi-day clearance bottlenecks."*
  * **Pink Badge:** `Data Confidentiality Gap` $\rightarrow$ *Refinery telemetry, P&IDs & MAWP calculations cannot leave on-premise network.*
* **THE PROPOSED SOLUTION (Bottom Left):**
  > *"A 100% self-hosted, portable AI workbench executing on the organization's own workstation or GPU server. Features an embedded Ollama runtime, dynamic hardware profiling across a 32-model catalog, universal file context injection (Word, Excel, PDF), and automated executive document generation with verifiable zero external network traffic."*
  * **Blue Badge:** `Secure Air-Gapped Processing` $\rightarrow$ *Generates official Word & PDF approval notes with 0 external calls.*

---

### 2. Diagram / Visual Elements:
* **Center Visual: 6-Spoke Circular Wheel (6 Core Pillars)**:
  * **Center Circle:** `Self-Hosted AI Workbench`
  * **Spoke 1 (Top / Orange):** `Multi-Step Work Planning` $\rightarrow$ *Audits Hydrocracker MAWP ($\le 150\text{ bar}$, $140\text{ bar}$ limit) and valve calibration windows.*
  * **Spoke 2 (Top-Right / Yellow):** `Local Tool Integration` $\rightarrow$ *Parses DOCX, Excel telemetry sheets, PDFs, and logs locally with zero cloud API dependencies.*
  * **Spoke 3 (Bottom-Right / Green):** `Open-Weight Model Hub` $\rightarrow$ *32-model catalog (Coding, DeepSeek-R1 Reasoning, Vision, SOP, 70B) with 1-click streaming pull.*
  * **Spoke 4 (Bottom / Cyan):** `Scan & Drawing Reasoning` $\rightarrow$ *LLaVA-7B and Moondream-1.8B inspect scanned P&ID schematics and equipment gauge dials.*
  * **Spoke 5 (Bottom-Left / Blue):** `Verifiable Zero-External Mode` $\rightarrow$ *Hardcoded loopback `127.0.0.1`; Wireshark confirms zero outbound packets during chat/export.*
  * **Spoke 6 (Top-Left / Pink):** `VRAM & Process Safety` $\rightarrow$ *Win32 Job Object (`KILL_ON_JOB_CLOSE`) + dynamic `keep_alive: 0` VRAM memory release.*

* **Right Visual: 5-Tier Stacked Cylinder (System Capability Scope)**:
  * **Tier 5 (Top / Green):** `Multimodal Deliverables` $\rightarrow$ *Exports styled ReportLab PDF & python-docx notes with executive approval sign-offs to Desktop.*
  * **Tier 4 (Blue):** `Local Knowledge Base` $\rightarrow$ *Grounding in MRPL refinery SOPs, Hydrocracker Unit-4 telemetry, and multi-turn session persistence.*
  * **Tier 3 (Pink):** `Permission-Gated Sandbox` $\rightarrow$ *Windows Job Object (`0x2000`), hidden subprocess execution (`CREATE_NO_WINDOW`), & orphan cleanup.*
  * **Tier 2 (Orange):** `Automatic Model Routing` $\rightarrow$ *Routes prompt complexity to optimal model (Qwen2.5-Coder for code, DeepSeek-R1 for math/audit).*
  * **Tier 1 (Base / Yellow):** `Model Recommendation` $\rightarrow$ *Calculates GPU VRAM/RAM fit ($\text{Mem} = \text{Params} \times 0.65 + 0.8\text{ GB}$) to prevent OOM crashes.*

---

### 3. 📸 Screenshot Slot for Slide 2:
* **Placement:** Bottom center-left (just below the Proposed Solution box).
* **What to capture:** Screenshot of the **Multi-File Context Attachment Bar** in the UI with a file chip attached (e.g., `Hydrocracker_Telemetry_Log.xlsx` or `P&ID_Schematic.pdf`).

---

---

# 📌 SLIDE 3: Technical Approach & Inspection Pipeline

### 1. Left Diagram: "Building a Comprehensive Local System" (Layered Architecture Stack)
* **Convert into 6 horizontal connected layers converging into a central hub:**
  * **Layer 1 (Cyan):** `Client UI (PyWebView)` $\rightarrow$ *Single-file SPA (HTML5/CSS3/JS); 5 luxury themes (Midnight Obsidian, Titanium, Azure, Matrix, Amethyst) + custom palette builder; stop generation button.*
  * **Layer 2 (Green):** `Orchestration (main.py)` $\rightarrow$ *Multithreaded TCP Server on port 8085; low-latency REST API + streaming NDJSON progress endpoints; daemon thread dispatch.*
  * **Layer 3 (Blue):** `Guardrail Supervisor` $\rightarrow$ *Win32 Job Object supervisor (`0x2000`) prevents background zombies; dynamic `keep_alive: 0` unloads models from VRAM immediately upon model switch or exit.*
  * **Layer 4 (Yellow):** `Tools & Ingestion Engine` $\rightarrow$ *Universal parser (`python-docx`, `openpyxl`, PDF, code, logs) & persistent multi-turn JSON session database (`data/sessions/sessions.json`).*
  * **Layer 5 (Orange):** `Inference Layer` $\rightarrow$ *Embedded standalone Ollama daemon (`port 11434`) running GGUF quantized models on CUDA GPU / CPU AVX2, with remote LAN & OpenAI-compatible fallback.*
  * **Layer 6 (Pink):** `Air-Gap Boundary` $\rightarrow$ *Hardcoded localhost `127.0.0.1` binding; zero cloud telemetry; direct OS-level desktop document generation and automated application dispatch.*
  * **Central Circle:** `Unified System` $\rightarrow$ *Decoupled server, engine, & storage integrated in a single portable bundle.*

---

### 2. Right Diagram: "Streamlined Inspection Process" (6-Step Flowchart)
* **Convert into a 6-step vertical/chevron flowchart leading to an output deliverable box:**
  * **Step 1 (Green):** `Report & Telemetry Upload` $\rightarrow$ *Operator attaches Hydrocracker logs, Excel sensor data, or P&ID drawings via attachment bar.*
  * **Step 2 (Yellow):** `Context Data Extraction` $\rightarrow$ *Universal parser extracts operating pressures ($142.5\text{ bar}$), temps ($410^\circ\text{C}$), and valve tags (`PSV-102A`).*
  * **Step 3 (Cyan):** `Hardware-Guided Model Routing` $\rightarrow$ *Engine matches prompt complexity to optimal installed model (`DeepSeek-R1` / `Qwen2.5` / `Moondream`).*
  * **Step 4 (Purple):** `SOP Compliance Reasoner` $\rightarrow$ *LLM audits telemetry against MRPL SOP rules ($\text{MAWP} \le 150\text{ bar}$; flags $142.5\text{ bar}$ secondary sensor calibration requirement).*
  * **Step 5 (Pink):** `Guardrail Approval & Formatting` $\rightarrow$ *AI formats structured findings with executive summary, compliance status, & sign-off block.*
  * **Step 6 (Orange):** `Desktop Document Dispatch` $\rightarrow$ *Builds formatted `.docx` and styled `.pdf` documents directly on Desktop and launches in Word/Acrobat.*
  * **Final Output Box:** `Comprehensive Inspection Report` $\rightarrow$ *Official MRPL compliance note with full telemetry audit trail & signature block.*

---

### 3. Bottom Banner: Implementation Components
* **Text to Replace:**
  > **Agent Harness:** Multithreaded Python server (`main.py`) + `sovereign_engine.py` | **Inference:** Embedded Portable Ollama + GGUF Quantized Models (CUDA / CPU AVX2) | **Ingestion:** Native `python-docx`, `openpyxl`, PDF text extractors | **UI:** Single-File HTML5/CSS3 SPA in PyWebView (5 Luxury Themes + Custom Palette) | **Safety & Isolation:** Windows Job Object (`KILL_ON_JOB_CLOSE`) + dynamic `keep_alive: 0` VRAM unloader + Zero-Egress Loopback

---

### 4. 📸 Screenshot Slot for Slide 3:
* **Placement:** Right side or middle lower section.
* **What to capture:** Screenshot of the **Live Chat Interface** showing an AI response analyzing Hydrocracker telemetry logs with the **Stop Response button** and **Export to PDF / Word buttons** visible.

---

---

# 📌 SLIDE 4: Feasibility, Stepper & Risk Mitigation

### 1. Top-Left Table: Feasibility on a Single Workstation
* **Create a 2-column table:**

| Feature | Implementation & Technical Verification |
| :--- | :--- |
| **Open-Weight Models** | Curated 32-model catalog (0.135B to 70B) covering Qwen2.5-Coder, DeepSeek-R1, Llama-3.2, Gemma-2, Moondream, Mistral, and Codestral. |
| **Quantized GGUF Models** | 4-bit (`Q4_K_M`) & 8-bit (`Q8_0`) quantization allows full GPU offload on modest 2GB VRAM cards (Nvidia GT 710) up to multi-GPU servers. |
| **Embedded Runtime** | Standalone embedded Ollama daemon with isolated storage (`data/models/`); runs with zero administrative rights or external dependencies. |
| **Portable Data Storage** | All settings, sessions, models, and exports stored in portable `data/` folder packaged in a single ~21MB ZIP archive. |

---

### 2. Top-Right: "Build an Agentic System" (6-Box Stepper Flow)
* **Convert into 6 horizontal connected cards:**
  1. `1. Foundation` $\rightarrow$ *Portable EXE + multithreaded server + embedded Ollama runtime*
  2. `2. Model Layer` $\rightarrow$ *Hardware profiler (nvidia-smi) + 32-model hub streaming pull*
  3. `3. Guardrails` $\rightarrow$ *Win32 Job Object supervisor + keep_alive=0 dynamic VRAM release*
  4. `4. Agentic Core` $\rightarrow$ *Multi-turn session memory + MRPL SOP system prompt grounding*
  5. `5. Multimodal` $\rightarrow$ *DOCX, Excel telemetry, PDF, code, & P&ID diagram parsing*
  6. `6. Hardening` $\rightarrow$ *ReportLab PDF & Word export + offline zero-egress proof*

---

### 3. Bottom-Center Table: Risk and Mitigation Evidence
* **Create a 3-column table:**

| Identified Risk | Control Implemented | Live Demo Evidence |
| :--- | :--- | :--- |
| **Model Exceeds VRAM (OOM)** | Hardware profiler + formula: $\text{Mem} = \text{Params} \times 0.65 + 0.8\text{ GB}$ | Dynamic badges (`⭐ Optimal`, `🟢 Full GPU`, `🟡 CPU Mode`) prevent OOM crashes. |
| **VRAM Leak / Zombie Tasks** | Win32 Job Object (`0x2000`) + dynamic `keep_alive: 0` unloader | Verified via `nvidia-smi`: exactly 0 MB VRAM retained after exit or model switch. |
| **Cloud Data Leakage** | Hardcoded loopback (`127.0.0.1:8085` / `11434`) with zero external calls | Wireshark network monitor records 0 outbound packets during chat and export. |
| **Interrupted Downloads** | NDJSON streaming reader with thread cancellation event (`_active_pulls`) | Live byte progress bar in Model Hub with instant responsive **Cancel** button. |

---

### 4. Bottom-Right: Validation Gates (Target Rings)
* **Create 5 concentric target circles or stacked badges:**
  * **Gate 5 (Outer):** `Exportable Proof` $\rightarrow$ *Final Word & PDF approval deliverable generated on Desktop.*
  * **Gate 4:** `End-to-End Agent Task` $\rightarrow$ *Full SOP audit of Hydrocracker telemetry logs.*
  * **Gate 3:** `Approved Action` $\rightarrow$ *Gated execution & dynamic model switching.*
  * **Gate 2:** `Demonstrated Routing` $\rightarrow$ *Hardware memory formula scores 32 models live.*
  * **Gate 1 (Center):** `Basic Local Answer` $\rightarrow$ *Sub-second token generation on local CPU/GPU.*

---

### 5. 📸 Screenshot Slot for Slide 4:
* **Placement:** Beside the Risk table or above the Validation Gates.
* **What to capture:** Screenshot of the **Model Hub UI** showing 1-click model cards with parameter counts, sizes, category pills (Coding, Reasoning, Vision), and live download progress bars with the Cancel button.

---

---

# 📌 SLIDE 5: Impact, 5-Gear Workflow & Value Matrix

### 1. Left Diagram: "Workflow Execution Process" (5 Interlocking Gears Flow)
* **Create 5 connected circular gear icons / cards:**
  * **Gear 1 (Blue):** `1. Route Task Types` $\rightarrow$ *Hardware profiler routes coding to Qwen2.5-Coder & audits to DeepSeek-R1 based on VRAM fit.*
  * **Gear 2 (Teal):** `2. Act in Sandbox` $\rightarrow$ *Python automation & telemetry scripts execute in hidden, permissioned subprocesses (`CREATE_NO_WINDOW`).*
  * **Gear 3 (Cyan):** `3. Read Report Data` $\rightarrow$ *Extracts Hydrocracker telemetry ($142.5\text{ bar}$, $410^\circ\text{C}$) and P&ID valve tags (`PSV-102A`) from files.*
  * **Gear 4 (Yellow):** `4. Deliver Approval Note` $\rightarrow$ *Exports formal MRPL compliance note directly to Desktop in styled DOCX/PDF format.*
  * **Gear 5 (Purple):** `5. Prove Zero Egress` $\rightarrow$ *Wireshark monitor confirms exactly 0 outbound packets in air-gapped localhost mode.*

---

### 2. Right Top: "Possible Use Case Scenario" (Stakeholder Before vs. After)
* **Text to Replace:**
  * `• Plant Engineers: ` *Manual log cross-checking ($3\text{–}4\text{ hrs}$) $\rightarrow$ Instant local telemetry ingestion, automated SOP checks, & evidence-backed drafts ($5\text{ seconds}$).*
  * `• Safety Approvers: ` *Multi-day document review cycles $\rightarrow$ Standardized, auditable compliance notes generated in seconds with signature lines.*
  * `• IT & Cybersecurity: ` *Severe cloud AI data policy exposure $\rightarrow$ 100% on-premise air-gapped execution with verifiable zero egress traffic.*
  * `• Executive Management: ` *High recurring cloud API subscription fees $\rightarrow$ Zero recurring costs, portable USB-deployable solution across all plant units.*

---

### 3. Right Bottom: "Why It Matters" (4 Value Pillars)
* **Text to Replace:**
  * `• 100% Confidentiality: ` *Protects proprietary refinery telemetry, P&IDs, tariff calculations, and code by keeping all inference on-device.*
  * `• 10x Productivity: ` *Transforms multi-hour manual reviews into an automated flow that extracts, audits, drafts, and saves final reports.*
  * `• Total Operator Control: ` *Keeps operators in charge of consequential approvals with structured audit trails, local tools, and process supervisors.*
  * `• Universal Hardware: ` *Runs seamlessly on modest 2GB VRAM GPUs (Nvidia GT 710) up to multi-GPU enterprise AI workstations.*

---

### 4. 📸 Screenshot Slot for Slide 5:
* **Placement:** Bottom-right corner or center.
* **What to capture:** Screenshot of the **Generated Microsoft Word (`.docx`) or Styled PDF Report** (e.g. `Hydrocracker_Inspection_Approval_Note.docx`) showing the MRPL header, SOP compliance table, and signature line.

---

---

# 📌 SLIDE 6: Research, References & Implementation Roadmap

### 1. Left Section: "Demonstration Checklist" (5 Tested Gates)
* **Create 5 cards with green checkmark icons (`✓`):**
  1. `[✓] Hardware Recommendation Matches Demo PC` $\rightarrow$ *Detects CPU cores, RAM, and GPU VRAM via nvidia-smi to recommend optimal model.*
  2. `[✓] Model Routing Switches Installed Models` $\rightarrow$ *Dynamically routes tasks across 32-model catalog and unloads VRAM via `keep_alive: 0`.*
  3. `[✓] Agent Reads Scan, Consults SOP, Exports Note` $\rightarrow$ *Ingests Hydrocracker telemetry, audits MAWP limits, and writes Word note to Desktop.*
  4. `[✓] Sandboxed Subprocess Runs Safely` $\rightarrow$ *Windows Job Object (`0x2000`) supervises subprocesses with `CREATE_NO_WINDOW`.*
  5. `[✓] Strict-Mode Monitor Proves 0 Outbound Calls` $\rightarrow$ *Wireshark and network telemetry confirm 100% air-gapped localhost loopback operation.*

---

### 2. Right Section: "Local Workbench Implementation Strategy" ($2 \times 5$ Grid Roadmap)
* **Create a $2 \times 5$ grid of connected stage cards:**
  * **Row 1:**
    1. `1. Scope (SIH26117)` $\rightarrow$ *Air-gapped, multimodal, multi-model AI workbench for MRPL.*
    2. `2. Architecture` $\rightarrow$ *Portable single-file EXE with decoupled TCP server & SPA.*
    3. `3. Embedded Engine` $\rightarrow$ *Bundled Ollama daemon with isolated `OLLAMA_MODELS` storage.*
    4. `4. File Ingestion` $\rightarrow$ *Native python-docx, openpyxl, PDF & diagram text extractors.*
    5. `5. Strict Air-Gap` $\rightarrow$ *Hardcoded loopback `127.0.0.1` with zero outbound network calls.*
  * **Row 2:**
    6. `6. Hardware Scoring` $\rightarrow$ *Memory formula scoring with dynamic visual fit badges.*
    7. `7. 32-Model Hub` $\rightarrow$ *Categorized catalog with 1-click download & live progress.*
    8. `8. VRAM Guard` $\rightarrow$ *Win32 Job Object supervisor & dynamic `keep_alive: 0` unloader.*
    9. `9. Document Engine` $\rightarrow$ *ReportLab PDF formatting & direct desktop Word dispatch.*
    10. `10. Luxury SPA UI` $\rightarrow$ *5 preset themes, custom palette builder, & stop response button.*

---

### 3. Bottom Footer URLs:
* **Text to Replace:**
  > `URLs: sih.gov.in/sih2026PS (SIH26117) | github.com/ollama/ollama | github.com/reportlab/reportlab | github.com/python-openxml/python-docx | github.com/qwenlm/Qwen2.5 | github.com/deepseek-ai/DeepSeek-R1`

---

### 4. 📸 Screenshot Slot for Slide 6:
* **Placement:** Bottom left corner (below the checklist).
* **What to capture:** Screenshot of the **Theme Customizer / Color Palette Modal** in the UI showing the 5 luxury themes (Midnight Obsidian, Titanium, Azure, Matrix, Amethyst) and custom color pickers.
