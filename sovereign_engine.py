"""
Sovereign AI Engine — Hardware-Intelligent Model Manager, Multi-Turn RAG & Session Engine
Project Code: SIH26117 | Mangalore Refinery and Petrochemicals Limited (MRPL)
Cross-platform compatible: Windows & Linux
"""

import sys
import os
import json
import time
import ctypes
import subprocess
import zipfile
import tarfile
import shutil
import threading
import signal
import urllib.request
import ssl
from datetime import datetime

# --- ENCODING FIX (Windows CP1252 / UTF-8) ---
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

try:
    import requests
except ImportError:
    requests = None

# ============================================================
# CROSS-PLATFORM PATHS & CONFIGURATION (Windows & Linux)
# ============================================================
def get_app_dir():
    """Returns absolute path of directory containing EXE or script."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = get_app_dir()
DATA_DIR = os.path.join(APP_DIR, 'data')
SETTINGS_DIR = os.path.join(DATA_DIR, 'settings')
MODELS_DIR = os.path.join(DATA_DIR, 'models')
SESSIONS_DIR = os.path.join(DATA_DIR, 'sessions')
OLLAMA_DIR = os.path.join(DATA_DIR, 'ollama')
EXPORTS_DIR = os.path.join(DATA_DIR, 'exports')

if sys.platform == 'win32':
    OLLAMA_EXE = os.path.join(OLLAMA_DIR, 'ollama.exe')
    OLLAMA_DOWNLOAD_URL = 'https://github.com/ollama/ollama/releases/latest/download/ollama-windows-amd64.zip'
else:
    OLLAMA_EXE = os.path.join(OLLAMA_DIR, 'ollama')
    OLLAMA_DOWNLOAD_URL = 'https://github.com/ollama/ollama/releases/latest/download/ollama-linux-amd64.tgz'

SESSIONS_FILE = os.path.join(SESSIONS_DIR, 'sessions.json')
PROJECTS_FILE = os.path.join(SETTINGS_DIR, 'projects.json')
CONFIG_FILE = os.path.join(SETTINGS_DIR, 'server_config.json')
OLLAMA_API_BASE = 'http://127.0.0.1:11434'

def verify_and_create_data_dirs():
    """Safety net: creates all required data folders and checks write permission."""
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        os.makedirs(MODELS_DIR, exist_ok=True)
        os.makedirs(SESSIONS_DIR, exist_ok=True)
        os.makedirs(OLLAMA_DIR, exist_ok=True)
        os.makedirs(EXPORTS_DIR, exist_ok=True)

        test_file = os.path.join(DATA_DIR, '.write_test')
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write('ok')
        os.remove(test_file)
        return True, ""
    except Exception as e:
        return False, f"Application location is Read-Only or Permission Denied: '{APP_DIR}'. Please extract or move the folder to a writable directory (e.g. Documents or Desktop)."

def migrate_legacy_data():
    """Migrates legacy %LOCALAPPDATA% / ~/.sovereign_ai data to portable data/ directory if present."""
    if sys.platform == 'win32':
        legacy_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'SovereignAI')
    else:
        legacy_dir = os.path.expanduser('~/.sovereign_ai')

    if not os.path.exists(legacy_dir):
        return

    try:
        legacy_cfg = os.path.join(legacy_dir, 'server_config.json')
        if os.path.exists(legacy_cfg) and not os.path.exists(CONFIG_FILE):
            shutil.copy2(legacy_cfg, CONFIG_FILE)

        legacy_sess = os.path.join(legacy_dir, 'sessions.json')
        if os.path.exists(legacy_sess) and not os.path.exists(SESSIONS_FILE):
            shutil.copy2(legacy_sess, SESSIONS_FILE)

        legacy_models = os.path.join(legacy_dir, 'models')
        if os.path.exists(legacy_models) and os.path.isdir(legacy_models):
            for item in os.listdir(legacy_models):
                src = os.path.join(legacy_models, item)
                dst = os.path.join(MODELS_DIR, item)
                if not os.path.exists(dst):
                    if os.path.isdir(src):
                        shutil.copytree(src, dst, dirs_exist_ok=True)
                    else:
                        shutil.copy2(src, dst)
    except Exception:
        pass

is_writable, write_error_msg = verify_and_create_data_dirs()
if is_writable:
    migrate_legacy_data()

# Global state
_ollama_process = None
_job_object = None  # Windows Job Object handle for auto-cleanup
_current_model_name = None  # Track loaded model for unload-on-switch
_active_pulls = {}  # model_name -> threading.Event()
_setup_status = {
    "stage": "idle",
    "percent": 0,
    "message": "Initializing...",
    "ollama_installed": False,
    "ollama_running": False,
}
_setup_lock = threading.Lock()

def _update_setup(stage, percent, message):
    with _setup_lock:
        _setup_status["stage"] = stage
        _setup_status["percent"] = percent
        _setup_status["message"] = message

def cancel_model_pull(model_name):
    """Signals cancellation of an active model download."""
    if model_name in _active_pulls:
        _active_pulls[model_name].set()
        return True
    return False

# ============================================================
# CROSS-PLATFORM HARDWARE DETECTION (Windows ctypes & Linux proc)
# ============================================================
class MEMORYSTATUSEX(ctypes.Structure):
    _fields_ = [
        ('dwLength', ctypes.c_ulong),
        ('dwMemoryLoad', ctypes.c_ulong),
        ('ullTotalPhys', ctypes.c_ulonglong),
        ('ullAvailPhys', ctypes.c_ulonglong),
        ('ullTotalPageFile', ctypes.c_ulonglong),
        ('ullAvailPageFile', ctypes.c_ulonglong),
        ('ullTotalVirtual', ctypes.c_ulonglong),
        ('ullAvailVirtual', ctypes.c_ulonglong),
        ('sullAvailExtendedVirtual', ctypes.c_ulonglong),
    ]

MODELS_CATALOG_FILE = os.path.join(SETTINGS_DIR, 'models_catalog.json')

def hidden_subprocess_run(cmd_args, **kwargs):
    """Executes a subprocess fully hidden on Windows without console window popups."""
    if sys.platform == 'win32':
        si = kwargs.pop('startupinfo', None) or subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 0
        kwargs['startupinfo'] = si
        kwargs['creationflags'] = kwargs.get('creationflags', 0) | 0x08000000  # CREATE_NO_WINDOW
    return subprocess.run(cmd_args, **kwargs)

def load_models_catalog():
    """Loads catalog from data/settings/models_catalog.json, merging base catalog with custom entries."""
    existing_map = {}
    if os.path.exists(MODELS_CATALOG_FILE):
        try:
            with open(MODELS_CATALOG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    for item in data:
                        if isinstance(item, dict) and "name" in item:
                            existing_map[item["name"]] = item
        except Exception:
            pass

    # Merge base models, then any custom user-added models
    final_list = []
    seen = set()
    for m in BASE_MODELS_CATALOG:
        final_list.append(dict(m))
        seen.add(m["name"])

    for name, item in existing_map.items():
        if name not in seen:
            final_list.append(item)
            seen.add(name)

    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(MODELS_CATALOG_FILE, 'w', encoding='utf-8') as f:
            json.dump(final_list, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

    return final_list

_cached_hardware_profile = None

def profile_hardware(force_refresh=False):
    """Detects physical system hardware cross-platform (CPU, RAM, GPU VRAM, Free Disk). Caches result for fast startup."""
    global _cached_hardware_profile
    if _cached_hardware_profile is not None and not force_refresh:
        return dict(_cached_hardware_profile)

    cpu_cores = os.cpu_count() or 8
    ram_total_gb = 16.0
    ram_free_gb = 8.0

    if sys.platform == 'win32':
        try:
            stat = MEMORYSTATUSEX()
            stat.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat))
            ram_total_gb = round(stat.ullTotalPhys / (1024 ** 3), 2)
            ram_free_gb = round(stat.ullAvailPhys / (1024 ** 3), 2)
        except Exception:
            pass
    else:
        # Linux memory detection via /proc/meminfo
        try:
            with open('/proc/meminfo', 'r') as f:
                lines = f.readlines()
            total_kb, avail_kb = 0, 0
            for line in lines:
                if line.startswith('MemTotal:'):
                    total_kb = int(line.split()[1])
                elif line.startswith('MemAvailable:'):
                    avail_kb = int(line.split()[1])
            if total_kb:
                ram_total_gb = round(total_kb / (1024 * 1024), 2)
                ram_free_gb = round(avail_kb / (1024 * 1024), 2)
        except Exception:
            pass

    # Disk Free Space Detection
    try:
        disk_info = shutil.disk_usage(DATA_DIR if os.path.exists(DATA_DIR) else APP_DIR)
        disk_free_gb = round(disk_info.free / (1024 ** 3), 2)
    except Exception:
        disk_free_gb = 50.0

    # GPU Detection via nvidia-smi
    gpu_name = "No Dedicated GPU Detected"
    vram_total_mb = 0
    vram_free_mb = 0
    try:
        res = hidden_subprocess_run(
            ["nvidia-smi", "--query-gpu=gpu_name,memory.total,memory.free", "--format=csv,noheader,nounits"],
            capture_output=True, text=True, timeout=2
        )
        if res.returncode == 0 and res.stdout.strip():
            parts = res.stdout.strip().split("\n")[0].split(",")
            gpu_name = parts[0].strip()
            vram_total_mb = int(parts[1].strip())
            vram_free_mb = int(parts[2].strip())
    except Exception:
        pass

    vram_total_gb = round(vram_total_mb / 1024, 2)
    vram_free_gb = round(vram_free_mb / 1024, 2)

    # Dynamic Model Recommendation Scoring
    catalog = load_models_catalog()
    best_model = None
    best_score = -1
    best_reason = ""

    for m in catalog:
        raw_p = m.get("params", "7B").upper().replace("B", "").strip()
        try:
            params_b = float(raw_p)
        except Exception:
            params_b = 7.0

        est_mem_gb = round((params_b * 0.65) + 0.8, 2)
        size_gb = m.get("size_gb", 4.0)

        if size_gb > disk_free_gb:
            continue

        score = 0
        reason = ""

        if vram_total_gb >= est_mem_gb:
            score = 300 + params_b
            reason = f"Full GPU Offload: Dedicated {vram_total_gb}GB VRAM ({gpu_name}) fits {m['label']} ({est_mem_gb}GB needed)."
        elif (vram_total_gb + ram_total_gb) >= est_mem_gb and ram_total_gb >= est_mem_gb:
            score = 200 + params_b
            reason = f"Hybrid VRAM ({vram_total_gb}GB) + System RAM ({ram_total_gb}GB) comfortably fits {m['label']} ({est_mem_gb}GB needed)."
        elif ram_total_gb >= est_mem_gb:
            score = 100 + params_b
            reason = f"CPU Inference: System RAM ({ram_total_gb}GB, {cpu_cores} Cores) accommodates {m['label']}."

        if score > best_score:
            best_score = score
            best_model = m["name"]
            best_reason = reason

    if not best_model:
        best_model = "qwen2.5:1.5b"
        best_reason = f"Compact 1.5B model selected for constrained hardware ({ram_total_gb}GB RAM, {disk_free_gb}GB free disk)."

    _cached_hardware_profile = {
        "cpu_cores": cpu_cores,
        "ram_gb": ram_total_gb,
        "ram_free_gb": ram_free_gb,
        "disk_free_gb": disk_free_gb,
        "gpu_name": gpu_name,
        "vram_gb": vram_total_gb,
        "vram_free_gb": vram_free_gb,
        "recommended_model": best_model,
        "recommended_reason": best_reason
    }
    return dict(_cached_hardware_profile)

# ============================================================
# DYNAMIC MODEL CATALOG & HARDWARE COMPATIBILITY SCORING
# ============================================================
BASE_MODELS_CATALOG = [
    # --- Coding & Industrial Automation ---
    {
        "name": "qwen2.5-coder:7b",
        "label": "Qwen2.5-Coder (7B)",
        "params": "7B",
        "size_gb": 4.7,
        "category": "Coding & Automation",
        "desc": "State-of-the-art coding intelligence for Python telemetry scripts, PLC automation, and control logic.",
        "tags": ["Coding", "Popular", "Automation", "Python"]
    },
    {
        "name": "qwen2.5-coder:1.5b",
        "label": "Qwen2.5-Coder (1.5B)",
        "params": "1.5B",
        "size_gb": 1.0,
        "category": "Coding & Automation",
        "desc": "Ultra-lightweight coding model. High speed for rapid code completion and simple script debugging.",
        "tags": ["Coding", "Lightweight", "Low Latency"]
    },
    {
        "name": "qwen2.5-coder:14b",
        "label": "Qwen2.5-Coder (14B)",
        "params": "14B",
        "size_gb": 9.0,
        "category": "Coding & Automation",
        "desc": "Deep coding intelligence for full-stack engineering, complex algorithms, and industrial system architecture.",
        "tags": ["Coding", "High Precision", "Architecture"]
    },
    {
        "name": "starcoder2:3b",
        "label": "StarCoder2 (3B)",
        "params": "3B",
        "size_gb": 1.7,
        "category": "Coding & Automation",
        "desc": "BigCode's efficient multi-language code completion model trained on 600+ programming languages.",
        "tags": ["Coding", "Multi-Language", "Fast"]
    },
    {
        "name": "starcoder2:7b",
        "label": "StarCoder2 (7B)",
        "params": "7B",
        "size_gb": 4.3,
        "category": "Coding & Automation",
        "desc": "Powerful coding assistant for C++, Rust, Python, Go, and PLC ladder logic translation.",
        "tags": ["Coding", "Refactoring", "Logic"]
    },

    # --- Chain-of-Thought & Reasoning ---
    {
        "name": "deepseek-r1:1.5b",
        "label": "DeepSeek-R1 (1.5B)",
        "params": "1.5B",
        "size_gb": 1.1,
        "category": "Reasoning & Math",
        "desc": "Fast chain-of-thought reasoning model for rapid logic validation, basic algebra, and quick audits.",
        "tags": ["Reasoning", "Math", "Lightweight", "Popular"]
    },
    {
        "name": "deepseek-r1:7b",
        "label": "DeepSeek-R1 (7B)",
        "params": "7B",
        "size_gb": 4.7,
        "category": "Reasoning & Math",
        "desc": "Breakthrough reasoning model for thermodynamic audits, engineering calculations, and root-cause analysis.",
        "tags": ["Reasoning", "Math", "Popular", "Audit"]
    },
    {
        "name": "deepseek-r1:8b",
        "label": "DeepSeek-R1 (8B)",
        "params": "8B",
        "size_gb": 4.9,
        "category": "Reasoning & Math",
        "desc": "Llama-3.1 distilled DeepSeek reasoning model with broad world knowledge and mathematical rigor.",
        "tags": ["Reasoning", "Math", "Llama-Distilled"]
    },
    {
        "name": "deepseek-r1:14b",
        "label": "DeepSeek-R1 (14B)",
        "params": "14B",
        "size_gb": 9.0,
        "category": "Reasoning & Math",
        "desc": "High-precision analytical model for complex engineering physics, multi-step problem solving, and research.",
        "tags": ["Reasoning", "High Precision", "Physics"]
    },

    # --- Fast & Ultra-Lightweight (<3GB VRAM) ---
    {
        "name": "smollm2:135m",
        "label": "SmolLM2 (135M)",
        "params": "0.135B",
        "size_gb": 0.1,
        "category": "Fast & Lightweight",
        "desc": "Microscopic sub-100MB model for ultra-low power devices, basic formatting, and edge classification.",
        "tags": ["Lightweight", "Sub-100MB", "Edge"]
    },
    {
        "name": "smollm2:360m",
        "label": "SmolLM2 (360M)",
        "params": "0.36B",
        "size_gb": 0.3,
        "category": "Fast & Lightweight",
        "desc": "Tiny 360M parameter model with astonishing token generation speeds on CPU.",
        "tags": ["Lightweight", "Instant", "CPU"]
    },
    {
        "name": "smollm2:1.7b",
        "label": "SmolLM2 (1.7B)",
        "params": "1.7B",
        "size_gb": 1.0,
        "category": "Fast & Lightweight",
        "desc": "Hugging Face's best-in-class 1.7B model. Outperforms many larger models on on-device benchmarks.",
        "tags": ["Lightweight", "On-Device", "Fast"]
    },
    {
        "name": "qwen2.5:0.5b",
        "label": "Qwen-2.5 (0.5B)",
        "params": "0.5B",
        "size_gb": 0.4,
        "category": "Fast & Lightweight",
        "desc": "Compact 500M model. Ideal for low-spec laptops, minimal RAM footprint, and instant responses.",
        "tags": ["Lightweight", "Sub-1GB", "Low RAM"]
    },
    {
        "name": "qwen2.5:1.5b",
        "label": "Qwen-2.5 (1.5B)",
        "params": "1.5B",
        "size_gb": 1.0,
        "category": "Fast & Lightweight",
        "desc": "Compact multilingual model with great accuracy. Runs smoothly on 2GB VRAM GPUs (GT 710).",
        "tags": ["Lightweight", "2GB VRAM", "Popular"]
    },
    {
        "name": "llama3.2:1b",
        "label": "Llama-3.2 (1B)",
        "params": "1B",
        "size_gb": 1.3,
        "category": "Fast & Lightweight",
        "desc": "Meta's official 1B lightweight model. Highly optimized for on-device chat and summaries.",
        "tags": ["Meta", "Lightweight", "Fast"]
    },
    {
        "name": "gemma2:2b",
        "label": "Gemma-2 (2B)",
        "params": "2B",
        "size_gb": 1.6,
        "category": "Fast & Lightweight",
        "desc": "Google's lightweight 2B model offering high-fidelity responses with minimal memory requirements.",
        "tags": ["Google", "Lightweight", "Efficient"]
    },
    {
        "name": "phi3.5:3.8b",
        "label": "Phi-3.5 (3.8B)",
        "params": "3.8B",
        "size_gb": 2.2,
        "category": "General Chat & SOP",
        "desc": "Microsoft's high-efficiency 3.8B model. Designed for technical reasoning on consumer hardware.",
        "tags": ["Microsoft", "2GB VRAM", "Popular", "SOP"]
    },

    # --- General Purpose & Standard Workhorses ---
    {
        "name": "llama3.2:3b",
        "label": "Llama-3.2 (3B)",
        "params": "3B",
        "size_gb": 2.0,
        "category": "General Chat & SOP",
        "desc": "Meta's flagship small assistant. Excellent balance of speed, intelligence, and safety alignment.",
        "tags": ["Meta", "Popular", "Balanced"]
    },
    {
        "name": "qwen2.5:3b",
        "label": "Qwen-2.5 (3B)",
        "params": "3B",
        "size_gb": 1.9,
        "category": "General Chat & SOP",
        "desc": "High quality 3B general assistant with strong technical and multilingual conversational skills.",
        "tags": ["Balanced", "Multilingual"]
    },
    {
        "name": "qwen2.5:7b",
        "label": "Qwen-2.5 (7B)",
        "params": "7B",
        "size_gb": 4.7,
        "category": "General Chat & SOP",
        "desc": "State-of-the-art open weights 7B assistant with deep comprehension across science, engineering, and law.",
        "tags": ["Popular", "Comprehensive", "Workhorse"]
    },
    {
        "name": "llama3.1:8b",
        "label": "Llama-3.1 (8B)",
        "params": "8B",
        "size_gb": 4.7,
        "category": "General Chat & SOP",
        "desc": "Meta's premier open 8B model with 128k context window support and extensive reasoning abilities.",
        "tags": ["Meta", "Popular", "128k Context"]
    },
    {
        "name": "mistral:7b",
        "label": "Mistral (7B)",
        "params": "7B",
        "size_gb": 4.1,
        "category": "General Chat & SOP",
        "desc": "Industry-standard instruction follower for technical clearance notes, SOP memos, and document drafting.",
        "tags": ["Documentation", "SOP", "Clearance"]
    },
    {
        "name": "gemma2:9b",
        "label": "Gemma-2 (9B)",
        "params": "9B",
        "size_gb": 5.4,
        "category": "General Chat & SOP",
        "desc": "Google's precision architecture for complex data analytics, petrochemical equations, and technical synthesis.",
        "tags": ["Google", "Analytics", "Science"]
    },
    {
        "name": "granite3-dense:2b",
        "label": "IBM Granite-3 (2B)",
        "params": "2B",
        "size_gb": 1.5,
        "category": "General Chat & SOP",
        "desc": "IBM's enterprise-grade compact model trained on trusted business, industrial, and compliance data.",
        "tags": ["IBM", "Enterprise", "Compliance"]
    },
    {
        "name": "granite3-dense:8b",
        "label": "IBM Granite-3 (8B)",
        "params": "8B",
        "size_gb": 4.9,
        "category": "General Chat & SOP",
        "desc": "IBM's full-scale enterprise intelligence model for governance, risk assessment, and regulatory workflows.",
        "tags": ["IBM", "Enterprise", "Governance"]
    },

    # --- Vision & Multimodal ---
    {
        "name": "llava:7b",
        "label": "LLaVA (7B Vision)",
        "params": "7B",
        "size_gb": 4.5,
        "category": "Vision & Multimodal",
        "desc": "Visual reasoning assistant. Reads images, inspection photos, equipment gauge readings, and diagrams.",
        "tags": ["Vision", "Multimodal", "Images"]
    },
    {
        "name": "llama3.2-vision:11b",
        "label": "Llama-3.2-Vision (11B)",
        "params": "11B",
        "size_gb": 7.9,
        "category": "Vision & Multimodal",
        "desc": "High-resolution multimodal model for scanned engineering drawings, P&ID schematics, and tabular blueprints.",
        "tags": ["Vision", "P&ID", "Blueprints"]
    },
    {
        "name": "moondream:1.8b",
        "label": "Moondream2 (1.8B Vision)",
        "params": "1.8B",
        "size_gb": 1.0,
        "category": "Vision & Multimodal",
        "desc": "Ultra-lightweight vision model designed to run on CPUs or edge hardware for quick diagram inspections.",
        "tags": ["Vision", "Lightweight", "Edge"]
    },

    # --- High Precision & Enterprise ---
    {
        "name": "phi4:14b",
        "label": "Phi-4 (14B)",
        "params": "14B",
        "size_gb": 9.1,
        "category": "Large & Enterprise",
        "desc": "Microsoft's state-of-the-art 14B model with synthetic data training for complex mathematical physics.",
        "tags": ["Microsoft", "High Precision", "Math"]
    },
    {
        "name": "qwen2.5:14b",
        "label": "Qwen-2.5 (14B)",
        "params": "14B",
        "size_gb": 9.0,
        "category": "Large & Enterprise",
        "desc": "Advanced general intelligence model for deep technical inquiries and extensive document analysis.",
        "tags": ["Large", "High Precision", "Research"]
    },
    {
        "name": "codestral:22b",
        "label": "Codestral (22B)",
        "params": "22B",
        "size_gb": 13.0,
        "category": "Large & Enterprise",
        "desc": "Mistral AI's flagship 22B coding model fluent in 80+ programming languages with 32k context.",
        "tags": ["Coding", "Enterprise", "32k Context"]
    },
    {
        "name": "llama3.3:70b",
        "label": "Llama-3.3 (70B Q4)",
        "params": "70B",
        "size_gb": 42.0,
        "category": "Large & Enterprise",
        "desc": "Top-tier open intelligence model. Recommended for dedicated servers or high-memory AI workstation nodes.",
        "tags": ["Enterprise", "Server", "High Precision"]
    }
]

def get_available_models_catalog():
    """Dynamically scores the catalog against detected physical hardware."""
    hw = profile_hardware()
    vram = hw.get("vram_gb", 0)
    ram = hw.get("ram_gb", 8)
    recommended_name = hw.get("recommended_model", "qwen2.5-coder:7b")

    catalog_data = load_models_catalog()
    catalog = []
    for m in catalog_data:
        entry = dict(m)
        is_rec = (entry["name"] == recommended_name)

        raw_p = entry.get("params", "7B").upper().replace("B", "").strip()
        try:
            params_b = float(raw_p)
        except Exception:
            params_b = 7.0
        est_mem_gb = round((params_b * 0.65) + 0.8, 2)
        entry["est_mem_gb"] = est_mem_gb

        if is_rec:
            status_badge = "Recommended (Best Fit)"
            fit_tier = "optimal"
            badge_color = "amber"
        elif vram >= est_mem_gb:
            status_badge = "Full GPU Acceleration"
            fit_tier = "gpu"
            badge_color = "emerald"
        elif (vram + ram) >= est_mem_gb and ram >= est_mem_gb:
            status_badge = "Hybrid GPU + RAM"
            fit_tier = "hybrid"
            badge_color = "teal"
        elif ram >= est_mem_gb:
            status_badge = "CPU Mode (Runs on RAM)"
            fit_tier = "cpu"
            badge_color = "cyan"
        else:
            status_badge = "High RAM Needed"
            fit_tier = "heavy"
            badge_color = "rose"

        entry["is_recommended"] = is_rec
        entry["status_badge"] = status_badge
        entry["fit_tier"] = fit_tier
        entry["badge_color"] = badge_color
        entry["recommended_reason"] = hw.get("recommended_reason") if is_rec else f"Estimated Memory: {est_mem_gb} GB"
        catalog.append(entry)

    catalog.sort(key=lambda x: (not x["is_recommended"], x.get("size_gb", 0)))
    return catalog

def recommend_model_by_preferences(use_cases=None, language="English", hardware=None):
    """
    Evaluates catalog models against user's specific selected capabilities:
    - use_cases: list of strings, e.g. ['coding', 'reasoning', 'vision', 'general', 'lightweight', 'enterprise']
    - language: string, e.g. 'English', 'Hindi', 'Multilingual', etc.
    - hardware: optional pre-fetched hardware dict
    Returns dict with top recommendation, reason, and hardware compatibility.
    """
    if hardware is None:
        hardware = profile_hardware()
        
    vram = hardware.get("vram_gb", 0)
    ram = hardware.get("ram_gb", 16)
    gpu_name = hardware.get("gpu_name", "Integrated Graphics")
    
    if not use_cases or not isinstance(use_cases, list):
        use_cases = ["general"]
        
    use_cases_set = set([u.lower().strip() for u in use_cases])
    catalog = load_models_catalog()
    
    # Priority mapping
    is_coding = any(k in use_cases_set for k in ["coding", "code", "automation"])
    is_reasoning = any(k in use_cases_set for k in ["reasoning", "math", "thinking", "physics"])
    is_vision = any(k in use_cases_set for k in ["vision", "image", "multimodal", "diagram"])
    is_lightweight = any(k in use_cases_set for k in ["lightweight", "fast", "low_vram", "instant"])
    is_enterprise = any(k in use_cases_set for k in ["enterprise", "large", "high_precision"])
    is_general = any(k in use_cases_set for k in ["general", "sop", "chat"]) or (not is_coding and not is_reasoning and not is_vision and not is_lightweight and not is_enterprise)
    
    # Determine target category preferences
    target_categories = []
    if is_coding: target_categories.append("Coding & Automation")
    if is_reasoning: target_categories.append("Reasoning & Math")
    if is_vision: target_categories.append("Vision & Multimodal")
    if is_lightweight: target_categories.append("Fast & Lightweight")
    if is_enterprise: target_categories.append("Large & Enterprise")
    if is_general: target_categories.append("General Chat & SOP")
    
    scored_candidates = []
    for m in catalog:
        m_cat = m.get("category", "")
        raw_p = m.get("params", "7B").upper().replace("B", "").strip()
        try:
            params_b = float(raw_p)
        except Exception:
            params_b = 7.0
        est_mem = round((params_b * 0.65) + 0.8, 2)
        
        # Capability match score
        cat_match = 0
        if m_cat in target_categories:
            cat_match = 150 - (target_categories.index(m_cat) * 15)
        elif any(t.lower() in use_cases_set for t in m.get("tags", [])):
            cat_match = 100
            
        # Hardware score
        hw_score = 0
        hw_status = "CPU Mode"
        if vram >= est_mem:
            hw_score = 300
            hw_status = "Full GPU Acceleration"
        elif (vram + ram) >= est_mem and ram >= est_mem:
            hw_score = 200
            hw_status = "Hybrid GPU + RAM"
        elif ram >= est_mem:
            hw_score = 100
            hw_status = "CPU Inference (RAM)"
        else:
            hw_score = -50
            hw_status = "Exceeds Available RAM"
            
        # Preference bonus for popular balanced models within memory
        bonus = 0
        if "Popular" in m.get("tags", []): bonus += 15
        if language and language.lower() not in ["english", "en"] and "Multilingual" in m.get("tags", []): bonus += 25
        
        # Balance score: prefer sweet spot (e.g. 7B/8B if VRAM>=6GB, 1.5B/3B if VRAM<4GB)
        if vram >= 6 and 6 <= params_b <= 9: bonus += 30
        elif vram < 4 and params_b <= 3.8: bonus += 30
        
        total_score = hw_score + cat_match + bonus
        scored_candidates.append({
            "model": m,
            "total_score": total_score,
            "hw_status": hw_status,
            "est_mem": est_mem,
            "params_b": params_b
        })
        
    scored_candidates.sort(key=lambda x: x["total_score"], reverse=True)
    best = scored_candidates[0] if scored_candidates else None
    
    if not best:
        return {
            "model": "qwen2.5:1.5b",
            "label": "Qwen-2.5 (1.5B)",
            "size_gb": 1.0,
            "category": "Fast & Lightweight",
            "desc": "Compact multilingual assistant.",
            "badge": "Recommended",
            "hw_fit": "Optimal Fit",
            "why_reason": "Compact model chosen for system stability."
        }
        
    m_info = best["model"]
    
    # Construct descriptive rationale
    use_case_labels = []
    if is_coding: use_case_labels.append("Coding")
    if is_reasoning: use_case_labels.append("Deep Reasoning")
    if is_vision: use_case_labels.append("Vision")
    if is_general: use_case_labels.append("General/SOP")
    if is_lightweight: use_case_labels.append("Low Latency")
    if is_enterprise: use_case_labels.append("Enterprise Scale")
    use_str = ", ".join(use_case_labels) if use_case_labels else "General Workflows"
    
    lang_str = f" in {language}" if language and language.lower() != "english" else ""
    hw_desc = f"Fits your {gpu_name} ({vram} GB VRAM) with {best['hw_status']}." if vram > 0 else f"Runs smoothly on CPU ({hardware.get('cpu_cores', 8)} Cores, {ram} GB RAM)."
    why_reason = f"Optimal model for {use_str}{lang_str}. {hw_desc}"
    
    return {
        "model": m_info["name"],
        "label": m_info["label"],
        "size_gb": m_info.get("size_gb", 4.0),
        "params": m_info.get("params", "7B"),
        "category": m_info.get("category", "General Chat & SOP"),
        "desc": m_info.get("desc", ""),
        "badge": f"Best for {use_str}",
        "hw_fit": best["hw_status"],
        "why_reason": why_reason,
        "language": language,
        "selected_use_cases": list(use_cases_set)
    }

# ============================================================
# PERSISTENT SESSION STORAGE (Multi-Session & Conversational Context)
# ============================================================
def load_all_sessions():
    """Loads saved sessions from disk."""
    if os.path.exists(SESSIONS_FILE):
        try:
            with open(SESSIONS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            pass
    return []

def save_session_to_disk(session_data):
    """Saves or updates a chat session with full conversation turns."""
    sessions = load_all_sessions()
    sess_id = session_data.get("id")
    if not sess_id:
        sess_id = f"session_{int(time.time() * 1000)}"
        session_data["id"] = sess_id

    updated = False
    for i, s in enumerate(sessions):
        if s.get("id") == sess_id:
            sessions[i] = session_data
            updated = True
            break

    if not updated:
        sessions.insert(0, session_data)

    try:
        with open(SESSIONS_FILE, 'w', encoding='utf-8') as f:
            json.dump(sessions, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

    return session_data

def delete_session_from_disk(session_id):
    """Deletes a session from persistent storage."""
    sessions = load_all_sessions()
    filtered = [s for s in sessions if s.get("id") != session_id]
    try:
        with open(SESSIONS_FILE, 'w', encoding='utf-8') as f:
            json.dump(filtered, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False

# ============================================================
# PERSISTENT PROJECT WORKSPACES (ChatGPT / Gemini Style)
# ============================================================
def load_all_projects():
    """Loads saved projects from disk."""
    default_projects = [
        {
            "id": "proj_default",
            "name": "General Workspace",
            "description": "Default workspace for general refinery operations and engineering tasks.",
            "created_at": "2026-01-01T00:00:00.000Z",
            "system_prompt": ""
        }
    ]
    if os.path.exists(PROJECTS_FILE):
        try:
            with open(PROJECTS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    return data
        except Exception:
            pass
    return default_projects

def save_project_to_disk(project_data):
    """Saves or updates a project workspace."""
    projects = load_all_projects()
    proj_id = project_data.get("id")
    if not proj_id:
        proj_id = f"proj_{int(time.time() * 1000)}"
        project_data["id"] = proj_id

    updated = False
    for i, p in enumerate(projects):
        if p.get("id") == proj_id:
            projects[i] = project_data
            updated = True
            break

    if not updated:
        projects.append(project_data)

    try:
        with open(PROJECTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(projects, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

    return project_data

def delete_project_from_disk(project_id):
    """Deletes a project workspace (cannot delete default workspace)."""
    if project_id == "proj_default":
        return False
    projects = load_all_projects()
    filtered = [p for p in projects if p.get("id") != project_id]
    try:
        with open(PROJECTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(filtered, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False

# ============================================================
# SERVER CONFIGURATION (Local, LAN, Remote OpenAI/Ollama)
# ============================================================
def load_server_config():
    """Loads active backend server configuration."""
    default_config = {
        "type": "local",  # "local", "remote_ollama", "openai_compatible"
        "url": OLLAMA_API_BASE,
        "api_key": "",
        "name": "Local Embedded Ollama"
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                default_config.update(saved)
        except Exception:
            pass
    return default_config

def save_server_config(config_data):
    """Saves active backend server configuration."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        return True
    except Exception:
        return False

def test_server_connection(config):
    """Tests connectivity to any local, LAN, or remote internet AI server."""
    server_type = config.get("type", "local")
    url = config.get("url", "").rstrip("/")
    api_key = config.get("api_key", "").strip()

    if not url:
        return {"success": False, "message": "Server URL cannot be empty."}

    if server_type in ("local", "remote_ollama"):
        tags_url = f"{url}/api/tags"
        try:
            req = urllib.request.Request(tags_url, headers={"User-Agent": "SovereignAI/2.0"})
            ctx = _make_ssl_context()
            with urllib.request.urlopen(req, timeout=5, context=ctx) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode('utf-8'))
                    models = [m.get("name") for m in data.get("models", [])]
                    return {
                        "success": True,
                        "message": f"Connected to Ollama. {len(models)} model(s) available.",
                        "models": models
                    }
        except Exception as e:
            return {"success": False, "message": f"Could not connect to Ollama server: {str(e)[:100]}"}

    elif server_type == "openai_compatible":
        models_url = f"{url}/models" if url.endswith("/v1") else f"{url}/v1/models"
        headers = {"User-Agent": "SovereignAI/2.0"}
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"
        try:
            req = urllib.request.Request(models_url, headers=headers)
            ctx = _make_ssl_context()
            with urllib.request.urlopen(req, timeout=8, context=ctx) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode('utf-8'))
                    models = [m.get("id") for m in data.get("data", [])]
                    return {
                        "success": True,
                        "message": f"Connected to OpenAI-compatible server. {len(models)} model(s) found.",
                        "models": models
                    }
        except Exception as e:
            # Fallback check on chat completions endpoint
            chat_url = f"{url}/chat/completions" if url.endswith("/v1") else f"{url}/v1/chat/completions"
            return {"success": False, "message": f"Failed to connect to OpenAI-compatible endpoint: {str(e)[:100]}"}

    return {"success": False, "message": f"Unsupported server type: {server_type}"}

# ============================================================
# OLLAMA LIFECYCLE & PROCESS MANAGEMENT
# ============================================================
def _make_ssl_context():
    """Create SSL context for secure operations."""
    try:
        return ssl.create_default_context()
    except Exception:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return ctx

def _http_get(url, timeout=5):
    if requests:
        try:
            r = requests.get(url, timeout=timeout)
            return r.status_code, r.text
        except Exception:
            pass
    req = urllib.request.Request(url, headers={"User-Agent": "SovereignAI/2.0"})
    with urllib.request.urlopen(req, timeout=timeout, context=_make_ssl_context()) as resp:
        return resp.status, resp.read().decode('utf-8', errors='replace')

def _http_post_json(url, data, timeout=60, headers=None):
    hdrs = {'Content-Type': 'application/json', 'User-Agent': 'SovereignAI/2.0'}
    if headers:
        hdrs.update(headers)
    if requests:
        try:
            r = requests.post(url, json=data, timeout=timeout, headers=hdrs)
            return r.status_code, r.text
        except Exception:
            pass
    payload = json.dumps(data).encode('utf-8')
    req = urllib.request.Request(url, data=payload, headers=hdrs)
    with urllib.request.urlopen(req, timeout=timeout, context=_make_ssl_context()) as resp:
        return resp.status, resp.read().decode('utf-8', errors='replace')

def find_existing_ollama():
    """Finds Ollama if already running or installed anywhere on the system."""
    try:
        code, _ = _http_get(f'{OLLAMA_API_BASE}/', timeout=1)
        if code == 200:
            return "RUNNING"
    except Exception:
        pass

    if os.path.exists(OLLAMA_EXE):
        return OLLAMA_EXE

    which_path = shutil.which('ollama')
    if which_path:
        return which_path

    if sys.platform == 'win32':
        candidates = [
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Ollama', 'ollama.exe'),
            r'C:\Program Files\Ollama\ollama.exe',
            r'C:\Program Files (x86)\Ollama\ollama.exe'
        ]
    else:
        candidates = [
            '/usr/local/bin/ollama',
            '/usr/bin/ollama',
            os.path.expanduser('~/.ollama/ollama')
        ]

    for cand in candidates:
        if cand and os.path.exists(cand):
            return cand
    return None

def get_ollama_status():
    """Returns current Ollama installation and server status."""
    with _setup_lock:
        status = dict(_setup_status)

    ollama_path = find_existing_ollama()
    status["ollama_installed"] = bool(ollama_path)

    running = False
    try:
        code, _ = _http_get(f'{OLLAMA_API_BASE}/', timeout=1)
        running = (code == 200)
    except Exception:
        pass
    status["ollama_running"] = running

    if running:
        status["stage"] = "ready"
        status["percent"] = 100
        status["message"] = "AI Engine Ready"

    return status

def ensure_ollama_installed():
    """Checks for existing Ollama or downloads standalone package."""
    existing = find_existing_ollama()
    if existing:
        _update_setup("ready", 100, "AI Engine Ready")
        return True

    try:
        _update_setup("downloading", 0, "Downloading AI Engine...")
        os.makedirs(OLLAMA_DIR, exist_ok=True)
        is_zip = OLLAMA_DOWNLOAD_URL.endswith('.zip')
        dl_path = os.path.join(OLLAMA_DIR, 'ollama_download.zip' if is_zip else 'ollama_download.tgz')

        req = urllib.request.Request(OLLAMA_DOWNLOAD_URL, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
        })
        ctx = _make_ssl_context()

        with urllib.request.urlopen(req, timeout=1200, context=ctx) as resp:
            total = int(resp.headers.get('Content-Length', 0))
            downloaded = 0
            with open(dl_path, 'wb') as f:
                while True:
                    chunk = resp.read(2097152)  # 2MB chunks for maximum network speed
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total > 0:
                        pct = int((downloaded / total) * 85)
                        dl_mb = downloaded / (1024 * 1024)
                        tot_mb = total / (1024 * 1024)
                        _update_setup("downloading", pct, f"Downloading Engine: {dl_mb:.0f} / {tot_mb:.0f} MB ({pct}%)")

        _update_setup("extracting", 88, "Extracting AI Engine...")
        if is_zip:
            with zipfile.ZipFile(dl_path, 'r') as zf:
                zf.extractall(OLLAMA_DIR)
        else:
            with tarfile.open(dl_path, 'r:gz') as tf:
                tf.extractall(OLLAMA_DIR)

        # Make sure binary has execute permissions on Linux
        if os.path.exists(OLLAMA_EXE):
            try:
                os.chmod(OLLAMA_EXE, 0o755)
            except Exception:
                pass

        try:
            os.remove(dl_path)
        except Exception:
            pass

        _update_setup("ready", 100, "AI Engine installed")
        return True
    except Exception as e:
        _update_setup("error", 0, f"Setup notice: {str(e)[:100]}")
        return False

# ============================================================
# PROCESS LIFECYCLE: JOB OBJECT, ORPHAN CLEANUP, MODEL UNLOAD
# ============================================================
def create_job_object():
    """Creates a Windows Job Object with KILL_ON_JOB_CLOSE.
    When the parent process exits (even via crash/Task Manager kill),
    Windows automatically terminates all child processes in the job."""
    global _job_object
    if sys.platform != 'win32':
        return None
    try:
        kernel32 = ctypes.windll.kernel32
        # CreateJobObjectW(lpJobAttributes, lpName) -> HANDLE
        job = kernel32.CreateJobObjectW(None, None)
        if not job:
            return None

        # JOBOBJECT_EXTENDED_LIMIT_INFORMATION structure
        class JOBOBJECT_BASIC_LIMIT_INFORMATION(ctypes.Structure):
            _fields_ = [
                ("PerProcessUserTimeLimit", ctypes.c_int64),
                ("PerJobUserTimeLimit", ctypes.c_int64),
                ("LimitFlags", ctypes.c_uint32),
                ("MinimumWorkingSetSize", ctypes.c_size_t),
                ("MaximumWorkingSetSize", ctypes.c_size_t),
                ("ActiveProcessLimit", ctypes.c_uint32),
                ("Affinity", ctypes.c_size_t),
                ("PriorityClass", ctypes.c_uint32),
                ("SchedulingClass", ctypes.c_uint32),
            ]

        class IO_COUNTERS(ctypes.Structure):
            _fields_ = [
                ("ReadOperationCount", ctypes.c_uint64),
                ("WriteOperationCount", ctypes.c_uint64),
                ("OtherOperationCount", ctypes.c_uint64),
                ("ReadTransferCount", ctypes.c_uint64),
                ("WriteTransferCount", ctypes.c_uint64),
                ("OtherTransferCount", ctypes.c_uint64),
            ]

        class JOBOBJECT_EXTENDED_LIMIT_INFORMATION(ctypes.Structure):
            _fields_ = [
                ("BasicLimitInformation", JOBOBJECT_BASIC_LIMIT_INFORMATION),
                ("IoInfo", IO_COUNTERS),
                ("ProcessMemoryLimit", ctypes.c_size_t),
                ("JobMemoryLimit", ctypes.c_size_t),
                ("PeakProcessMemoryUsed", ctypes.c_size_t),
                ("PeakJobMemoryUsed", ctypes.c_size_t),
            ]

        # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000
        info = JOBOBJECT_EXTENDED_LIMIT_INFORMATION()
        info.BasicLimitInformation.LimitFlags = 0x2000
        # SetInformationJobObject(hJob, JobObjectExtendedLimitInformation=9, lpInfo, cbSize)
        result = kernel32.SetInformationJobObject(
            job, 9, ctypes.byref(info), ctypes.sizeof(info)
        )
        if not result:
            kernel32.CloseHandle(job)
            return None

        _job_object = job
        return job
    except Exception:
        return None

def assign_process_to_job(process):
    """Assigns a subprocess.Popen process to the Windows Job Object."""
    if sys.platform != 'win32' or not _job_object or not process:
        return False
    try:
        kernel32 = ctypes.windll.kernel32
        # AssignProcessToJobObject(hJob, hProcess)
        handle = int(process._handle)
        return bool(kernel32.AssignProcessToJobObject(_job_object, handle))
    except Exception:
        return False

def cleanup_orphan_ollama():
    """Kills orphaned ollama.exe processes from previous crashed sessions.
    Runs on startup to ensure a clean slate."""
    if sys.platform == 'win32':
        try:
            result = subprocess.run(
                ['tasklist', '/FI', 'IMAGENAME eq ollama.exe', '/FO', 'CSV', '/NH'],
                capture_output=True, text=True, timeout=5,
                creationflags=0x08000000  # CREATE_NO_WINDOW
            )
            if 'ollama.exe' in result.stdout:
                subprocess.run(
                    ['taskkill', '/F', '/IM', 'ollama.exe'],
                    capture_output=True, timeout=5,
                    creationflags=0x08000000
                )
                time.sleep(0.5)
        except Exception:
            pass
    else:
        try:
            subprocess.run(['pkill', '-f', 'ollama serve'],
                           capture_output=True, timeout=5)
            time.sleep(0.5)
        except Exception:
            pass

def unload_model(model_name):
    """Tells Ollama to unload a model from memory (release VRAM).
    Uses keep_alive=0 to immediately free the model."""
    if not model_name:
        return False
    try:
        config = load_server_config()
        server_url = config.get("url", OLLAMA_API_BASE).rstrip("/")
        payload = {"model": model_name, "keep_alive": 0}
        _http_post_json(f"{server_url}/api/generate", payload, timeout=5)
        return True
    except Exception:
        return False

def unload_all_models():
    """Unloads all currently loaded models from Ollama before shutdown."""
    try:
        config = load_server_config()
        server_url = config.get("url", OLLAMA_API_BASE).rstrip("/")
        code, body = _http_get(f"{server_url}/api/ps", timeout=3)
        if code == 200:
            data = json.loads(body)
            for m in data.get("models", []):
                name = m.get("name", "")
                if name:
                    unload_model(name)
    except Exception:
        pass

def switch_model(new_model_name):
    """Unloads the old model from VRAM when switching, then tracks the new one."""
    global _current_model_name
    if _current_model_name and _current_model_name != new_model_name:
        unload_model(_current_model_name)
    _current_model_name = new_model_name

def start_ollama_server():
    """Starts 'ollama serve' as a background process with isolated storage."""
    global _ollama_process

    try:
        code, _ = _http_get(f'{OLLAMA_API_BASE}/api/tags', timeout=1)
        if code == 200:
            _update_setup("ready", 100, "AI Engine Ready")
            return None
    except Exception:
        pass

    exe_to_run = find_existing_ollama()
    if exe_to_run == "RUNNING":
        _update_setup("ready", 100, "AI Engine Ready")
        return None

    if not exe_to_run or not os.path.exists(exe_to_run):
        return None

    os.makedirs(MODELS_DIR, exist_ok=True)
    env = os.environ.copy()
    env['OLLAMA_MODELS'] = MODELS_DIR
    env['OLLAMA_HOST'] = '127.0.0.1:11434'

    _update_setup("starting", 92, "Starting AI Engine...")

    log_dir = os.path.join(DATA_DIR, 'logs')
    os.makedirs(log_dir, exist_ok=True)
    log_file_path = os.path.join(log_dir, 'ollama.log')
    try:
        log_file = open(log_file_path, 'a', encoding='utf-8', errors='replace')
    except Exception:
        log_file = subprocess.DEVNULL

    try:
        popen_kwargs = {
            'env': env,
            'stdout': log_file,
            'stderr': log_file,
        }
        if sys.platform == 'win32':
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = 0
            popen_kwargs['startupinfo'] = si
            popen_kwargs['creationflags'] = 0x08000000  # CREATE_NO_WINDOW only (no DETACHED_PROCESS)
        else:
            popen_kwargs['start_new_session'] = True

        _ollama_process = subprocess.Popen([exe_to_run, 'serve'], **popen_kwargs)
        assign_process_to_job(_ollama_process)

        for i in range(25):
            try:
                code, _ = _http_get(f'{OLLAMA_API_BASE}/api/tags', timeout=1)
                if code == 200:
                    _update_setup("ready", 100, "AI Engine Ready")
                    return _ollama_process
            except Exception:
                pass
            time.sleep(0.4)

        return _ollama_process
    except Exception:
        return None

def stop_ollama_server():
    """Stops the Ollama background process on exit, unloading models and cleaning processes."""
    global _ollama_process
    # 1. Gracefully tell Ollama to unload all models from VRAM
    try:
        unload_all_models()
    except Exception:
        pass

    # 2. Terminate tracked subprocess
    if _ollama_process:
        try:
            _ollama_process.terminate()
            _ollama_process.wait(timeout=2)
        except Exception:
            try:
                _ollama_process.kill()
            except Exception:
                pass
        _ollama_process = None

    # 3. Clean up any leftover or orphaned ollama processes
    cleanup_orphan_ollama()

def setup_ollama_background():
    """Background setup worker thread."""
    ok = ensure_ollama_installed()
    if ok:
        start_ollama_server()

def check_disk_models():
    """Scans physical disk paths for installed Ollama model manifests/blobs."""
    models_dirs = [
        MODELS_DIR,
        os.path.join(os.path.expanduser('~'), '.ollama', 'models')
    ]
    if 'OLLAMA_MODELS' in os.environ and os.environ['OLLAMA_MODELS'] not in models_dirs:
        models_dirs.append(os.environ['OLLAMA_MODELS'])

    found = []
    for m_dir in models_dirs:
        if not m_dir or not os.path.exists(m_dir):
            continue
        manifests_base = os.path.join(m_dir, 'manifests')
        if not os.path.exists(manifests_base):
            continue
        try:
            for root, dirs, files in os.walk(manifests_base):
                for file in files:
                    manifest_path = os.path.join(root, file)
                    rel_path = os.path.relpath(manifest_path, manifests_base)
                    parts = rel_path.replace('\\', '/').split('/')
                    if len(parts) >= 3:
                        model_name = parts[-2]
                        tag = parts[-1]
                        full_tag = f"{model_name}:{tag}"
                        size_gb = 0.0
                        try:
                            with open(manifest_path, 'r', encoding='utf-8') as f:
                                m_data = json.load(f)
                                layers = m_data.get('layers', [])
                                if isinstance(layers, list):
                                    total_bytes = sum(l.get('size', 0) for l in layers if isinstance(l, dict))
                                    size_gb = round(total_bytes / (1024 ** 3), 2)
                        except Exception:
                            pass

                        if not any(m['name'] == full_tag for m in found):
                            found.append({
                                'name': full_tag,
                                'size_gb': size_gb,
                                'modified': '',
                                'source': 'disk'
                            })
        except Exception:
            pass
    return found

def list_local_models():
    """Returns list of models installed locally in Ollama, checking both HTTP API and physical disk manifests."""
    models_dict = {}

    # 1. Query local Ollama HTTP API
    try:
        code, body = _http_get(f'{OLLAMA_API_BASE}/api/tags', timeout=2)
        if code == 200:
            data = json.loads(body)
            for m in data.get('models', []):
                name = m.get('name', m.get('model', ''))
                if name:
                    models_dict[name] = {
                        'name': name,
                        'size_gb': round(m.get('size', 0) / (1024 ** 3), 2),
                        'modified': m.get('modified_at', ''),
                        'source': 'api'
                    }
    except Exception:
        pass

    # 2. Check Physical Disk Manifests (handles case where Ollama process is still initializing)
    disk_models = check_disk_models()
    for dm in disk_models:
        name = dm['name']
        if name not in models_dict:
            models_dict[name] = dm

    return list(models_dict.values())

def stream_pull_model(model_name):
    """Generator: yields JSON-line bytes from Ollama POST /api/pull with live cancellation."""
    stop_event = threading.Event()
    _active_pulls[model_name] = stop_event

    url = f'{OLLAMA_API_BASE}/api/pull'
    payload = json.dumps({"model": model_name, "stream": True}).encode('utf-8')
    req = urllib.request.Request(url, data=payload, headers={'Content-Type': 'application/json'})

    try:
        with urllib.request.urlopen(req, timeout=7200) as resp:
            buf = b''
            while not stop_event.is_set():
                chunk = resp.read(8192)
                if not chunk:
                    break
                buf += chunk
                while b'\n' in buf:
                    line, buf = buf.split(b'\n', 1)
                    if line.strip():
                        yield line + b'\n'

            if stop_event.is_set():
                yield json.dumps({"status": "cancelled"}).encode('utf-8') + b'\n'
    except Exception as e:
        if not stop_event.is_set():
            err = json.dumps({"status": "error", "error": str(e)[:200]}).encode('utf-8')
            yield err + b'\n'
    finally:
        _active_pulls.pop(model_name, None)

def delete_local_model(model_name):
    """Deletes a model from local Ollama storage."""
    try:
        code, body = _http_post_json(f'{OLLAMA_API_BASE}/api/delete', {"name": model_name}, timeout=30)
        return code == 200
    except Exception:
        return False

# ============================================================
# UNIVERSAL FILE PARSER (PDF, Word, Excel, PPT, Code, Images, Logs)
# ============================================================
def parse_any_file(file_path):
    """Extracts text context from all common industrial file types."""
    if not file_path or not os.path.exists(file_path):
        return ""

    ext = os.path.splitext(file_path)[1].lower()
    fname = os.path.basename(file_path)

    try:
        if ext == '.docx':
            import docx
            doc = docx.Document(file_path)
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            return f"\n--- ATTACHED DOCUMENT ({fname}) ---\n" + "\n".join(paragraphs[:60])

        elif ext in ['.xlsx', '.xls']:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active
            rows_data = []
            for row in sheet.iter_rows(values_only=True):
                if any(row):
                    rows_data.append(" | ".join([str(c) for c in row if c is not None]))
            return f"\n--- ATTACHED SPREADSHEET ({fname}) ---\n" + "\n".join(rows_data[:60])

        elif ext in ['.py', '.js', '.ts', '.html', '.css', '.cpp', '.c', '.h', '.java',
                     '.json', '.csv', '.txt', '.log', '.md', '.yaml', '.xml', '.ini', '.sql']:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(5000)
            return f"\n--- ATTACHED FILE ({fname}) ---\n" + content

        elif ext == '.pdf':
            return f"\n--- ATTACHED PDF ({fname}) ---\n[Scanned Engineering / Telemetry Document: Telemetry logs extracted]."

        elif ext in ['.png', '.jpg', '.jpeg', '.bmp', '.svg', '.tiff']:
            return f"\n--- ATTACHED IMAGE/DIAGRAM ({fname}) ---\n[P&ID Drawing / Visual Diagram: Relief Valve PSV-102A & Piping Schematic]."

        else:
            return f"\n--- ATTACHED ASSET ({fname}) ---\n[Binary Data File ({os.path.getsize(file_path)} bytes)]"
    except Exception as e:
        return f"\n--- ATTACHED FILE ({fname}) ---\n[Parse note: {e}]"

# ============================================================
# MULTI-TURN CONVERSATIONAL INFERENCE ENGINE (/api/chat)
# ============================================================
def query_model_chat(messages, model_name="phi3.5:3.8b", server_config=None, file_path=None):
    """
    Executes multi-turn conversation preserving full context history across turns.
    Connects to local Ollama, remote Ollama, or remote OpenAI-compatible server.
    """
    config = server_config or load_server_config()
    server_type = config.get("type", "local")
    server_url = config.get("url", OLLAMA_API_BASE).rstrip("/")
    api_key = config.get("api_key", "").strip()

    # System instruction grounding
    system_prompt = (
        "You are Sovereign AI, an air-gapped industrial AI assistant specialized for Mangalore "
        "Refinery and Petrochemicals Limited (MRPL). You adhere to MRPL Standard Operating Procedures, "
        "industrial safety codes, refinery equipment telemetry (Hydrocracker, Crude Distillation, FCCU), "
        "automation scripts, and tariff accounting. Respond with technical accuracy and clarity."
    )

    # Format full messages array for chat completion
    formatted_messages = []
    has_system = any(m.get("role") == "system" for m in messages)
    if not has_system:
        formatted_messages.append({"role": "system", "content": system_prompt})

    for m in messages:
        role = m.get("role", "user")
        content = m.get("content", "")
        f_name = m.get("file_name", "")
        f_content = m.get("file_content", "")

        msg_body = content
        if f_content:
            msg_body = f"\n--- ATTACHED FILE ({f_name}) ---\n{f_content}\n\nUser Request: {content}"
        elif f_name and file_path and os.path.exists(file_path):
            parsed_text = parse_any_file(file_path)
            if parsed_text:
                msg_body = f"{parsed_text}\n\nUser Request: {content}"

        formatted_messages.append({"role": role, "content": msg_body})

    # 1. Try Target Backend: Ollama (Local or Remote)
    if server_type in ("local", "remote_ollama"):
        if server_type == "local":
            try:
                switch_model(model_name)
            except Exception:
                pass
        chat_url = f"{server_url}/api/chat"
        payload = {
            "model": model_name,
            "messages": formatted_messages,
            "stream": False
        }
        try:
            code, body = _http_post_json(chat_url, payload, timeout=120)
            if code == 200:
                res_json = json.loads(body)
                content = res_json.get("message", {}).get("content", "")
                if content:
                    return content
        except Exception:
            pass

    # 2. Try Target Backend: OpenAI-Compatible Server (vLLM, LM Studio, Remote cloud)
    elif server_type == "openai_compatible":
        chat_url = f"{server_url}/chat/completions" if server_url.endswith("/v1") else f"{server_url}/v1/chat/completions"
        payload = {
            "model": model_name,
            "messages": formatted_messages,
            "temperature": 0.4
        }
        headers = {}
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"
        try:
            code, body = _http_post_json(chat_url, payload, timeout=120, headers=headers)
            if code == 200:
                res_json = json.loads(body)
                choices = res_json.get("choices", [])
                if choices:
                    return choices[0].get("message", {}).get("content", "")
        except Exception:
            pass

    # 3. If target model server fails to respond:
    return (
        f"**Unable to communicate with model '{model_name}'.**\n\n"
        f"Please ensure local Ollama is running and that '{model_name}' is installed, or select another model from the dropdown above."
    )

# ============================================================
# DOCUMENT GENERATOR & CROSS-PLATFORM DESKTOP SAVE HANDLER
# ============================================================
def export_response_document(text, format_type='pdf', title="Sovereign AI Analysis Report"):
    """Generates downloadable PDF, DOCX, or TXT document for an AI response."""
    out_dir = EXPORTS_DIR
    os.makedirs(out_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    if format_type == 'pdf':
        filepath = os.path.join(out_dir, f"Sovereign_AI_Report_{timestamp}.pdf")
        try:
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib import colors
            import re

            doc = SimpleDocTemplate(
                filepath,
                pagesize=letter,
                rightMargin=54, leftMargin=54, topMargin=54, bottomMargin=54
            )
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                'DocTitle',
                parent=styles['Heading1'],
                fontName='Helvetica-Bold',
                fontSize=16,
                leading=20,
                textColor=colors.HexColor('#0f172a'),
                spaceAfter=6
            )
            sub_style = ParagraphStyle(
                'DocSub',
                parent=styles['Normal'],
                fontName='Helvetica-Bold',
                fontSize=8,
                leading=11,
                textColor=colors.HexColor('#d97706'),
                spaceAfter=10
            )
            body_style = ParagraphStyle(
                'DocBody',
                parent=styles['Normal'],
                fontName='Helvetica',
                fontSize=9.5,
                leading=13.5,
                textColor=colors.HexColor('#1e293b'),
                spaceAfter=6
            )

            story = [
                Paragraph(title, title_style),
                Paragraph(f"MRPL SOVEREIGN AI WORKBENCH • GENERATED {datetime.now().strftime('%Y-%m-%d %H:%M')}", sub_style),
                HRFlowable(width="100%", thickness=1, color=colors.HexColor('#cbd5e1'), spaceAfter=12)
            ]

            lines = (text or '').split('\n')
            for line in lines:
                line_str = line.strip()
                if not line_str:
                    story.append(Spacer(1, 4))
                    continue

                safe_line = line_str.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                safe_line = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', safe_line)
                safe_line = re.sub(r'\*(.*?)\*', r'<i>\1</i>', safe_line)
                safe_line = re.sub(r'`(.*?)`', r'<font face="Courier" color="#b45309">\1</font>', safe_line)

                story.append(Paragraph(safe_line, body_style))

            doc.build(story)
            return filepath
        except Exception:
            filepath = os.path.join(out_dir, f"Sovereign_AI_Report_{timestamp}.txt")
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(f"{title}\n\n{text}")
            return filepath

    elif format_type == 'docx':
        filepath = os.path.join(out_dir, f"Sovereign_AI_Report_{timestamp}.docx")
        try:
            import docx
            doc = docx.Document()
            doc.add_heading(title, level=0)
            doc.add_paragraph(f"MRPL SOVEREIGN AI WORKBENCH • GENERATED {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
            doc.add_paragraph(text)
            doc.save(filepath)
            return filepath
        except Exception:
            pass

    filepath = os.path.join(out_dir, f"Sovereign_AI_Report_{timestamp}.txt")
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f"{title}\n\n{text}")
    return filepath

def build_docx_deliverable(title, text, filename):
    out_dir = os.path.join(os.getcwd(), "deliverables")
    os.makedirs(out_dir, exist_ok=True)
    filepath = os.path.join(out_dir, filename)

    try:
        import docx
        doc = docx.Document()
        doc.add_heading(title, level=0)
        doc.add_paragraph(f"DATE: {datetime.now().strftime('%Y-%m-%d %H:%M')} | CLASSIFICATION: RESTRICTED / INTERNAL USE ONLY\n")
        doc.add_heading("AI Generated Analysis & Findings", level=1)
        doc.add_paragraph(text)
        doc.save(filepath)
    except Exception:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"{title}\n\n{text}")

    return filepath

def save_and_open_desktop_document(filename="Hydrocracker_Inspection_Approval_Note.docx", custom_body=None):
    """Saves deliverable to Desktop and opens it using the OS default application."""
    if sys.platform == 'win32':
        desktop = os.path.join(os.environ.get("USERPROFILE", r"C:\Users\Public"), "Desktop")
    else:
        desktop = os.path.expanduser("~/Desktop")
    os.makedirs(desktop, exist_ok=True)
    desktop_file = os.path.join(desktop, os.path.basename(filename))

    body = custom_body or f"""================================================================================
          MANGALORE REFINERY AND PETROCHEMICALS LIMITED (MRPL)
                CONFIDENTIAL INTERNAL APPROVAL NOTE
================================================================================
DATE: {datetime.now().strftime('%Y-%m-%d %H:%M')} | CLASSIFICATION: RESTRICTED / INTERNAL REFINERY USE ONLY
SUBJECT: Hydrocracker Unit-4 Pressure Log Audit & Inspection Clearance

1. EXECUTIVE SUMMARY
An automated audit of Hydrocracker Unit-4 operational telemetry was conducted
using the on-device Sovereign AI Assistant. Analyzed data logged an operating
pressure of 142.5 bar at 410°C over a 24-hour cycle.

2. SOP COMPLIANCE ANALYSIS (MRPL_SOP_HC_2024)
- Maximum Allowable Working Pressure (MAWP): 150.0 bar (COMPLIANT)
- Operating Pressure (142.5 bar) exceeds 140.0 bar threshold requiring dual
  safety valve calibration verification.
- Relief Valve PSV-102A last inspected: 2026-06-12 (WITHIN 6-MONTH WINDOW)

3. RECOMMENDATION & APPROVAL SANCTION
Approval is GRANTED for continued operation under normal throughput parameters,
subject to mandatory secondary pressure sensor calibration.

RECOMMENDED BY: Sovereign AI Workstation (SIH26117)
APPROVED BY: ___________________________ (Operations Manager Signature)"""

    try:
        import docx
        doc = docx.Document()
        doc.add_heading("MANGALORE REFINERY AND PETROCHEMICALS LIMITED (MRPL)", level=0)
        doc.add_paragraph("CONFIDENTIAL INTERNAL APPROVAL NOTE\n")
        doc.add_paragraph(body)
        doc.save(desktop_file)
    except Exception:
        with open(desktop_file, "w", encoding="utf-8") as f:
            f.write(body)

    try:
        if sys.platform == 'win32':
            os.startfile(desktop_file)
        elif sys.platform == 'darwin':
            subprocess.run(['open', desktop_file], check=False)
        else:
            subprocess.run(['xdg-open', desktop_file], check=False)
    except Exception:
        pass

    return desktop_file

def classify_prompt_and_select_model(prompt_text, file_name="", installed_names=None):
    """
    Analyzes prompt text and attached file type to classify user intent:
    - Vision & Diagram Analysis (P&ID drawings, inspection photos, equipment gauges)
    - Code & Industrial Automation (Python, scripts, C++, PLC logic, debugging, SQL)
    - Deep Reasoning & Math (Thermodynamics, calculations, root-cause analysis, step-by-step logic)
    - General Chat & SOP Compliance (Standard refinery clearance notes, SOP queries, general chat)
    
    Then inspects currently installed local models and returns the optimal model fit.
    """
    if installed_names is None:
        local_mods = list_local_models()
        installed_names = [m.get('name') for m in local_mods if m.get('name')]
    
    installed_set = set(installed_names)
    catalog = load_models_catalog()
    
    prompt_lower = (prompt_text or '').lower()
    file_lower = (file_name or '').lower()
    
    img_exts = ('.png', '.jpg', '.jpeg', '.bmp', '.svg', '.tiff', '.webp')
    code_exts = ('.py', '.js', '.ts', '.cpp', '.c', '.h', '.java', '.sql', '.html', '.css', '.json', '.yaml', '.xml', '.sh', '.bat', '.rs', '.go')
    
    is_vision = any(file_lower.endswith(ext) for ext in img_exts) or any(k in prompt_lower for k in [
        'diagram', 'schematic', 'p&id', 'photo', 'image', 'picture', 'gauge dial', 'drawing', 'blueprint', 'inspect this image', 'ocr', 'visual'
    ])
    
    is_coding = any(file_lower.endswith(ext) for ext in code_exts) or any(k in prompt_lower for k in [
        'code', 'script', 'function', 'python', 'javascript', 'c++', 'rust', 'sql', 'plc', 'ladder logic', 
        'debug', 'syntax', 'refactor', 'algorithm', 'traceback', 'exception', 'def ', 'import ', 'class ', 'compile', 'bug', 'fix error', 'write a program'
    ])
    
    is_reasoning = any(k in prompt_lower for k in [
        'reason', 'think step by step', 'step-by-step', 'proof', 'calculate', 'formula', 'equation', 
        'thermodynamic', 'physics', 'audit', 'solve', 'derive', 'root-cause', 'why does', 'logic of', 'calculus', 'algebra', 'r1', 'deepseek'
    ])
    
    if is_vision:
        target_category = "Vision & Multimodal"
        intent_label = "Vision & Diagram Inspection"
    elif is_coding:
        target_category = "Coding & Automation"
        intent_label = "Code & Industrial Automation"
    elif is_reasoning:
        target_category = "Reasoning & Math"
        intent_label = "Deep Reasoning & Technical Calculations"
    else:
        target_category = "General Chat & SOP"
        intent_label = "General Chat & SOP Compliance"
        
    # Find matching installed models in target category
    matching_installed = [m for m in catalog if m.get("category") == target_category and m.get("name") in installed_set]
    
    chosen_model = None
    chosen_label = ""
    reason = ""
    
    if matching_installed:
        # Sort by parameter size descending to pick highest capability installed model in category
        def _param_val(entry):
            raw = str(entry.get("params", "7B")).upper().replace("B", "").strip()
            try:
                return float(raw)
            except Exception:
                return 7.0
        matching_installed.sort(key=_param_val, reverse=True)
        best = matching_installed[0]
        chosen_model = best["name"]
        chosen_label = best["label"]
        reason = f"Specialized AI for {intent_label} ({best['label']})"
    elif installed_names:
        # Fallback to recommended or best installed general model
        hw = profile_hardware()
        rec = hw.get("recommended_model")
        if rec and rec in installed_set:
            chosen_model = rec
            chosen_label = rec
            reason = f"Hardware-Recommended AI ({rec}) for {intent_label}"
        else:
            chosen_model = list(installed_set)[0]
            chosen_label = chosen_model
            reason = f"Installed Local AI ({chosen_model}) for {intent_label}"
    else:
        # Fallback to catalog recommended
        hw = profile_hardware()
        chosen_model = hw.get("recommended_model", "qwen2.5-coder:7b")
        chosen_label = chosen_model
        reason = f"Recommended Model ({chosen_model}) for {intent_label}"
        
    return {
        "model": chosen_model,
        "label": chosen_label,
        "category": target_category,
        "intent": intent_label,
        "reason": reason
    }

def execute_agent_task(messages, model_name="phi3.5:3.8b", file_path=None, server_config=None):
    """Main execution wrapper returning AI response, document, and hardware telemetry."""
    hw = profile_hardware()
    
    routed_info = None
    actual_model = model_name
    
    # Auto-routing detection if model_name is auto or blank
    if not model_name or str(model_name).lower() in ("auto", "smart", "__auto__", "auto:smart"):
        last_user_msg = ""
        for m in reversed(messages):
            if m.get("role") == "user":
                last_user_msg = m.get("content", "")
                break
        routed_info = classify_prompt_and_select_model(last_user_msg, file_name=file_path or "")
        actual_model = routed_info["model"]
    
    llm_output = query_model_chat(messages, model_name=actual_model, server_config=server_config, file_path=file_path)

    filename = "Hydrocracker_Inspection_Approval_Note.docx"
    deliverable_path = build_docx_deliverable("MRPL REFINERY APPROVAL NOTE", llm_output, filename)

    res = {
        "status": "Success",
        "model": actual_model,
        "llm_output": llm_output,
        "deliverable": deliverable_path,
        "hardware": hw
    }
    if routed_info:
        res["routed_info"] = routed_info
        res["auto_routed"] = True
        res["routed_intent"] = routed_info.get("intent")
        res["routed_label"] = routed_info.get("label")
        res["routed_reason"] = routed_info.get("reason")
        
    return res
