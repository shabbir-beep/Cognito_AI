"""
Generate SIH 2026 Presentation PDF for Problem Statement SIH26117
Cognito On-Premise Agentic AI Workbench - MRPL
"""

import os
import sys
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super(NumberedCanvas, self).__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super(NumberedCanvas, self).showPage()
        super(NumberedCanvas, self).save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        
        # Header banner bar
        self.setFillColor(colors.HexColor("#0f172a")) # Slate 900
        self.rect(0, 580, 792, 32, fill=1, stroke=0)
        
        self.setFillColor(colors.HexColor("#f59e0b")) # Amber 500
        self.rect(0, 577, 792, 3, fill=1, stroke=0)
        
        # Header text
        self.setFont("Helvetica-Bold", 10)
        self.setFillColor(colors.white)
        self.drawString(36, 592, "SMART INDIA HACKATHON 2026  •  COGNITO ON-PREMISE AGENTIC AI WORKBENCH")
        
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.HexColor("#cbd5e1"))
        self.drawRightString(756, 592, "PS ID: SIH26117  |  MRPL")
        
        # Footer banner bar
        self.setFillColor(colors.HexColor("#f8fafc"))
        self.rect(0, 0, 792, 28, fill=1, stroke=0)
        self.setStrokeColor(colors.HexColor("#e2e8f0"))
        self.setLineWidth(1)
        self.line(0, 28, 792, 28)
        
        # Footer text
        self.setFont("Helvetica-Bold", 8)
        self.setFillColor(colors.HexColor("#d97706"))
        self.drawString(36, 10, "@SIH Idea submission - Template  |  Mangalore Refinery and Petrochemicals Limited")
        
        self.setFont("Helvetica-Bold", 8)
        self.setFillColor(colors.HexColor("#475569"))
        self.drawRightString(756, 10, f"Slide {self._pageNumber} of {page_count}")
        
        self.restoreState()

def build_presentation_pdf(output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(letter), # 792 x 612
        leftMargin=36,
        rightMargin=36,
        topMargin=48,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    
    # Custom palette styles
    title_style = ParagraphStyle(
        'SlideTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#0f172a'),
        alignment=1, # Center
        spaceAfter=12
    )
    
    section_heading = ParagraphStyle(
        'SecHeading',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=15,
        textColor=colors.HexColor('#0f172a'),
        spaceAfter=6
    )

    card_title = ParagraphStyle(
        'CardTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#0f172a'),
        spaceAfter=4
    )

    body_text = ParagraphStyle(
        'SlideBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11.5,
        textColor=colors.HexColor('#334155'),
        spaceAfter=4
    )
    
    body_bold = ParagraphStyle(
        'SlideBodyBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11.5,
        textColor=colors.HexColor('#0f172a'),
        spaceAfter=4
    )

    badge_style = ParagraphStyle(
        'Badge',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#b45309'),
        alignment=1
    )

    table_header = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=1
    )

    table_cell = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#1e293b')
    )

    table_cell_bold = ParagraphStyle(
        'TableCellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#0f172a')
    )

    story = []

    # =========================================================================
    # SLIDE 1: Title Slide
    # =========================================================================
    story.append(Spacer(1, 15))
    story.append(Paragraph("SMART INDIA HACKATHON 2026", ParagraphStyle('SIHHeader', fontName='Helvetica-Bold', fontSize=22, leading=26, textColor=colors.HexColor('#d97706'), alignment=1)))
    story.append(Paragraph("COGNITO ON-PREMISE AGENTIC AI WORKBENCH", ParagraphStyle('SIHSub', fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=colors.HexColor('#0f172a'), alignment=1)))
    story.append(Spacer(1, 15))

    info_data = [
        [
            Paragraph("<b>Problem Statement ID:</b>", body_bold),
            Paragraph("<b>SIH26117</b>", ParagraphStyle('SIHBadge', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#d97706')))
        ],
        [
            Paragraph("<b>Problem Statement Title:</b>", body_bold),
            Paragraph("Cognito On-Premise Agentic AI Workbench using Open-Weight Multimodal LLMs for Confidential Industrial Work", body_text)
        ],
        [
            Paragraph("<b>Organization:</b>", body_bold),
            Paragraph("<b>Mangalore Refinery and Petrochemicals Limited (MRPL)</b>", body_text)
        ],
        [
            Paragraph("<b>Theme & PS Category:</b>", body_bold),
            Paragraph("Smart Automation  |  Software", body_text)
        ],
        [
            Paragraph("<b>Team ID & Name:</b>", body_bold),
            Paragraph("[Registered Team ID]  —  [Registered Team Name]", body_text)
        ],
        [
            Paragraph("<b>Platform & Architecture:</b>", body_bold),
            Paragraph("Air-Gapped, Zero-Install Portable Windows/Linux Executable with Embedded Ollama Runtime", body_text)
        ]
    ]

    t_info = Table(info_data, colWidths=[180, 520])
    t_info.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8fafc')),
        ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor('#cbd5e1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
    ]))
    story.append(t_info)
    
    story.append(Spacer(1, 20))
    summary_box = [
        [Paragraph(
            "<b>Project Scope Summary:</b> An industrial-grade, fully sovereign, on-premise desktop AI workbench "
            "designed for MRPL refineries. Provides automated multi-turn engineering reasoning, SOP compliance audits, "
            "P&ID drawing inspection, hardware-aware 32-model management, and automatic executive report generation "
            "under strict zero-external-network isolation.",
            ParagraphStyle('SumText', fontName='Helvetica', fontSize=8.5, leading=12, textColor=colors.HexColor('#0f172a'))
        )]
    ]
    t_sum = Table(summary_box, colWidths=[700])
    t_sum.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#fef3c7')), # Amber 100
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#f59e0b')),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
    ]))
    story.append(t_sum)

    # =========================================================================
    # SLIDE 2: Problem & Proposed Solution
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("COGNITO AI WORKBENCH — PROBLEM & SOLUTION", title_style))
    story.append(Spacer(1, 4))

    col1 = [
        Paragraph("<b>THE PROBLEM</b>", ParagraphStyle('RedTitle', fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=colors.HexColor('#b91c1c'))),
        Spacer(1, 4),
        Paragraph("• <b>Confidential Data Risk:</b> Sensitive approval notes, P&IDs, Hydrocracker / FCCU telemetry logs, and internal operating code cannot leave refinery premises.", body_text),
        Paragraph("• <b>Cloud AI Policy Gap:</b> Commercial cloud AI assistants (OpenAI, Claude) create catastrophic data leakage and cybersecurity compliance violations.", body_text),
        Paragraph("• <b>Operational Bottleneck:</b> Manual log verification and cross-referencing against heavy SOP manuals delays maintenance clearance and reduces plant throughput.", body_text),
        Spacer(1, 4),
        Paragraph("<b>THE PROPOSED SOLUTION</b>", ParagraphStyle('GreenTitle', fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=colors.HexColor('#15803d'))),
        Spacer(1, 4),
        Paragraph("• <b>100% Self-Hosted & Cognito:</b> Fully offline desktop workstation running on refinery hardware with verifiable zero external network traffic.", body_text),
        Paragraph("• <b>Hardware-Aware Intelligence:</b> Auto-detects real CPU/RAM/VRAM and scores a 32-model catalog to run the best-fit model comfortably.", body_text),
        Paragraph("• <b>Multi-Format File Ingestion:</b> Native extraction from Word, Excel telemetry sheets, PDFs, and P&ID diagrams.", body_text),
        Paragraph("• <b>Automated Deliverable Dispatch:</b> One-click generation of signed executive approval notes in DOCX and PDF directly to Desktop.", body_text),
    ]

    col2 = [
        Paragraph("<b>CORE AGENTIC PILLARS & WORKBENCH CAPABILITIES</b>", ParagraphStyle('BlueTitle', fontName='Helvetica-Bold', fontSize=10.5, leading=13, textColor=colors.HexColor('#0369a1'))),
        Spacer(1, 4),
        Paragraph("<b>1. Multi-Step Industrial Planning:</b> AI audits live telemetry against MRPL SOP limits (e.g., MAWP $\\le$ 150 bar, 6-month PSV calibration window) and formats structured approval notes.", body_text),
        Paragraph("<b>2. Local Tool & File Integration:</b> Universal parser extracts text, tables, and sensor metrics from `.docx`, `.xlsx`, `.pdf`, code, and logs.", body_text),
        Paragraph("<b>3. Open-Weight 32-Model Hub:</b> Categorized repository (Coding, DeepSeek-R1 Reasoning, Vision, General SOP, Enterprise) with 1-click NDJSON streaming downloads and cancellation.", body_text),
        Paragraph("<b>4. Scan & Drawing Understanding:</b> Multimodal visual models (LLaVA, Moondream, Llama-3.2-Vision) inspect P&ID schematics and equipment gauges.", body_text),
        Paragraph("<b>5. VRAM & Process Safety Sandbox:</b> Windows Job Object (`KILL_ON_JOB_CLOSE`), orphan cleanup, and `keep_alive=0` memory release.", body_text),
        Paragraph("<b>6. Executive Deliverable Export:</b> Styled ReportLab PDF and python-docx approval note generator dispatched directly to OS Desktop.", body_text)
    ]

    t_s2 = Table([[col1, col2]], colWidths=[345, 355])
    t_s2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#f8fafc')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#f0fdf4')),
        ('BOX', (0,0), (0,0), 1, colors.HexColor('#cbd5e1')),
        ('BOX', (1,0), (1,0), 1, colors.HexColor('#86efac')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t_s2)

    # =========================================================================
    # SLIDE 3: Technical Approach
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("TECHNICAL APPROACH & ARCHITECTURE", title_style))
    story.append(Spacer(1, 4))

    arch_left = [
        Paragraph("<b>1. LAYERED LOCAL WORKBENCH STACK</b>", card_title),
        Paragraph("• <b>Presentation Layer:</b> Single-Page Desktop App in PyWebView (HTML5/CSS3/JS); 5 luxury preset themes (Midnight Obsidian, Titanium, Azure, Matrix, Amethyst) + custom palette builder; real-time stop generation button.", body_text),
        Paragraph("• <b>Backend Server (main.py):</b> Multithreaded TCP Server on port 8085; REST API + NDJSON streaming progress endpoints; daemon thread dispatch.", body_text),
        Paragraph("• <b>Core Engine (sovereign_engine.py):</b> Hardware detection, 32-model catalog scorer, multi-turn chat orchestrator, universal file parser.", body_text),
        Paragraph("• <b>Process & Memory Supervisor:</b> Win32 Job Object (`0x2000`), startup orphan process killer, dynamic VRAM model unloader (`keep_alive=0`).", body_text),
        Paragraph("• <b>Inference Layer:</b> Embedded portable Ollama daemon (`port 11434`) running quantized GGUF weights on CUDA GPU / CPU AVX2.", body_text),
        Paragraph("• <b>Air-Gap Isolation:</b> Hardcoded loopback (`127.0.0.1`), zero outbound calls, standalone `data/` directory structure.", body_text)
    ]

    arch_right = [
        Paragraph("<b>2. STREAMLINED INSPECTION & APPROVAL WORKFLOW</b>", card_title),
        Paragraph("<b>Step 1: Telemetry & Report Ingestion:</b> Operator attaches Hydrocracker / Crude Distillation logs, Excel sensor data, or P&ID drawings.", body_text),
        Paragraph("<b>Step 2: Context Parsing:</b> Universal parser extracts operating pressures (142.5 bar), temperatures (410°C), and valve tags (PSV-102A).", body_text),
        Paragraph("<b>Step 3: Hardware-Guided Model Routing:</b> Engine matches prompt complexity to the optimal installed model (e.g., DeepSeek-R1 / Qwen2.5).", body_text),
        Paragraph("<b>Step 4: SOP Compliance Audit:</b> LLM checks operating pressure against MRPL SOP limits (MAWP: 150 bar, 140 bar threshold audit).", body_text),
        Paragraph("<b>Step 5: Executive Deliverable Generation:</b> Auto-formats structured findings with executive summary, compliance status, and signature lines.", body_text),
        Paragraph("<b>Step 6: Direct Desktop Dispatch:</b> Builds formatted `.docx` and styled `.pdf` documents directly on Desktop and launches in Word.", body_text)
    ]

    t_s3 = Table([[arch_left, arch_right]], colWidths=[345, 355])
    t_s3.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#f8fafc')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#eff6ff')),
        ('BOX', (0,0), (0,0), 1, colors.HexColor('#cbd5e1')),
        ('BOX', (1,0), (1,0), 1, colors.HexColor('#93c5fd')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t_s3)

    story.append(Spacer(1, 8))
    comp_banner = [
        [Paragraph(
            "<b>Implementation Components:</b> "
            "<b>Agent Harness:</b> Multithreaded Python (`main.py` + `sovereign_engine.py`) | "
            "<b>Inference:</b> Embedded Portable Ollama + GGUF Quantized Models (CUDA / CPU AVX2) | "
            "<b>File Ingestion:</b> Native `python-docx`, `openpyxl`, PDF & text parsers | "
            "<b>UI Framework:</b> PyWebView + Luxury Responsive HTML5 SPA (5 Themes) | "
            "<b>Safety & Isolation:</b> Win32 Job Object (`KILL_ON_JOB_CLOSE`) + Dynamic VRAM Unloader (`keep_alive: 0`) + Zero-Egress Loopback",
            ParagraphStyle('CompText', fontName='Helvetica', fontSize=7.5, leading=10, textColor=colors.HexColor('#0f172a'))
        )]
    ]
    t_comp = Table(comp_banner, colWidths=[700])
    t_comp.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f1f5f9')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#94a3b8')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_comp)

    # =========================================================================
    # SLIDE 4: Feasibility and Viability
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("FEASIBILITY, VIABILITY & RISK MITIGATION", title_style))
    story.append(Spacer(1, 4))

    # Stepper Flow
    stepper_data = [
        [
            Paragraph("<b>1. Foundation</b><br/>Portable EXE, server & embedded engine", ParagraphStyle('Step1', fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=1)),
            Paragraph("<b>2. Model Hub</b><br/>32 models, 1-click streaming pull", ParagraphStyle('Step2', fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=1)),
            Paragraph("<b>3. Guardrails</b><br/>Job Object, VRAM unloader & orphan killer", ParagraphStyle('Step3', fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=1)),
            Paragraph("<b>4. Agentic Core</b><br/>Multi-turn memory & MRPL SOP prompt", ParagraphStyle('Step4', fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=1)),
            Paragraph("<b>5. File Ingestion</b><br/>DOCX, XLSX, PDF & P&ID parser", ParagraphStyle('Step5', fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=1)),
            Paragraph("<b>6. Deliverables</b><br/>ReportLab PDF & Word to Desktop", ParagraphStyle('Step6', fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=1))
        ]
    ]
    t_step = Table(stepper_data, colWidths=[116, 116, 116, 116, 116, 116])
    t_step.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8fafc')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#cbd5e1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_step)
    story.append(Spacer(1, 8))

    # Feasibility Table + Risk Table
    feas_head = [Paragraph("<b>Single Workstation Feasibility</b>", table_header), Paragraph("<b>Implementation & Verification</b>", table_header)]
    feas_rows = [
        feas_head,
        [Paragraph("<b>Open-Weight Models</b>", table_cell_bold), Paragraph("Curated 32-model catalog (0.135B to 70B) covering Qwen2.5, DeepSeek-R1, Llama-3.2, Gemma-2, Moondream, Mistral.", table_cell)],
        [Paragraph("<b>Quantized GGUF</b>", table_cell_bold), Paragraph("4-bit/8-bit quantization fits low-end GPUs (2GB VRAM GT 710) up to multi-GPU high-memory servers.", table_cell)],
        [Paragraph("<b>Embedded Runtime</b>", table_cell_bold), Paragraph("Standalone Ollama daemon with isolated storage (<code>data/models/</code>); zero admin rights or installs needed.", table_cell)],
        [Paragraph("<b>Portable Storage</b>", table_cell_bold), Paragraph("All settings, sessions, models, and exports stored in portable <code>data/</code> folder packaged in single ~21MB ZIP.", table_cell)],
    ]
    t_feas = Table(feas_rows, colWidths=[120, 220])
    t_feas.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0f172a')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#cbd5e1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))

    risk_head = [Paragraph("<b>Identified Risk</b>", table_header), Paragraph("<b>Control Implemented</b>", table_header), Paragraph("<b>Live Demo Evidence</b>", table_header)]
    risk_rows = [
        risk_head,
        [
            Paragraph("<b>Model Exceeds VRAM (OOM)</b>", table_cell_bold),
            Paragraph("Hardware profiler + formula: $\\text{Mem} = \\text{Params} \\times 0.65 + 0.8\\text{ GB}$", table_cell),
            Paragraph("Dynamic compatibility badges (<code>⭐ Optimal</code>, <code>🟢 Full GPU</code>, <code>🟡 CPU</code>) prevent OOM.", table_cell)
        ],
        [
            Paragraph("<b>VRAM Leak / Zombie Processes</b>", table_cell_bold),
            Paragraph("Win32 Job Object (<code>0x2000</code>) + dynamic <code>keep_alive: 0</code> model unloader", table_cell),
            Paragraph("Verified via nvidia-smi: exactly 0 MB VRAM retained after exit or model switch.", table_cell)
        ],
        [
            Paragraph("<b>Data Leakage via Cloud</b>", table_cell_bold),
            Paragraph("Hardcoded loopback (<code>127.0.0.1</code>) with zero external calls", table_cell),
            Paragraph("Wireshark network monitor records 0 outbound packets during chat and export.", table_cell)
        ],
        [
            Paragraph("<b>Interrupted Downloads</b>", table_cell_bold),
            Paragraph("NDJSON streaming reader with thread cancellation event", table_cell),
            Paragraph("Live byte progress bar with instant responsive <b>Cancel</b> button.", table_cell)
        ],
    ]
    t_risk = Table(risk_rows, colWidths=[100, 125, 125])
    t_risk.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#b45309')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#cbd5e1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))

    t_s4_tables = Table([[t_feas, t_risk]], colWidths=[345, 355])
    t_s4_tables.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(t_s4_tables)

    # =========================================================================
    # SLIDE 5: Impact and Benefits
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("IMPACT, BENEFITS & INDUSTRIAL VALUE", title_style))
    story.append(Spacer(1, 4))

    col_impact_1 = [
        Paragraph("<b>WORKFLOW TRANSFORMATION</b>", card_title),
        Paragraph("<b>⚙️ Gear 1: Profile & Route:</b> Hardware profiler auto-detects GPU/RAM and routes task to optimal model from 32-model hub.", body_text),
        Paragraph("<b>⚙️ Gear 2: Read & Ingest:</b> Universal parser extracts telemetry metrics from Excel sensor logs, DOCX, and P&ID drawings.", body_text),
        Paragraph("<b>⚙️ Gear 3: Audit in Sandbox:</b> Local LLM evaluates operational parameters against MRPL SOP limits (MAWP threshold, valve dates).", body_text),
        Paragraph("<b>⚙️ Gear 4: Deliver & Clear:</b> Exports signed executive approval note to Desktop and releases VRAM immediately.", body_text),
        Spacer(1, 4),
        Paragraph("<b>WHY IT MATTERS</b>", card_title),
        Paragraph("• <b>Absolute Confidentiality:</b> Telemetry logs, tariff accounting, and internal memos never leave refinery boundaries.", body_text),
        Paragraph("• <b>Zero-Friction Portability:</b> Single ~21MB ZIP runs out-of-the-box on Windows/Linux without admin privileges or pip installs.", body_text),
        Paragraph("• <b>Hardware Inclusivity:</b> Runs on modest 2GB VRAM cards (GT 710) up to multi-GPU enterprise AI workstations.", body_text),
    ]

    col_impact_2 = [
        Paragraph("<b>STAKEHOLDER VALUE MATRIX (BEFORE VS. AFTER)</b>", card_title),
        Spacer(1, 2),
        Paragraph("<b>👨‍💼 Plant & Operations Engineers:</b>", body_bold),
        Paragraph("<i>Before:</i> Manual spreadsheet parsing & cross-referencing against 200+ page SOP manuals.<br/><b>After:</b> Instant file ingestion, automated telemetry audits, and evidence-backed draft notes.", body_text),
        Spacer(1, 2),
        Paragraph("<b>🛡️ Operations & Safety Approvers:</b>", body_bold),
        Paragraph("<i>Before:</i> Slow, multi-day document review cycles causing operational delays.<br/><b>After:</b> Structured, standardized, and traceable compliance memos generated in seconds.", body_text),
        Spacer(1, 2),
        Paragraph("<b>🔒 IT & Cybersecurity Officers:</b>", body_bold),
        Paragraph("<i>Before:</i> Massive risk of proprietary refinery data leakage via public cloud LLMs.<br/><b>After:</b> 100% on-premise air-gapped execution with verifiable zero egress traffic.", body_text),
        Spacer(1, 2),
        Paragraph("<b>🏢 Executive Management:</b>", body_bold),
        Paragraph("<i>Before:</i> High recurring cloud API costs, vendor lock-in, and compliance penalties.<br/><b>After:</b> Zero license costs, offline USB-deployable solution across all refinery units.", body_text),
    ]

    t_s5 = Table([[col_impact_1, col_impact_2]], colWidths=[345, 355])
    t_s5.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#f8fafc')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#fefce8')),
        ('BOX', (0,0), (0,0), 1, colors.HexColor('#cbd5e1')),
        ('BOX', (1,0), (1,0), 1, colors.HexColor('#fde047')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t_s5)

    # =========================================================================
    # SLIDE 6: Research, References & Implementation Roadmap
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("RESEARCH, REFERENCES & IMPLEMENTATION ROADMAP", title_style))
    story.append(Spacer(1, 4))

    col_chk = [
        Paragraph("<b>LIVE DEMONSTRATION CHECKLIST</b>", card_title),
        Paragraph("<b>[x] Hardware Auto-Profiling:</b> Accurately detects CPU cores, RAM, free disk, and GPU VRAM via nvidia-smi and auto-recommends optimal model.", body_text),
        Paragraph("<b>[x] Interactive 32-Model Hub:</b> 1-click NDJSON streaming download with real-time progress bar and responsive cancel support.", body_text),
        Paragraph("<b>[x] Multi-Format Context Ingestion:</b> Extracts and injects text from Word documents, Excel telemetry sheets, PDFs, and code files into the prompt.", body_text),
        Paragraph("<b>[x] Executive Deliverable Dispatch:</b> One-click generation of styled ReportLab PDFs and formatted Microsoft Word (`.docx`) approval memos directly to Desktop.", body_text),
        Paragraph("<b>[x] 100% VRAM & Process Safety:</b> Windows Job Object (`KILL_ON_JOB_CLOSE`) + `keep_alive: 0` model unloading ensures 0 MB VRAM leakage and zero orphan processes on exit.", body_text),
    ]

    col_strat = [
        Paragraph("<b>10-STAGE IMPLEMENTATION ROADMAP</b>", card_title),
        Paragraph("<b>1. SIH Scope:</b> Defined air-gapped, multimodal sovereign AI scope for MRPL refinery.<br/>"
                  "<b>2. Portable Architecture:</b> Built self-contained <code>data/</code> folder structure and standalone EXE.<br/>"
                  "<b>3. Embedded Inference:</b> Bundled standalone Ollama daemon with isolated <code>OLLAMA_MODELS</code>.<br/>"
                  "<b>4. Universal Ingestion:</b> Native python-docx, openpyxl, PDF and diagram text extractors.<br/>"
                  "<b>5. Strict Air-Gap Proof:</b> Hardcoded loopback (<code>127.0.0.1</code>) with zero outbound calls.<br/>"
                  "<b>6. Hardware Auto-Scoring:</b> Memory formula scoring with dynamic visual fit badges.<br/>"
                  "<b>7. Curated 32-Model Hub:</b> Categorized catalog with 1-click download and live progress bars.<br/>"
                  "<b>8. VRAM & Process Guard:</b> Win32 Job Object supervisor and dynamic <code>keep_alive: 0</code> unloader.<br/>"
                  "<b>9. Document Generator:</b> ReportLab PDF formatting and direct desktop Word dispatch.<br/>"
                  "<b>10. Luxury Modern UI:</b> 5 preset themes, custom palette builder, and stop response controller.",
                  ParagraphStyle('StratText', fontName='Helvetica', fontSize=7.5, leading=9.8, textColor=colors.HexColor('#0f172a')))
    ]

    t_s6 = Table([[col_chk, col_strat]], colWidths=[335, 365])
    t_s6.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#f0fdf4')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#f8fafc')),
        ('BOX', (0,0), (0,0), 1, colors.HexColor('#86efac')),
        ('BOX', (1,0), (1,0), 1, colors.HexColor('#cbd5e1')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t_s6)

    story.append(Spacer(1, 8))
    ref_banner = [
        [Paragraph(
            "<b>Technical References & URLs:</b> "
            "<b>SIH Portal:</b> sih.gov.in/sih2026PS (Problem Statement SIH26117) | "
            "<b>Inference Engine:</b> ollama.com, github.com/ollama/ollama, github.com/ggml-org/llama.cpp | "
            "<b>Parsing & Docs:</b> github.com/python-openxml/python-docx, github.com/openpyxl/openpyxl, github.com/reportlab/reportlab | "
            "<b>Open-Weight Models:</b> github.com/qwenlm/Qwen2.5, github.com/deepseek-ai/DeepSeek-R1, github.com/meta-llama/llama-models, github.com/vikhyat/moondream",
            ParagraphStyle('RefText', fontName='Helvetica', fontSize=7, leading=9.5, textColor=colors.HexColor('#475569'))
        )]
    ]
    t_ref = Table(ref_banner, colWidths=[700])
    t_ref.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f1f5f9')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#94a3b8')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_ref)

    # Build document
    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Presentation PDF successfully generated at: {output_path}")

if __name__ == '__main__':
    root_dir = os.path.dirname(os.path.abspath(__file__))
    out_pdf = os.path.join(root_dir, "SIH2026_SIH26117_Cognito_AI_Workbench_Presentation.pdf")
    build_presentation_pdf(out_pdf)
