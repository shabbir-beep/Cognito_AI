"""
Build Complete SIH 2026 PowerPoint Presentation with Dedicated Chart/Graph Layouts
Project: SIH26117 | MRPL Sovereign On-Premise Agentic AI Workbench
Format: Exact 16:9 Widescreen SIH Template with distinct visual containers,
structured tables, wheel diagrams, stacked cylinder, flowcharts, and risk matrices.
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

def build_presentation(output_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- Color Palette ---
    C_WHITE      = RGBColor(255, 255, 255)
    C_DARK_TEXT  = RGBColor(15, 23, 42)     # #0F172A
    C_SLATE_900  = RGBColor(15, 23, 42)     # #0F172A
    C_SLATE_800  = RGBColor(30, 41, 59)     # #1E293B
    C_SLATE_700  = RGBColor(51, 65, 85)     # #334155
    C_SLATE_500  = RGBColor(100, 116, 139)  # #64748B
    C_SLATE_300  = RGBColor(203, 213, 225)  # #CBD5E1
    C_SLATE_100  = RGBColor(241, 245, 249)  # #F1F5F9
    C_SLATE_50   = RGBColor(248, 250, 252)  # #F8FAFC

    C_SIH_BLUE   = RGBColor(2, 132, 199)    # #0284C7
    C_ORANGE_600 = RGBColor(234, 88, 12)    # #EA580C
    C_ORANGE_500 = RGBColor(249, 115, 22)   # #F97316
    C_ORANGE_100 = RGBColor(255, 237, 213)  # #FFEDD5
    
    C_GREEN_700  = RGBColor(21, 128, 61)    # #15803D
    C_GREEN_600  = RGBColor(22, 163, 74)    # #16A34A
    C_GREEN_100  = RGBColor(220, 252, 231)  # #DCFCE7
    C_GREEN_50   = RGBColor(240, 253, 244)  # #F0FDF4

    C_CYAN_700   = RGBColor(14, 116, 144)   # #0E7490
    C_CYAN_500   = RGBColor(6, 182, 212)    # #06B6D4
    C_CYAN_100   = RGBColor(207, 250, 254)  # #CFFAFE

    C_YELLOW_600 = RGBColor(202, 138, 4)    # #CA8A04
    C_YELLOW_500 = RGBColor(234, 179, 8)    # #EAB308
    C_YELLOW_100 = RGBColor(254, 240, 138)  # #FEF08A

    C_PINK_600   = RGBColor(219, 39, 119)   # #DB2777
    C_PINK_500   = RGBColor(236, 72, 153)   # #EC4899
    C_PINK_100   = RGBColor(252, 231, 243)  # #FCE7F3
    C_PINK_50    = RGBColor(253, 242, 248)  # #FDF2F8

    C_BLUE_700   = RGBColor(29, 78, 216)    # #1D4ED8
    C_BLUE_500   = RGBColor(59, 130, 246)   # #3B82F6
    C_BLUE_100   = RGBColor(219, 234, 254)  # #DBEAFE
    C_BLUE_50    = RGBColor(239, 246, 255)  # #EFF6FF

    C_PURPLE_600 = RGBColor(147, 51, 234)   # #9333EA
    C_PURPLE_100 = RGBColor(243, 232, 255)  # #F3E8FF

    def draw_sih_logo(slide, left, top):
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(2.2), Inches(0.9))
        box.fill.solid()
        box.fill.fore_color.rgb = C_WHITE
        box.line.color.rgb = C_SLATE_300
        box.line.width = Pt(0.5)

        tf = box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = "SMART INDIA"
        p1.font.name = "Arial"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = C_DARK_TEXT
        p1.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = "HACKATHON 2026"
        p2.font.name = "Arial"
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = C_ORANGE_600
        p2.alignment = PP_ALIGN.CENTER

        p3 = tf.add_paragraph()
        p3.text = "[ SIH26117 | MRPL ]"
        p3.font.name = "Arial"
        p3.font.size = Pt(8)
        p3.font.bold = True
        p3.font.color.rgb = C_GREEN_700
        p3.alignment = PP_ALIGN.CENTER

    def add_common_header_footer(slide, title_text, page_num):
        # 1. Top-Left Team Name Oval Badge
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.4), Inches(0.3), Inches(1.5), Inches(0.95))
        oval.fill.solid()
        oval.fill.fore_color.rgb = C_WHITE
        oval.line.color.rgb = C_SLATE_700
        oval.line.width = Pt(1.5)
        tf_ov = oval.text_frame
        p_ov1 = tf_ov.paragraphs[0]
        p_ov1.text = "Team"
        p_ov1.font.name = "Arial"
        p_ov1.font.size = Pt(9.5)
        p_ov1.font.color.rgb = C_DARK_TEXT
        p_ov1.alignment = PP_ALIGN.CENTER
        p_ov2 = tf_ov.add_paragraph()
        p_ov2.text = "Name"
        p_ov2.font.name = "Arial"
        p_ov2.font.size = Pt(9.5)
        p_ov2.font.color.rgb = C_DARK_TEXT
        p_ov2.alignment = PP_ALIGN.CENTER

        # 2. Top-Center Title
        tx_title = slide.shapes.add_textbox(Inches(2.1), Inches(0.38), Inches(8.5), Inches(0.8))
        tf_t = tx_title.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.name = "Georgia"
        p_t.font.size = Pt(21)
        p_t.font.bold = True
        p_t.font.color.rgb = C_DARK_TEXT
        p_t.alignment = PP_ALIGN.CENTER

        # 3. Top-Right SIH Logo
        draw_sih_logo(slide, Inches(10.7), Inches(0.3))

        # 4. Bottom Blue Footer Banner
        ft = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4))
        ft.fill.solid()
        ft.fill.fore_color.rgb = C_SIH_BLUE
        ft.line.color.rgb = C_SIH_BLUE

        tx_ft = slide.shapes.add_textbox(Inches(0.5), Inches(7.12), Inches(12.333), Inches(0.35))
        tf_ft = tx_ft.text_frame
        p_ft = tf_ft.paragraphs[0]
        p_ft.text = f"@SIH Idea submission- Template  |  Mangalore Refinery and Petrochemicals Limited (MRPL)                                           {page_num}"
        p_ft.font.name = "Arial"
        p_ft.font.size = Pt(9.5)
        p_ft.font.color.rgb = C_WHITE

    # =========================================================================
    # SLIDE 1: Title Slide (Metadata + Hardware Memory Allocation Chart)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)

    # Top Center Title
    tx_s1_t = s1.shapes.add_textbox(Inches(1.5), Inches(0.45), Inches(8.5), Inches(1.4))
    tf_s1_t = tx_s1_t.text_frame
    p_t1 = tf_s1_t.paragraphs[0]
    p_t1.text = "SMART INDIA HACKATHON 2026"
    p_t1.font.name = "Georgia"
    p_t1.font.size = Pt(23)
    p_t1.font.bold = True
    p_t1.font.color.rgb = C_DARK_TEXT
    p_t1.alignment = PP_ALIGN.CENTER

    p_t2 = tf_s1_t.add_paragraph()
    p_t2.text = "SOVEREIGN ON-PREMISE\nAGENTIC AI WORKBENCH"
    p_t2.font.name = "Georgia"
    p_t2.font.size = Pt(17)
    p_t2.font.bold = True
    p_t2.font.color.rgb = C_SIH_BLUE
    p_t2.alignment = PP_ALIGN.CENTER

    draw_sih_logo(s1, Inches(10.7), Inches(0.35))

    # Left Column: Detailed Metadata Bullets
    tx_meta = s1.shapes.add_textbox(Inches(0.6), Inches(1.95), Inches(6.6), Inches(5.0))
    tf_m = tx_meta.text_frame
    tf_m.word_wrap = True

    meta_bullets = [
        ("• Problem Statement ID: ", "SIH26117", True),
        ("• Problem Statement Title: ", "Sovereign On-Premise Agentic AI Workbench using Open-Weight Multimodal LLMs for Confidential Industrial Work", False),
        ("• Organization: ", "Mangalore Refinery and Petrochemicals Limited (MRPL)", True),
        ("• Theme & Category: ", "Smart Automation  |  Software", False),
        ("• Team ID & Name: ", "[To be updated / Registered Team ID] — [Registered Team Name]", False),
        ("• Tech Stack: ", "Python 3.10-3.14 (Multithreaded Server), PyWebView SPA (HTML5/CSS3/JS), Standalone Embedded Ollama (Port 11434), ReportLab PDF, python-docx, openpyxl", False),
        ("• Deployment: ", "Zero-Install Single-File Portable Windows Executable (SovereignAIWorkbench.exe, ~21MB) with self-contained data/ folder structure", False),
    ]

    for idx, (label, val, highlight) in enumerate(meta_bullets):
        p = tf_m.paragraphs[0] if idx == 0 else tf_m.add_paragraph()
        p.text = label + val
        p.font.name = "Arial"
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_DARK_TEXT
        p.font.bold = True if highlight else False
        p.space_after = Pt(7)

    # Right Area: Chart / Data Table for Hardware Memory Footprint & Model Fitting
    chart_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.4), Inches(1.95), Inches(5.4), Inches(5.0))
    chart_box.fill.solid()
    chart_box.fill.fore_color.rgb = C_SLATE_50
    chart_box.line.color.rgb = C_SLATE_300
    chart_box.line.width = Pt(1.5)

    tf_cb = chart_box.text_frame
    tf_cb.word_wrap = True
    p_cb1 = tf_cb.paragraphs[0]
    p_cb1.text = "📊 [CHART DATA] HARDWARE VRAM FIT MATRIX"
    p_cb1.font.name = "Arial"
    p_cb1.font.size = Pt(10.5)
    p_cb1.font.bold = True
    p_cb1.font.color.rgb = C_ORANGE_600
    p_cb1.alignment = PP_ALIGN.CENTER

    # Memory allocation sub-table
    t_hw_shape = s1.shapes.add_table(6, 3, Inches(7.55), Inches(2.45), Inches(5.1), Inches(2.6))
    t_hw = t_hw_shape.table
    t_hw.columns[0].width = Inches(1.8)
    t_hw.columns[1].width = Inches(1.3)
    t_hw.columns[2].width = Inches(2.0)

    hw_rows = [
        ("Model Tier", "Size / VRAM", "Target Hardware / Fit"),
        ("Qwen2.5 (0.5B - 1.5B)", "0.4 - 1.0 GB", "🟢 Full GPU on 2GB GT 710"),
        ("Phi-3.5 / Llama-3.2 (3B)", "1.9 - 2.2 GB", "⚡ Hybrid GPU (2GB) + RAM"),
        ("Qwen2.5-Coder / R1 (7B)", "4.7 GB", "🟢 Full GPU on 6GB RTX 3060"),
        ("DeepSeek-R1 / Phi-4 (14B)", "9.0 - 9.1 GB", "🟢 Full GPU on 12GB RTX 4070"),
        ("Llama-3.3 (70B Q4)", "42.0 GB", "🏢 Multi-GPU Workstation / Server")
    ]
    for r_idx, (c1, c2, c3) in enumerate(hw_rows):
        cell1, cell2, cell3 = t_hw.cell(r_idx, 0), t_hw.cell(r_idx, 1), t_hw.cell(r_idx, 2)
        if r_idx == 0:
            for cell, text in [(cell1, c1), (cell2, c2), (cell3, c3)]:
                cell.fill.solid(); cell.fill.fore_color.rgb = C_SLATE_900
                p = cell.text_frame.paragraphs[0]; p.text = text; p.font.bold = True; p.font.color.rgb = C_WHITE; p.font.size = Pt(8)
        else:
            for cell, text in [(cell1, c1), (cell2, c2), (cell3, c3)]:
                cell.fill.solid(); cell.fill.fore_color.rgb = C_WHITE
                p = cell.text_frame.paragraphs[0]; p.text = text; p.font.color.rgb = C_SLATE_800; p.font.size = Pt(7.5)

    # Core Metric Note below table
    tx_note = s1.shapes.add_textbox(Inches(7.55), Inches(5.15), Inches(5.1), Inches(1.6))
    tf_nt = tx_note.text_frame
    tf_nt.word_wrap = True
    p_nt1 = tf_nt.paragraphs[0]
    p_nt1.text = "• Auto-Scoring Formula: Memory (GB) = (Params * 0.65) + 0.8 GB\n• Zero-Install Binary: SovereignAIWorkbench.exe (20.87 MB)\n• Windows Job Object (0x2000): 0 MB leftover VRAM on exit"
    p_nt1.font.name = "Arial"
    p_nt1.font.size = Pt(8)
    p_nt1.font.color.rgb = C_SLATE_700

    # =========================================================================
    # SLIDE 2: Problem & Proposed Solution + Wheel + Cylinder
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s2, "SOVEREIGN AI WORKBENCH", 2)

    # Left Column: Problem & Proposed Solution
    tx_p = s2.shapes.add_textbox(Inches(0.4), Inches(1.35), Inches(3.6), Inches(2.6))
    tf_p = tx_p.text_frame
    tf_p.word_wrap = True
    p_ph = tf_p.paragraphs[0]
    p_ph.text = "THE PROBLEM"
    p_ph.font.name = "Arial"
    p_ph.font.size = Pt(12)
    p_ph.font.bold = True
    p_ph.font.color.rgb = C_DARK_TEXT
    
    p_pd = tf_p.add_paragraph()
    p_pd.text = "Confidential refinery telemetry across Hydrocracker Unit-4, FCCU, and Crude Distillation units, P&IDs, relief valve inspection notes (PSV-102A), MAWP calculations, and control code cannot leave premises. Public cloud LLMs create severe data leakage risks, while manual SOP cross-referencing causes massive clearance bottlenecks."
    p_pd.font.name = "Arial"
    p_pd.font.size = Pt(8.2)
    p_pd.font.color.rgb = C_SLATE_700

    # Mini badge: Data Confidentiality Gap
    bg_p = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(3.55), Inches(3.5), Inches(0.65))
    bg_p.fill.solid()
    bg_p.fill.fore_color.rgb = C_PINK_50
    bg_p.line.color.rgb = C_PINK_500
    tf_bgp = bg_p.text_frame
    p_bgp1 = tf_bgp.paragraphs[0]
    p_bgp1.text = "Data Confidentiality Gap"
    p_bgp1.font.bold = True
    p_bgp1.font.size = Pt(8.5)
    p_bgp1.font.color.rgb = C_PINK_600
    p_bgp2 = tf_bgp.add_paragraph()
    p_bgp2.text = "Industrial telemetry & P&IDs cannot leave refinery premises"
    p_bgp2.font.size = Pt(7.2)
    p_bgp2.font.color.rgb = C_SLATE_700

    # Solution
    tx_s = s2.shapes.add_textbox(Inches(0.4), Inches(4.25), Inches(3.6), Inches(2.0))
    tf_s = tx_s.text_frame
    tf_s.word_wrap = True
    p_sh = tf_s.paragraphs[0]
    p_sh.text = "THE PROPOSED SOLUTION"
    p_sh.font.name = "Arial"
    p_sh.font.size = Pt(12)
    p_sh.font.bold = True
    p_sh.font.color.rgb = C_DARK_TEXT

    p_sd = tf_s.add_paragraph()
    p_sd.text = "A 100% self-hosted, portable AI workbench executing on the organization's own workstation or GPU server. It plans multi-step work, executes local file parsing tools, dynamically selects among 32 installed open-weight models, understands P&ID drawings, and delivers usable Word/PDF deliverables with verifiable zero external network traffic."
    p_sd.font.name = "Arial"
    p_sd.font.size = Pt(8.2)
    p_sd.font.color.rgb = C_SLATE_700

    # Mini badge: Secure Data Processing
    bg_s = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(6.3), Inches(3.5), Inches(0.65))
    bg_s.fill.solid()
    bg_s.fill.fore_color.rgb = C_BLUE_50
    bg_s.line.color.rgb = C_BLUE_500
    tf_bgs = bg_s.text_frame
    p_bgs1 = tf_bgs.paragraphs[0]
    p_bgs1.text = "Secure Data Processing"
    p_bgs1.font.bold = True
    p_bgs1.font.size = Pt(8.5)
    p_bgs1.font.color.rgb = C_BLUE_700
    p_bgs2 = tf_bgs.add_paragraph()
    p_bgs2.text = "Executive Word & PDF files generated with 0 external calls"
    p_bgs2.font.size = Pt(7.2)
    p_bgs2.font.color.rgb = C_SLATE_700

    # Center: Circular Wheel Diagram
    wheel_title = s2.shapes.add_textbox(Inches(4.1), Inches(1.35), Inches(4.8), Inches(0.4))
    p_wt = wheel_title.text_frame.paragraphs[0]
    p_wt.text = "🎡 [WHEEL CHART] 6 AGENTIC PILLARS"
    p_wt.font.name = "Arial"
    p_wt.font.size = Pt(10)
    p_wt.font.bold = True
    p_wt.font.color.rgb = C_DARK_TEXT
    p_wt.alignment = PP_ALIGN.CENTER

    w_center = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.6), Inches(3.4), Inches(1.8), Inches(1.8))
    w_center.fill.solid()
    w_center.fill.fore_color.rgb = C_WHITE
    w_center.line.color.rgb = C_SLATE_700
    w_center.line.width = Pt(2)
    tf_wc = w_center.text_frame
    p_wc = tf_wc.paragraphs[0]
    p_wc.text = "Self-Hosted\nAI\nWorkbench"
    p_wc.font.name = "Arial"
    p_wc.font.size = Pt(9.5)
    p_wc.font.bold = True
    p_wc.font.color.rgb = C_DARK_TEXT
    p_wc.alignment = PP_ALIGN.CENTER

    pillar_nodes = [
        (Inches(5.1), Inches(1.7), Inches(2.8), Inches(0.8), "1. Multi-Step Work Planning", "Audits Hydrocracker MAWP (<=150 bar, 140 bar limit) & valve logs against MRPL SOPs", C_ORANGE_600, C_ORANGE_100),
        (Inches(7.2), Inches(2.55), Inches(2.3), Inches(0.8), "2. Local Tool Integration", "Parses DOCX, Excel telemetry, PDFs, code, & logs locally with zero cloud API dependencies", C_YELLOW_600, C_YELLOW_100),
        (Inches(7.2), Inches(4.7), Inches(2.3), Inches(0.8), "3. Open-Weight Selection", "32-model catalog (Coding, DeepSeek-R1, Vision, SOP, 70B) with 1-click streaming pull", C_GREEN_600, C_GREEN_100),
        (Inches(5.1), Inches(5.6), Inches(2.8), Inches(0.8), "4. Scan & Drawing Reasoning", "LLaVA-7B & Moondream-1.8B inspect scanned P&ID schematics and equipment gauge dials", C_CYAN_700, C_CYAN_100),
        (Inches(3.7), Inches(4.7), Inches(2.2), Inches(0.8), "5. Verifiable Zero-External", "Strict loopback 127.0.0.1; Wireshark confirms zero outbound packets during chat/export", C_BLUE_700, C_BLUE_100),
        (Inches(3.7), Inches(2.55), Inches(2.2), Inches(0.8), "6. VRAM & Process Safety", "Win32 Job Object (KILL_ON_JOB_CLOSE) + dynamic keep_alive=0 VRAM memory release", C_PINK_600, C_PINK_100),
    ]
    for left, top, width, height, title, desc, col, bg_col in pillar_nodes:
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_col
        card.line.color.rgb = col
        card.line.width = Pt(1)
        tf_c = card.text_frame
        p_c1 = tf_c.paragraphs[0]
        p_c1.text = title
        p_c1.font.bold = True
        p_c1.font.size = Pt(7.5)
        p_c1.font.color.rgb = col
        p_c1.alignment = PP_ALIGN.CENTER
        p_c2 = tf_c.add_paragraph()
        p_c2.text = desc
        p_c2.font.size = Pt(6.5)
        p_c2.font.color.rgb = C_SLATE_700
        p_c2.alignment = PP_ALIGN.CENTER

    # Right Column: Stacked Capabilities Cylinder
    cyl_title = s2.shapes.add_textbox(Inches(9.4), Inches(1.3), Inches(3.6), Inches(0.6))
    tf_cyl = cyl_title.text_frame
    tf_cyl.word_wrap = True
    p_cyl_t = tf_cyl.paragraphs[0]
    p_cyl_t.text = "🏛️ [CYLINDER HIERARCHY] SCOPE"
    p_cyl_t.font.name = "Arial"
    p_cyl_t.font.size = Pt(8.5)
    p_cyl_t.font.bold = True
    p_cyl_t.font.color.rgb = C_DARK_TEXT
    p_cyl_t.alignment = PP_ALIGN.CENTER

    layers = [
        ("Tier 5: Multimodal Deliverables", "Exports styled ReportLab PDF & python-docx notes with executive approval sign-offs to Desktop", C_GREEN_600, C_GREEN_100),
        ("Tier 4: Local Knowledge Base", "Grounding in MRPL refinery SOPs, Hydrocracker Unit-4 telemetry, and multi-turn session persistence", C_BLUE_700, C_BLUE_100),
        ("Tier 3: Permission-Gated Sandbox", "Windows Job Object (0x2000), hidden subprocess execution (CREATE_NO_WINDOW), & orphan cleanup", C_PINK_600, C_PINK_100),
        ("Tier 2: Automatic Model Routing", "Routes prompt complexity to optimal model (Qwen2.5-Coder for code, DeepSeek-R1 for math/audit)", C_ORANGE_600, C_ORANGE_100),
        ("Tier 1: Hardware Recommendation", "Calculates GPU VRAM/RAM fit (Mem = Params * 0.65 + 0.8 GB) to prevent out-of-memory crashes", C_YELLOW_600, C_YELLOW_100),
    ]

    for idx, (title, desc, col, bg_col) in enumerate(layers):
        l_top = Inches(2.0 + idx * 0.95)
        cyl_l = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.6), l_top, Inches(3.4), Inches(0.85))
        cyl_l.fill.solid()
        cyl_l.fill.fore_color.rgb = bg_col
        cyl_l.line.color.rgb = col
        cyl_l.line.width = Pt(1.5)

        tf_l = cyl_l.text_frame
        tf_l.word_wrap = True
        pl1 = tf_l.paragraphs[0]
        pl1.text = title
        pl1.font.bold = True
        pl1.font.size = Pt(8.5)
        pl1.font.color.rgb = col

        pl2 = tf_l.add_paragraph()
        pl2.text = desc
        pl2.font.size = Pt(7.0)
        pl2.font.color.rgb = C_SLATE_700

    # =========================================================================
    # SLIDE 3: Technical Approach (Layered Stack + 6-Step Workflow)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s3, "TECHNICAL APPROACH", 3)

    # Left: Building a Comprehensive Local System
    t_left = s3.shapes.add_textbox(Inches(0.6), Inches(1.35), Inches(5.8), Inches(0.4))
    p_tl = t_left.text_frame.paragraphs[0]
    p_tl.text = "🏗️ [STACK DIAGRAM] COMPREHENSIVE LOCAL SYSTEM"
    p_tl.font.name = "Arial"
    p_tl.font.size = Pt(10.5)
    p_tl.font.bold = True
    p_tl.font.color.rgb = C_DARK_TEXT
    p_tl.alignment = PP_ALIGN.CENTER

    left_nodes = [
        ("Layer 1: Client UI (PyWebView)", "Lightweight single-file SPA (HTML5/CSS3/JS); 5 luxury themes (Midnight Obsidian, Titanium, Azure, Matrix, Amethyst) + custom palette builder; real-time stop generation button.", C_CYAN_700, C_CYAN_100),
        ("Layer 2: Server (main.py)", "Multithreaded TCP Server on port 8085; low-latency REST API + streaming NDJSON progress endpoints; daemon thread dispatch; signal handlers (SIGINT/SIGTERM/SIGBREAK).", C_GREEN_700, C_GREEN_100),
        ("Layer 3: Guardrail Supervisor", "Win32 Job Object supervisor (0x2000) prevents background zombies; dynamic keep_alive=0 unloads models from VRAM immediately upon model switch or application exit.", C_BLUE_700, C_BLUE_100),
        ("Layer 4: Tools & Parsing Engine", "Universal parser (DOCX python-docx, Excel openpyxl, PDF, code, logs) & persistent multi-turn JSON session database (data/sessions/sessions.json).", C_YELLOW_600, C_YELLOW_100),
        ("Layer 5: Inference Runtime", "Embedded standalone Ollama daemon (port 11434) running GGUF quantized models on CUDA GPU / CPU AVX2, with remote LAN & OpenAI-compatible fallback.", C_ORANGE_600, C_ORANGE_100),
        ("Layer 6: Air-Gap Isolation", "Hardcoded localhost 127.0.0.1 binding; zero cloud telemetry; direct OS-level desktop document generation and automated application dispatch.", C_PINK_600, C_PINK_100),
    ]
    for idx, (title, desc, col, bg_col) in enumerate(left_nodes):
        node_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.8 + idx * 0.65), Inches(3.8), Inches(0.58))
        node_box.fill.solid()
        node_box.fill.fore_color.rgb = bg_col
        node_box.line.color.rgb = col
        tf_nb = node_box.text_frame
        p1 = tf_nb.paragraphs[0]
        p1.text = title + ": "
        p1.font.bold = True
        p1.font.size = Pt(7.8)
        p1.font.color.rgb = col
        p2 = tf_nb.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(6.8)
        p2.font.color.rgb = C_SLATE_700

    # Hub: Unified System
    hub1 = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.7), Inches(3.2), Inches(1.7), Inches(1.7))
    hub1.fill.solid()
    hub1.fill.fore_color.rgb = C_BLUE_50
    hub1.line.color.rgb = C_BLUE_700
    hub1.line.width = Pt(1.5)
    tf_h1 = hub1.text_frame
    ph1_t = tf_h1.paragraphs[0]
    ph1_t.text = "Unified System"
    ph1_t.font.bold = True
    ph1_t.font.size = Pt(8.5)
    ph1_t.font.color.rgb = C_BLUE_700
    ph1_t.alignment = PP_ALIGN.CENTER
    ph1_d = tf_h1.add_paragraph()
    ph1_d.text = "Decoupled server, engine, & storage integrated in a single portable bundle."
    ph1_d.font.size = Pt(6.2)
    ph1_d.font.color.rgb = C_SLATE_700
    ph1_d.alignment = PP_ALIGN.CENTER

    # Right: Streamlined Inspection Process
    t_right = s3.shapes.add_textbox(Inches(6.8), Inches(1.35), Inches(6.0), Inches(0.4))
    p_tr = t_right.text_frame.paragraphs[0]
    p_tr.text = "🔄 [FLOWCHART] 6-STEP REFINERY INSPECTION PIPELINE"
    p_tr.font.name = "Arial"
    p_tr.font.size = Pt(10.5)
    p_tr.font.bold = True
    p_tr.font.color.rgb = C_DARK_TEXT
    p_tr.alignment = PP_ALIGN.CENTER

    right_steps = [
        ("Step 1: Telemetry & Log Upload", "Operator attaches Hydrocracker logs, Excel sensor data, or P&ID drawings via attachment bar.", C_GREEN_700, C_GREEN_100),
        ("Step 2: Universal Context Extraction", "Parser extracts operating pressures (142.5 bar), temps (410°C), and valve tags (PSV-102A).", C_YELLOW_600, C_YELLOW_100),
        ("Step 3: Hardware-Guided Model Routing", "Engine matches prompt complexity to optimal installed model (DeepSeek-R1 / Qwen2.5 / Moondream).", C_CYAN_700, C_CYAN_100),
        ("Step 4: SOP Compliance Reasoner", "LLM audits telemetry against MRPL SOP rules (MAWP <= 150 bar; flags 142.5 bar calibration requirement).", C_PURPLE_600, C_PURPLE_100),
        ("Step 5: Guardrail Approval & Formatting", "AI formats structured findings with executive summary, compliance status, & sign-off block.", C_PINK_600, C_PINK_100),
        ("Step 6: Desktop Document Dispatch", "Builds formatted .docx and styled .pdf documents directly on Desktop and launches in Word/Acrobat.", C_ORANGE_600, C_ORANGE_100),
    ]
    for idx, (title, desc, col, bg_col) in enumerate(right_steps):
        step_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8 + idx * 0.65), Inches(3.8), Inches(0.58))
        step_box.fill.solid()
        step_box.fill.fore_color.rgb = bg_col
        step_box.line.color.rgb = col
        tf_sb = step_box.text_frame
        p1 = tf_sb.paragraphs[0]
        p1.text = title + ": "
        p1.font.bold = True
        p1.font.size = Pt(7.8)
        p1.font.color.rgb = col
        p2 = tf_sb.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(6.8)
        p2.font.color.rgb = C_SLATE_700

    # Hub: Comprehensive Inspection Report
    hub2 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.9), Inches(3.2), Inches(2.0), Inches(1.7))
    hub2.fill.solid()
    hub2.fill.fore_color.rgb = C_GREEN_50
    hub2.line.color.rgb = C_GREEN_700
    hub2.line.width = Pt(1.5)
    tf_h2 = hub2.text_frame
    ph2_t = tf_h2.paragraphs[0]
    ph2_t.text = "Comprehensive\nInspection Report"
    ph2_t.font.bold = True
    ph2_t.font.size = Pt(8.5)
    ph2_t.font.color.rgb = C_GREEN_700
    ph2_t.alignment = PP_ALIGN.CENTER
    ph2_d = tf_h2.add_paragraph()
    ph2_d.text = "Official MRPL compliance note with full telemetry audit trail & signature block."
    ph2_d.font.size = Pt(6.2)
    ph2_d.font.color.rgb = C_SLATE_700
    ph2_d.alignment = PP_ALIGN.CENTER

    # Bottom Banner: Implementation Components
    tx_ic = s3.shapes.add_textbox(Inches(0.6), Inches(5.9), Inches(12.133), Inches(1.0))
    tf_ic = tx_ic.text_frame
    tf_ic.word_wrap = True
    pic_h = tf_ic.paragraphs[0]
    pic_h.text = "IMPLEMENTATION COMPONENTS & INDUSTRIAL INTEGRATION"
    pic_h.font.name = "Arial"
    pic_h.font.size = Pt(10)
    pic_h.font.bold = True
    pic_h.font.color.rgb = C_DARK_TEXT
    pic_h.alignment = PP_ALIGN.CENTER

    pic_d = tf_ic.add_paragraph()
    pic_d.text = "Agent Harness: Multithreaded Python server (main.py) + sovereign_engine.py | Inference: Embedded Portable Ollama + GGUF Quantized Models (CUDA / CPU AVX2) | Ingestion: python-docx, openpyxl, PDF text extraction | UI: Single-File HTML5/CSS3 SPA in PyWebView (5 Luxury Themes + Custom Palette) | Isolation: Windows Job Object (0x2000) + keep_alive=0 VRAM release + Zero-Egress Loopback"
    pic_d.font.name = "Arial"
    pic_d.font.size = Pt(8.0)
    pic_d.font.color.rgb = C_SLATE_800
    pic_d.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 4: Feasibility, Stepper, Risk Matrix & Concentric Gates
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s4, "FEASIBILITY AND VIABILITY", 4)

    # Top-Left Title: Feasibility on a Single Workstation
    tx_f_t = s4.shapes.add_textbox(Inches(0.4), Inches(1.3), Inches(4.5), Inches(0.35))
    p_ft = tx_f_t.text_frame.paragraphs[0]
    p_ft.text = "📋 [TABLE] FEASIBILITY ON SINGLE WORKSTATION"
    p_ft.font.name = "Arial"
    p_ft.font.size = Pt(9.5)
    p_ft.font.bold = True
    p_ft.font.color.rgb = C_DARK_TEXT

    # Feasibility Table
    t_feas_shape = s4.shapes.add_table(5, 2, Inches(0.4), Inches(1.65), Inches(4.4), Inches(5.15))
    tf_tab = t_feas_shape.table
    tf_tab.columns[0].width = Inches(1.5)
    tf_tab.columns[1].width = Inches(2.9)

    feas_rows = [
        ("Feature", "Implementation & Technical Verification", C_SLATE_100, C_DARK_TEXT, True),
        ("Open-Weight Models", "Curated 32-model catalog (0.135B to 70B) covering Qwen2.5-Coder, DeepSeek-R1, Llama-3.2, Gemma-2, Moondream, Mistral, and Codestral.", C_YELLOW_100, C_YELLOW_600, False),
        ("Quantized GGUF Models", "4-bit (Q4_K_M) & 8-bit (Q8_0) quantization allows full GPU offload on modest 2GB VRAM cards (Nvidia GT 710) up to multi-GPU servers.", C_ORANGE_100, C_ORANGE_600, False),
        ("Embedded Runtime", "Standalone embedded Ollama daemon with isolated storage (data/models/); runs with zero administrative rights or external dependencies.", C_PINK_100, C_PINK_600, False),
        ("Portable Data Storage", "All settings, sessions, models, and exports stored in portable data/ folder packaged in a single ~21MB ZIP archive.", C_PINK_50, C_PINK_600, False),
    ]
    for r_idx, (c1, c2, bg_col, txt_col, is_head) in enumerate(feas_rows):
        cell1, cell2 = tf_tab.cell(r_idx, 0), tf_tab.cell(r_idx, 1)
        cell1.fill.solid(); cell1.fill.fore_color.rgb = bg_col
        cell2.fill.solid(); cell2.fill.fore_color.rgb = C_WHITE if is_head else C_SLATE_50
        p1 = cell1.text_frame.paragraphs[0]; p1.text = c1; p1.font.bold = True; p1.font.size = Pt(8.0); p1.font.color.rgb = txt_col
        p2 = cell2.text_frame.paragraphs[0]; p2.text = c2; p2.font.bold = is_head; p2.font.size = Pt(7.5); p2.font.color.rgb = C_SLATE_700

    # Top-Right: Build an Agentic System Stepper (6 cards)
    tx_stepper = s4.shapes.add_textbox(Inches(5.0), Inches(1.3), Inches(7.8), Inches(0.35))
    p_stp = tx_stepper.text_frame.paragraphs[0]
    p_stp.text = "📈 [STEPPER FLOW] 6-STAGE AGENTIC ROADMAP"
    p_stp.font.name = "Arial"
    p_stp.font.size = Pt(10)
    p_stp.font.bold = True
    p_stp.font.color.rgb = C_DARK_TEXT
    p_stp.alignment = PP_ALIGN.CENTER

    stepper_nodes = [
        ("1. Foundation", "Portable EXE + multithreaded server + embedded Ollama runtime", C_BLUE_700, C_BLUE_100),
        ("2. Model Layer", "Hardware profiler (nvidia-smi) + 32-model hub streaming pull", C_CYAN_700, C_CYAN_100),
        ("3. Guardrails", "Win32 Job Object supervisor + keep_alive=0 dynamic VRAM release", C_GREEN_700, C_GREEN_100),
        ("4. Agentic Core", "Multi-turn session memory + MRPL SOP system prompt grounding", C_YELLOW_600, C_YELLOW_100),
        ("5. Multimodal", "DOCX, Excel telemetry, PDF, code, & P&ID diagram parsing", C_ORANGE_600, C_ORANGE_100),
        ("6. Hardening", "ReportLab PDF & Word export + offline zero-egress proof", C_PINK_600, C_PINK_100),
    ]
    card_w = Inches(1.23)
    for idx, (title, desc, col, bg_col) in enumerate(stepper_nodes):
        bx = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0 + idx * 1.3), Inches(1.65), card_w, Inches(1.3))
        bx.fill.solid(); bx.fill.fore_color.rgb = bg_col
        bx.line.color.rgb = col
        tfx = bx.text_frame
        tfx.word_wrap = True
        p1 = tfx.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(7.2); p1.font.color.rgb = col; p1.alignment = PP_ALIGN.CENTER
        p2 = tfx.add_paragraph(); p2.text = desc; p2.font.size = Pt(6.2); p2.font.color.rgb = C_SLATE_700; p2.alignment = PP_ALIGN.CENTER

    # Bottom-Center: Risk and Mitigation Evidence Table
    tx_rme = s4.shapes.add_textbox(Inches(5.0), Inches(3.05), Inches(4.8), Inches(0.35))
    p_rme = tx_rme.text_frame.paragraphs[0]
    p_rme.text = "🛡️ [MATRIX] RISK & MITIGATION EVIDENCE"
    p_rme.font.name = "Arial"
    p_rme.font.size = Pt(9.5)
    p_rme.font.bold = True
    p_rme.font.color.rgb = C_DARK_TEXT
    p_rme.alignment = PP_ALIGN.CENTER

    t_risk_shape = s4.shapes.add_table(5, 3, Inches(5.0), Inches(3.45), Inches(4.8), Inches(3.35))
    t_rk = t_risk_shape.table
    t_rk.columns[0].width = Inches(1.5)
    t_rk.columns[1].width = Inches(1.65)
    t_rk.columns[2].width = Inches(1.65)

    risk_data = [
        ("Identified Risk", "Control Implemented", "Live Demo Evidence", C_SLATE_100, C_DARK_TEXT),
        ("Model Exceeds VRAM (OOM)", "Hardware profiler + formula: Mem = Params * 0.65 + 0.8 GB", "Dynamic badges (Optimal, Full GPU, CPU) prevent OOM crashes.", C_YELLOW_100, C_YELLOW_600),
        ("VRAM Leak / Zombie Tasks", "Win32 Job Object (0x2000) + dynamic keep_alive: 0 unloader", "Verified via nvidia-smi: exactly 0 MB VRAM retained after exit.", C_ORANGE_100, C_ORANGE_600),
        ("Cloud Data Leakage", "Hardcoded loopback (127.0.0.1) with zero external calls", "Wireshark records 0 outbound packets during chat and export.", C_PINK_100, C_PINK_600),
        ("Interrupted Downloads", "NDJSON streaming reader with thread cancellation event", "Live byte progress bar with instant responsive Cancel button.", C_PINK_50, C_PINK_600),
    ]
    for r_idx, (c1, c2, c3, bg_col, txt_col) in enumerate(risk_data):
        cell1, cell2, cell3 = t_rk.cell(r_idx, 0), t_rk.cell(r_idx, 1), t_rk.cell(r_idx, 2)
        cell1.fill.solid(); cell1.fill.fore_color.rgb = bg_col
        cell2.fill.solid(); cell2.fill.fore_color.rgb = C_SLATE_50
        cell3.fill.solid(); cell3.fill.fore_color.rgb = C_WHITE
        p1 = cell1.text_frame.paragraphs[0]; p1.text = c1; p1.font.bold = True; p1.font.size = Pt(7.5); p1.font.color.rgb = txt_col
        p2 = cell2.text_frame.paragraphs[0]; p2.text = c2; p2.font.size = Pt(6.8); p2.font.color.rgb = C_SLATE_700
        p3 = cell3.text_frame.paragraphs[0]; p3.text = c3; p3.font.size = Pt(6.8); p3.font.color.rgb = C_SLATE_700

    # Bottom-Right: Validation Gates
    tx_vg = s4.shapes.add_textbox(Inches(10.1), Inches(3.05), Inches(2.8), Inches(0.35))
    p_vg = tx_vg.text_frame.paragraphs[0]
    p_vg.text = "🎯 [RADAR] VALIDATION GATES"
    p_vg.font.name = "Arial"
    p_vg.font.size = Pt(9.5)
    p_vg.font.bold = True
    p_vg.font.color.rgb = C_DARK_TEXT
    p_vg.alignment = PP_ALIGN.CENTER

    vg_items = [
        ("Gate 5: Exportable Proof", "Final Word & PDF approval deliverable on Desktop", C_GREEN_600, C_GREEN_100),
        ("Gate 4: End-to-End Agent Task", "Full SOP audit of Hydrocracker telemetry logs", C_CYAN_700, C_CYAN_100),
        ("Gate 3: Approved Action", "Gated execution & dynamic model switching", C_BLUE_700, C_BLUE_100),
        ("Gate 2: Demonstrated Routing", "Hardware memory formula scores 32 models live", C_BLUE_500, C_BLUE_50),
        ("Gate 1: Basic Local Answer", "Sub-second token generation on local CPU/GPU", C_YELLOW_600, C_YELLOW_100),
    ]
    for idx, (title, desc, col, bg_col) in enumerate(vg_items):
        v_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.1), Inches(3.45 + idx * 0.68), Inches(2.8), Inches(0.6))
        v_box.fill.solid(); v_box.fill.fore_color.rgb = bg_col
        v_box.line.color.rgb = col
        tf_vb = v_box.text_frame
        p1 = tf_vb.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(7.5); p1.font.color.rgb = col
        p2 = tf_vb.add_paragraph(); p2.text = desc; p2.font.size = Pt(6.5); p2.font.color.rgb = C_SLATE_700

    # =========================================================================
    # SLIDE 5: Impact, 5-Gear Workflow, & Stakeholder Transformation
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s5, "IMPACT AND BENEFITS", 5)

    # Left: Workflow Execution Process (5 Gears with Detailed Real-World Actions)
    tx_wep = s5.shapes.add_textbox(Inches(0.6), Inches(1.35), Inches(5.4), Inches(0.4))
    p_wep = tx_wep.text_frame.paragraphs[0]
    p_wep.text = "⚙️ [GEARS PROCESS] 5-STAGE PIPELINE"
    p_wep.font.name = "Arial"
    p_wep.font.size = Pt(10.5)
    p_wep.font.bold = True
    p_wep.font.color.rgb = C_DARK_TEXT
    p_wep.alignment = PP_ALIGN.CENTER

    gears_nodes = [
        (Inches(0.6), Inches(2.2), Inches(2.4), Inches(1.15), "1. Route Task Types", "Hardware profiler routes coding to Qwen2.5-Coder & audits to DeepSeek-R1 based on VRAM fit.", C_BLUE_700, C_BLUE_100),
        (Inches(2.1), Inches(1.8), Inches(2.4), Inches(1.15), "2. Act in Sandbox", "Python automation & telemetry scripts execute in hidden, permissioned subprocesses (CREATE_NO_WINDOW).", C_CYAN_700, C_CYAN_100),
        (Inches(1.0), Inches(4.3), Inches(2.4), Inches(1.15), "3. Read Report Data", "Extracts Hydrocracker telemetry (142.5 bar, 410°C) and P&ID valve tags (PSV-102A) from files.", C_CYAN_500, C_CYAN_100),
        (Inches(2.7), Inches(5.1), Inches(2.4), Inches(1.15), "4. Deliver Approval Note", "Exports formal MRPL compliance note directly to Desktop in styled DOCX/PDF format.", C_YELLOW_600, C_YELLOW_100),
        (Inches(3.8), Inches(3.2), Inches(2.4), Inches(1.15), "5. Prove Zero Egress", "Wireshark monitor confirms exactly 0 outbound packets in air-gapped localhost mode.", C_PURPLE_600, C_PURPLE_100),
    ]
    for left, top, width, height, title, desc, col, bg_col in gears_nodes:
        gbx = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        gbx.fill.solid(); gbx.fill.fore_color.rgb = bg_col
        gbx.line.color.rgb = col
        gbx.line.width = Pt(1.5)
        tf_g = gbx.text_frame
        tf_g.word_wrap = True
        p1 = tf_g.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(8.5); p1.font.color.rgb = col; p1.alignment = PP_ALIGN.CENTER
        p2 = tf_g.add_paragraph(); p2.text = desc; p2.font.size = Pt(6.8); p2.font.color.rgb = C_SLATE_700; p2.alignment = PP_ALIGN.CENTER

    # Right: Possible Use Case Scenario & Why It Matters
    tx_right_imp = s5.shapes.add_textbox(Inches(6.6), Inches(1.35), Inches(6.2), Inches(5.5))
    tf_ri = tx_right_imp.text_frame
    tf_ri.word_wrap = True

    p_sc_h = tf_ri.paragraphs[0]
    p_sc_h.text = "📊 [STAKEHOLDER MATRIX] BEFORE VS. AFTER"
    p_sc_h.font.name = "Arial"
    p_sc_h.font.size = Pt(11)
    p_sc_h.font.bold = True
    p_sc_h.font.color.rgb = C_DARK_TEXT
    p_sc_h.space_after = Pt(3)

    scenarios = [
        ("• Plant Engineers: ", "Manual log cross-checking (3-4 hrs) → Instant local telemetry ingestion, automated SOP checks, & evidence-backed drafts (5 secs)."),
        ("• Safety Approvers: ", "Multi-day document review cycles → Standardized, auditable compliance notes generated in seconds with signature lines."),
        ("• IT & Cybersecurity: ", "Severe cloud AI data policy exposure → 100% on-premise air-gapped execution with verifiable zero egress traffic."),
        ("• Management: ", "High recurring cloud API subscription fees → Zero recurring costs, portable USB-deployable solution across all plant units."),
    ]
    for lbl, desc in scenarios:
        p = tf_ri.add_paragraph()
        p.text = lbl + desc
        p.font.name = "Arial"
        p.font.size = Pt(8.8)
        p.font.color.rgb = C_SLATE_800
        p.space_after = Pt(2)

    p_w_h = tf_ri.add_paragraph()
    p_w_h.text = "\n💡 [VALUE PILLARS] WHY IT MATTERS"
    p_w_h.font.name = "Arial"
    p_w_h.font.size = Pt(11)
    p_w_h.font.bold = True
    p_w_h.font.color.rgb = C_DARK_TEXT
    p_w_h.space_after = Pt(3)

    why_points = [
        ("• 100% Confidentiality: ", "Protects proprietary refinery telemetry, P&IDs, tariff calculations, and code by keeping all inference on-device."),
        ("• 10x Productivity: ", "Transforms multi-hour manual reviews into an automated flow that extracts, audits, drafts, and saves final reports."),
        ("• Total Operator Control: ", "Keeps operators in charge of consequential approvals with structured audit trails, local tools, and process supervisors."),
        ("• Universal Hardware: ", "Runs seamlessly on modest 2GB VRAM GPUs (Nvidia GT 710) up to multi-GPU enterprise AI workstations."),
    ]
    for lbl, desc in why_points:
        p = tf_ri.add_paragraph()
        p.text = lbl + desc
        p.font.name = "Arial"
        p.font.size = Pt(8.5)
        p.font.color.rgb = C_SLATE_800
        p.space_after = Pt(2)

    # =========================================================================
    # SLIDE 6: Checklist, 10-Stage Strategy, and Technical References
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s6, "RESEARCH AND REFERENCES", 6)

    # Left: Demonstration Checklist
    tx_dcl = s6.shapes.add_textbox(Inches(0.4), Inches(1.35), Inches(4.5), Inches(0.4))
    p_dcl = tx_dcl.text_frame.paragraphs[0]
    p_dcl.text = "✅ [DEMO SCORECARD] 5 TESTED GATES"
    p_dcl.font.name = "Arial"
    p_dcl.font.size = Pt(10.5)
    p_dcl.font.bold = True
    p_dcl.font.color.rgb = C_DARK_TEXT

    chk_items = [
        ("Hardware Recommendation Matches Demo PC", "Detects CPU cores, RAM, and GPU VRAM via nvidia-smi to recommend optimal model."),
        ("Model Routing Switches Installed Models", "Dynamically routes tasks across 32-model catalog and unloads VRAM via keep_alive=0."),
        ("Agent Reads Scan, Consults SOP, Exports Note", "Ingests Hydrocracker telemetry, audits MAWP limits, and writes Word note to Desktop."),
        ("Sandboxed Subprocess Runs Safely", "Windows Job Object (0x2000) supervises subprocesses with CREATE_NO_WINDOW."),
        ("Strict-Mode Monitor Proves 0 Outbound Calls", "Wireshark and network telemetry confirm 100% air-gapped localhost loopback operation.")
    ]
    for idx, (title, detail) in enumerate(chk_items):
        # Green tick circle
        tk = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.4), Inches(1.9 + idx * 0.95), Inches(0.32), Inches(0.32))
        tk.fill.solid(); tk.fill.fore_color.rgb = C_GREEN_600
        tk.line.color.rgb = C_GREEN_700
        p_tk = tk.text_frame.paragraphs[0]; p_tk.text = "✓"; p_tk.font.bold = True; p_tk.font.size = Pt(8.5); p_tk.font.color.rgb = C_WHITE; p_tk.alignment = PP_ALIGN.CENTER

        # Item text
        tx_item = s6.shapes.add_textbox(Inches(0.8), Inches(1.85 + idx * 0.95), Inches(3.9), Inches(0.9))
        tf_it = tx_item.text_frame
        tf_it.word_wrap = True
        p_it1 = tf_it.paragraphs[0]
        p_it1.text = title
        p_it1.font.name = "Arial"
        p_it1.font.size = Pt(8.8)
        p_it1.font.bold = True
        p_it1.font.color.rgb = C_DARK_TEXT

        p_it2 = tf_it.add_paragraph()
        p_it2.text = detail
        p_it2.font.name = "Arial"
        p_it2.font.size = Pt(7.2)
        p_it2.font.color.rgb = C_SLATE_700

    # Right: Local Workbench Implementation Strategy (2 Rows of 5 Detailed Cards)
    tx_lwis = s6.shapes.add_textbox(Inches(5.0), Inches(1.35), Inches(7.8), Inches(0.4))
    p_lwis = tx_lwis.text_frame.paragraphs[0]
    p_lwis.text = "🗺️ [ROADMAP GRID] 10-STAGE DEPLOYMENT"
    p_lwis.font.name = "Arial"
    p_lwis.font.size = Pt(10.5)
    p_lwis.font.bold = True
    p_lwis.font.color.rgb = C_DARK_TEXT
    p_lwis.alignment = PP_ALIGN.CENTER

    row1_cards = [
        ("1. Scope (SIH26117)", "Air-gapped, multimodal, multi-model AI workbench for MRPL.", C_CYAN_700, C_CYAN_100),
        ("2. Architecture", "Portable single-file EXE with decoupled TCP server & SPA.", C_BLUE_700, C_BLUE_100),
        ("3. Embedded Engine", "Bundled Ollama daemon with isolated OLLAMA_MODELS storage.", C_BLUE_500, C_BLUE_50),
        ("4. File Ingestion", "Native python-docx, openpyxl, PDF & diagram text extractors.", C_CYAN_700, C_CYAN_100),
        ("5. Strict Air-Gap", "Hardcoded loopback 127.0.0.1 with zero outbound network calls.", C_BLUE_700, C_BLUE_100),
    ]

    row2_cards = [
        ("6. Hardware Scoring", "Memory formula scoring with dynamic visual fit badges.", C_CYAN_700, C_CYAN_100),
        ("7. 32-Model Hub", "Categorized catalog with 1-click download & live progress.", C_BLUE_700, C_BLUE_100),
        ("8. VRAM Guard", "Win32 Job Object supervisor & dynamic keep_alive=0 unloader.", C_BLUE_500, C_BLUE_50),
        ("9. Document Engine", "ReportLab PDF formatting & direct desktop Word dispatch.", C_CYAN_700, C_CYAN_100),
        ("10. Luxury SPA UI", "5 preset themes, custom palette builder, & stop response button.", C_BLUE_700, C_BLUE_100),
    ]

    c_width = Inches(1.48)
    for idx, (title, desc, col, bg_col) in enumerate(row1_cards):
        bx = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0 + idx * 1.55), Inches(1.9), c_width, Inches(1.75))
        bx.fill.solid(); bx.fill.fore_color.rgb = bg_col; bx.line.color.rgb = col
        tfx = bx.text_frame; tfx.word_wrap = True
        p1 = tfx.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(7.8); p1.font.color.rgb = col; p1.alignment = PP_ALIGN.CENTER
        p2 = tfx.add_paragraph(); p2.text = desc; p2.font.size = Pt(6.8); p2.font.color.rgb = C_SLATE_700; p2.alignment = PP_ALIGN.CENTER

    for idx, (title, desc, col, bg_col) in enumerate(row2_cards):
        bx = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0 + idx * 1.55), Inches(3.85), c_width, Inches(1.75))
        bx.fill.solid(); bx.fill.fore_color.rgb = bg_col; bx.line.color.rgb = col
        tfx = bx.text_frame; tfx.word_wrap = True
        p1 = tfx.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(7.8); p1.font.color.rgb = col; p1.alignment = PP_ALIGN.CENTER
        p2 = tfx.add_paragraph(); p2.text = desc; p2.font.size = Pt(6.8); p2.font.color.rgb = C_SLATE_700; p2.alignment = PP_ALIGN.CENTER

    # Bottom URLs Line
    tx_urls = s6.shapes.add_textbox(Inches(0.4), Inches(6.6), Inches(12.5), Inches(0.4))
    p_u = tx_urls.text_frame.paragraphs[0]
    p_u.text = "URLs: sih.gov.in/sih2026PS (SIH26117) | github.com/ollama/ollama | github.com/reportlab/reportlab | github.com/python-openxml/python-docx | github.com/qwenlm/Qwen2.5 | github.com/deepseek-ai/DeepSeek-R1"
    p_u.font.name = "Arial"
    p_u.font.size = Pt(7.8)
    p_u.font.color.rgb = C_SLATE_500
    p_u.alignment = PP_ALIGN.CENTER

    # Save
    prs.save(output_path)
    print(f"Full Charts PowerPoint Presentation successfully created at: {output_path}")

if __name__ == '__main__':
    root_dir = os.path.dirname(os.path.abspath(__file__))
    out_file = os.path.join(root_dir, "SIH2026_SIH26117_Exact_Format_Presentation.pptx")
    build_presentation(out_file)
