"""
Build Exact SIH 2026 Template PowerPoint Presentation for SIH26117
Sovereign On-Premise Agentic AI Workbench (MRPL)
Matches the exact slide-by-slide layout, shapes, diagrams, oval badges,
wheels, cylinders, and tables from the official Smart India Hackathon template.
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

def build_sih_exact_presentation(output_path):
    prs = Presentation()
    # 16:9 Widescreen (13.333" x 7.5")
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- Color Palette ---
    C_WHITE      = RGBColor(255, 255, 255)
    C_BLACK      = RGBColor(0, 0, 0)
    C_DARK_TEXT  = RGBColor(15, 23, 42)     # #0F172A
    C_SLATE_800  = RGBColor(30, 41, 59)     # #1E293B
    C_SLATE_700  = RGBColor(51, 65, 85)     # #334155
    C_SLATE_500  = RGBColor(100, 116, 139)  # #64748B
    C_SLATE_300  = RGBColor(203, 213, 225)  # #CBD5E1
    C_SLATE_100  = RGBColor(241, 245, 249)  # #F1F5F9
    C_SLATE_50   = RGBColor(248, 250, 252)  # #F8FAFC

    C_SIH_BLUE   = RGBColor(2, 132, 199)    # #0284C7 (Footer bar)
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

    C_YELLOW_500 = RGBColor(234, 179, 8)    # #EAB308
    C_YELLOW_100 = RGBColor(254, 240, 138)  # #FEF08A

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
        # Draw SIH 2026 logo badge in top right
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
        p3.text = "[ SIH ]"
        p3.font.name = "Arial"
        p3.font.size = Pt(8.5)
        p3.font.bold = True
        p3.font.color.rgb = C_GREEN_700
        p3.alignment = PP_ALIGN.CENTER

    def add_common_header_footer(slide, title_text, page_num):
        # 1. Top-Left Team Name Oval Badge
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.4), Inches(0.35), Inches(1.5), Inches(1.0))
        oval.fill.solid()
        oval.fill.fore_color.rgb = C_WHITE
        oval.line.color.rgb = C_SLATE_700
        oval.line.width = Pt(1.5)
        tf_ov = oval.text_frame
        p_ov1 = tf_ov.paragraphs[0]
        p_ov1.text = "Team"
        p_ov1.font.name = "Arial"
        p_ov1.font.size = Pt(10)
        p_ov1.font.color.rgb = C_DARK_TEXT
        p_ov1.alignment = PP_ALIGN.CENTER
        p_ov2 = tf_ov.add_paragraph()
        p_ov2.text = "Name"
        p_ov2.font.name = "Arial"
        p_ov2.font.size = Pt(10)
        p_ov2.font.color.rgb = C_DARK_TEXT
        p_ov2.alignment = PP_ALIGN.CENTER

        # 2. Top-Center Title (Bold Serif Caps)
        tx_title = slide.shapes.add_textbox(Inches(2.2), Inches(0.45), Inches(8.8), Inches(0.8))
        tf_t = tx_title.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.name = "Georgia"
        p_t.font.size = Pt(22)
        p_t.font.bold = True
        p_t.font.color.rgb = C_DARK_TEXT
        p_t.alignment = PP_ALIGN.CENTER

        # 3. Top-Right SIH Logo
        draw_sih_logo(slide, Inches(10.7), Inches(0.35))

        # 4. Bottom Blue Footer Banner
        ft = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4))
        ft.fill.solid()
        ft.fill.fore_color.rgb = C_SIH_BLUE
        ft.line.color.rgb = C_SIH_BLUE

        tx_ft = slide.shapes.add_textbox(Inches(0.5), Inches(7.12), Inches(12.333), Inches(0.35))
        tf_ft = tx_ft.text_frame
        p_ft = tf_ft.paragraphs[0]
        p_ft.text = f"@SIH Idea submission- Template                                                                                                                                                 {page_num}"
        p_ft.font.name = "Arial"
        p_ft.font.size = Pt(10)
        p_ft.font.color.rgb = C_WHITE

    # =========================================================================
    # SLIDE 1: Title Slide
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)

    # Top Center Title
    tx_s1_t = s1.shapes.add_textbox(Inches(1.5), Inches(0.6), Inches(8.5), Inches(1.4))
    tf_s1_t = tx_s1_t.text_frame
    p_t1 = tf_s1_t.paragraphs[0]
    p_t1.text = "SMART INDIA HACKATHON 2026"
    p_t1.font.name = "Georgia"
    p_t1.font.size = Pt(22)
    p_t1.font.bold = True
    p_t1.font.color.rgb = C_DARK_TEXT
    p_t1.alignment = PP_ALIGN.CENTER

    p_t2 = tf_s1_t.add_paragraph()
    p_t2.text = "SOVEREIGN ON-PREMISE\nAGENTIC AI WORKBENCH"
    p_t2.font.name = "Georgia"
    p_t2.font.size = Pt(16)
    p_t2.font.bold = True
    p_t2.font.color.rgb = C_SIH_BLUE
    p_t2.alignment = PP_ALIGN.CENTER

    draw_sih_logo(s1, Inches(10.7), Inches(0.4))

    # Left Column: Metadata Bullets
    tx_meta = s1.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(6.8), Inches(4.8))
    tf_m = tx_meta.text_frame
    tf_m.word_wrap = True

    bullets = [
        ("• Problem Statement ID: ", "SIH26117", True),
        ("• Problem Statement Title: ", "Sovereign On-Premise Agentic AI Workbench using Open-Weight Multimodal LLMs for Confidential Industrial Work", False),
        ("• Organization: ", "Mangalore Refinery and Petrochemicals Limited (MRPL)", True),
        ("• Theme: ", "Smart Automation", False),
        ("• PS Category: ", "Software", False),
        ("• Team ID: ", "[To be updated / Registered Team ID]", False),
        ("• Team Name: ", "[Registered Team Name]", False),
    ]

    for idx, (label, val, highlight) in enumerate(bullets):
        p = tf_m.paragraphs[0] if idx == 0 else tf_m.add_paragraph()
        p.text = label + val
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.color.rgb = C_DARK_TEXT
        p.font.bold = True if "ID" in label or "Organization" in label else False
        p.space_after = Pt(14)

    # Right Graphic: Lightbulb & Brain Graphic Container
    bulb_bg = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.0), Inches(2.1), Inches(4.6), Inches(4.8))
    bulb_bg.fill.solid()
    bulb_bg.fill.fore_color.rgb = C_SLATE_50
    bulb_bg.line.color.rgb = C_SLATE_300
    bulb_bg.line.width = Pt(1.5)

    tf_bg = bulb_bg.text_frame
    tf_bg.word_wrap = True
    p_bg1 = tf_bg.paragraphs[0]
    p_bg1.text = "💡 AIR-GAPPED ON-PREMISE AI"
    p_bg1.font.name = "Arial"
    p_bg1.font.size = Pt(13)
    p_bg1.font.bold = True
    p_bg1.font.color.rgb = C_ORANGE_600
    p_bg1.alignment = PP_ALIGN.CENTER

    p_bg2 = tf_bg.add_paragraph()
    p_bg2.text = "\n[ Circuit / Neural Brain Architecture ]\n• 100% Air-Gapped Local Workstation\n• Embedded Ollama (Port 11434)\n• Zero Egress & Zero Cloud Dependency\n• Real-Time GPU VRAM Profiling\n• 32 Curated Open-Weight Models\n• Windows Job Object Crash Safety"
    p_bg2.font.name = "Arial"
    p_bg2.font.size = Pt(10)
    p_bg2.font.color.rgb = C_SLATE_700
    p_bg2.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 2: Problem & Proposed Solution + Wheel + Cylinder
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s2, "SOVEREIGN AI WORKBENCH", 2)

    # Left Column: Problem & Proposed Solution
    tx_p = s2.shapes.add_textbox(Inches(0.4), Inches(1.5), Inches(3.6), Inches(2.6))
    tf_p = tx_p.text_frame
    tf_p.word_wrap = True
    p_ph = tf_p.paragraphs[0]
    p_ph.text = "THE PROBLEM"
    p_ph.font.name = "Arial"
    p_ph.font.size = Pt(13)
    p_ph.font.bold = True
    p_ph.font.color.rgb = C_DARK_TEXT
    
    p_pd = tf_p.add_paragraph()
    p_pd.text = "Sensitive approval notes, P&IDs, inspection reports, calculations and internal code cannot leave refinery, PSU, defence or government premises. Cloud assistants create a confidentiality gap, while manual work reduces productivity."
    p_pd.font.name = "Arial"
    p_pd.font.size = Pt(9.5)
    p_pd.font.color.rgb = C_SLATE_700

    # Mini badge: Data Confidentiality Gap
    bg_p = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(3.45), Inches(3.4), Inches(0.65))
    bg_p.fill.solid()
    bg_p.fill.fore_color.rgb = C_PINK_50
    bg_p.line.color.rgb = C_PINK_500
    tf_bgp = bg_p.text_frame
    p_bgp1 = tf_bgp.paragraphs[0]
    p_bgp1.text = "Data Confidentiality Gap"
    p_bgp1.font.bold = True
    p_bgp1.font.size = Pt(9)
    p_bgp1.font.color.rgb = C_PINK_500
    p_bgp2 = tf_bgp.add_paragraph()
    p_bgp2.text = "Sensitive data cannot leave premises"
    p_bgp2.font.size = Pt(8)
    p_bgp2.font.color.rgb = C_SLATE_700

    # Solution
    tx_s = s2.shapes.add_textbox(Inches(0.4), Inches(4.2), Inches(3.6), Inches(2.2))
    tf_s = tx_s.text_frame
    tf_s.word_wrap = True
    p_sh = tf_s.paragraphs[0]
    p_sh.text = "THE PROPOSED SOLUTION"
    p_sh.font.name = "Arial"
    p_sh.font.size = Pt(13)
    p_sh.font.bold = True
    p_sh.font.color.rgb = C_DARK_TEXT

    p_sd = tf_s.add_paragraph()
    p_sd.text = "A self-hosted AI workbench that runs on the organisation's own workstation or GPU server. It plans multi-step work, calls local tools, selects among installed open-weight models, understands scans and drawings, and delivers usable files with a verifiable zero-external-call mode."
    p_sd.font.name = "Arial"
    p_sd.font.size = Pt(9.5)
    p_sd.font.color.rgb = C_SLATE_700

    # Mini badge: Secure Data Processing
    bg_s = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(6.3), Inches(3.4), Inches(0.65))
    bg_s.fill.solid()
    bg_s.fill.fore_color.rgb = C_BLUE_50
    bg_s.line.color.rgb = C_BLUE_500
    tf_bgs = bg_s.text_frame
    p_bgs1 = tf_bgs.paragraphs[0]
    p_bgs1.text = "Secure Data Processing"
    p_bgs1.font.bold = True
    p_bgs1.font.size = Pt(9)
    p_bgs1.font.color.rgb = C_BLUE_700
    p_bgs2 = tf_bgs.add_paragraph()
    p_bgs2.text = "Usable files with zero external calls"
    p_bgs2.font.size = Pt(8)
    p_bgs2.font.color.rgb = C_SLATE_700

    # Center: Circular Wheel Container & 6 Sectors
    wheel_title = s2.shapes.add_textbox(Inches(4.1), Inches(1.35), Inches(4.8), Inches(0.4))
    p_wt = wheel_title.text_frame.paragraphs[0]
    p_wt.text = "Secure AI Workbench for Sensitive Data"
    p_wt.font.name = "Arial"
    p_wt.font.size = Pt(10.5)
    p_wt.font.bold = True
    p_wt.font.color.rgb = C_DARK_TEXT
    p_wt.alignment = PP_ALIGN.CENTER

    # Wheel Center Circle
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

    # 6 Surrounding Pillar Cards in Circle
    pillar_nodes = [
        (Inches(5.3), Inches(1.8), Inches(2.4), Inches(0.7), "Multi-Step Work Planning", "AI plans complex tasks", C_ORANGE_600, C_ORANGE_100),
        (Inches(7.2), Inches(2.6), Inches(2.2), Inches(0.7), "Local Tool Integration", "Calls tools on the server", C_YELLOW_500, C_YELLOW_100),
        (Inches(7.2), Inches(4.7), Inches(2.2), Inches(0.7), "Open-Weight Selection", "Chooses from installed models", C_GREEN_600, C_GREEN_100),
        (Inches(5.3), Inches(5.6), Inches(2.4), Inches(0.7), "Scan & Drawing Understanding", "AI interprets visual info", C_CYAN_700, C_CYAN_100),
        (Inches(3.8), Inches(4.7), Inches(2.0), Inches(0.7), "Verifiable Zero-External", "Ensures data stays internal", C_BLUE_700, C_BLUE_100),
        (Inches(3.8), Inches(2.6), Inches(2.0), Inches(0.7), "Confidentiality Engine", "Protects proprietary code", C_PINK_500, C_PINK_100),
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
        p_c1.font.size = Pt(8)
        p_c1.font.color.rgb = col
        p_c1.alignment = PP_ALIGN.CENTER
        p_c2 = tf_c.add_paragraph()
        p_c2.text = desc
        p_c2.font.size = Pt(7)
        p_c2.font.color.rgb = C_SLATE_700
        p_c2.alignment = PP_ALIGN.CENTER

    # Right Column: Stacked Capabilities Cylinder
    cyl_title = s2.shapes.add_textbox(Inches(9.4), Inches(1.3), Inches(3.6), Inches(0.6))
    tf_cyl = cyl_title.text_frame
    tf_cyl.word_wrap = True
    p_cyl_t = tf_cyl.paragraphs[0]
    p_cyl_t.text = "System capabilities range from device-specific to broad knowledge."
    p_cyl_t.font.name = "Arial"
    p_cyl_t.font.size = Pt(8.5)
    p_cyl_t.font.color.rgb = C_DARK_TEXT
    p_cyl_t.alignment = PP_ALIGN.CENTER

    # 5-Tier Cylinder Layer Blocks
    layers = [
        ("Multimodal Deliverables", "Exports various file types and calculations", C_GREEN_600, C_GREEN_100),
        ("Local Knowledge", "Searches manuals and retains context", C_BLUE_700, C_BLUE_100),
        ("Permission-Gated", "Asks before file writes or execution", C_PINK_500, C_PINK_100),
        ("Automatic Routing", "Routes tasks to suitable models", C_ORANGE_600, C_ORANGE_100),
        ("Model Recommendation", "Recommends models fitting device", C_YELLOW_500, C_YELLOW_100),
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
        pl1.font.size = Pt(9)
        pl1.font.color.rgb = col

        pl2 = tf_l.add_paragraph()
        pl2.text = desc
        pl2.font.size = Pt(7.5)
        pl2.font.color.rgb = C_SLATE_700

    # =========================================================================
    # SLIDE 3: Technical Approach
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s3, "TECHNICAL APPROACH", 3)

    # Left: Building a Comprehensive Local System
    t_left = s3.shapes.add_textbox(Inches(0.6), Inches(1.35), Inches(5.8), Inches(0.4))
    p_tl = t_left.text_frame.paragraphs[0]
    p_tl.text = "Building a Comprehensive Local System"
    p_tl.font.name = "Arial"
    p_tl.font.size = Pt(11)
    p_tl.font.bold = True
    p_tl.font.color.rgb = C_DARK_TEXT
    p_tl.alignment = PP_ALIGN.CENTER

    left_nodes = [
        ("Client", "User interface for interaction and task initiation.", C_CYAN_700, C_CYAN_100),
        ("Orchestration", "Manages task execution and resource allocation.", C_GREEN_700, C_GREEN_100),
        ("Guardrail", "Enforces security and access control policies.", C_BLUE_700, C_BLUE_100),
        ("Tools + Knowledge", "Provides essential resources for task completion.", C_YELLOW_500, C_YELLOW_100),
        ("Model + Multimodal", "Processes information and handles multimodal inputs.", C_ORANGE_600, C_ORANGE_100),
        ("Output + Network Boundary", "Generates outputs and controls network access.", C_PINK_500, C_PINK_100),
    ]
    for idx, (title, desc, col, bg_col) in enumerate(left_nodes):
        node_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.8 + idx * 0.65), Inches(3.8), Inches(0.55))
        node_box.fill.solid()
        node_box.fill.fore_color.rgb = bg_col
        node_box.line.color.rgb = col
        tf_nb = node_box.text_frame
        p1 = tf_nb.paragraphs[0]
        p1.text = title + ": "
        p1.font.bold = True
        p1.font.size = Pt(8.5)
        p1.font.color.rgb = col
        p2 = tf_nb.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(7.5)
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
    ph1_d.text = "Seamless integration of diverse components for efficient operation."
    ph1_d.font.size = Pt(6.5)
    ph1_d.font.color.rgb = C_SLATE_700
    ph1_d.alignment = PP_ALIGN.CENTER

    # Right: Streamlined Inspection Process
    t_right = s3.shapes.add_textbox(Inches(6.8), Inches(1.35), Inches(6.0), Inches(0.4))
    p_tr = t_right.text_frame.paragraphs[0]
    p_tr.text = "Streamlined Inspection Process"
    p_tr.font.name = "Arial"
    p_tr.font.size = Pt(11)
    p_tr.font.bold = True
    p_tr.font.color.rgb = C_DARK_TEXT
    p_tr.alignment = PP_ALIGN.CENTER

    right_steps = [
        ("Report Upload", "Upload scanned inspection report and SOPs.", C_GREEN_700, C_GREEN_100),
        ("Data Extraction", "OCR extracts findings; RAG retrieves matching procedures.", C_YELLOW_500, C_YELLOW_100),
        ("Document Routing", "Router selects appropriate document model.", C_CYAN_700, C_CYAN_100),
        ("Agent Drafting", "Agent drafts finding summary and approval note.", C_PURPLE_600, C_PURPLE_100),
        ("Guardrail Approval", "Guardrail requests approval for writing and attaching evidence.", C_PINK_500, C_PINK_100),
        ("Network Monitoring", "Network monitor confirms zero external calls.", C_ORANGE_600, C_ORANGE_100),
    ]
    for idx, (title, desc, col, bg_col) in enumerate(right_steps):
        step_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8 + idx * 0.65), Inches(3.8), Inches(0.55))
        step_box.fill.solid()
        step_box.fill.fore_color.rgb = bg_col
        step_box.line.color.rgb = col
        tf_sb = step_box.text_frame
        p1 = tf_sb.paragraphs[0]
        p1.text = title + ": "
        p1.font.bold = True
        p1.font.size = Pt(8.5)
        p1.font.color.rgb = col
        p2 = tf_sb.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(7.5)
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
    ph2_d.text = "Detailed report with findings, procedures, and audit trail."
    ph2_d.font.size = Pt(7)
    ph2_d.font.color.rgb = C_SLATE_700
    ph2_d.alignment = PP_ALIGN.CENTER

    # Bottom Banner: Implementation Components
    tx_ic = s3.shapes.add_textbox(Inches(0.6), Inches(5.9), Inches(12.133), Inches(1.0))
    tf_ic = tx_ic.text_frame
    tf_ic.word_wrap = True
    pic_h = tf_ic.paragraphs[0]
    pic_h.text = "IMPLEMENTATION COMPONENTS"
    pic_h.font.name = "Arial"
    pic_h.font.size = Pt(11)
    pic_h.font.bold = True
    pic_h.font.color.rgb = C_DARK_TEXT
    pic_h.alignment = PP_ALIGN.CENTER

    pic_d = tf_ic.add_paragraph()
    pic_d.text = "Agent harness: OpenCode style plan-act loop | Inference: llama.cpp + GGUF open-weight models | OCR: locally deployed PaddleOCR | Retrieval: local embeddings + vector index | Isolation: Docker / OS firewall / traffic logging"
    pic_d.font.name = "Arial"
    pic_d.font.size = Pt(9.5)
    pic_d.font.color.rgb = C_SLATE_800
    pic_d.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 4: Feasibility and Viability
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s4, "FEASIBILITY AND VIABILITY", 4)

    # Top-Left Title: Feasibility on a Single Workstation
    tx_f_t = s4.shapes.add_textbox(Inches(0.4), Inches(1.3), Inches(4.5), Inches(0.35))
    p_ft = tx_f_t.text_frame.paragraphs[0]
    p_ft.text = "Feasibility on a Single Workstation or GPU Server"
    p_ft.font.name = "Arial"
    p_ft.font.size = Pt(9.5)
    p_ft.font.bold = True
    p_ft.font.color.rgb = C_DARK_TEXT

    # Feasibility Table
    t_feas_shape = s4.shapes.add_table(5, 2, Inches(0.4), Inches(1.7), Inches(4.4), Inches(5.1))
    tf_tab = t_feas_shape.table
    tf_tab.columns[0].width = Inches(1.6)
    tf_tab.columns[1].width = Inches(2.8)

    feas_rows = [
        ("Feature", "Description", C_SLATE_100, C_DARK_TEXT, True),
        ("Open-weight models", "Avoid dependence on cloud APIs", C_YELLOW_100, C_YELLOW_500, False),
        ("Quantized GGUF models", "Match available VRAM; fallback to smaller models when venue hardware cannot host a larger model", C_ORANGE_100, C_ORANGE_600, False),
        ("llama.cpp", "Exposes local serving and supports CPU, NVIDIA CUDA, AMD HIP, Vulkan and CPU+GPU hybrid inference", C_PINK_100, C_PINK_500, False),
        ("Data storage", "All document indices, logs and generated artifacts remain on local storage", C_PINK_50, C_PINK_500, False),
    ]
    for r_idx, (c1, c2, bg_col, txt_col, is_head) in enumerate(feas_rows):
        cell1, cell2 = tf_tab.cell(r_idx, 0), tf_tab.cell(r_idx, 1)
        cell1.fill.solid(); cell1.fill.fore_color.rgb = bg_col
        cell2.fill.solid(); cell2.fill.fore_color.rgb = C_WHITE if is_head else C_SLATE_50
        p1 = cell1.text_frame.paragraphs[0]; p1.text = c1; p1.font.bold = True; p1.font.size = Pt(8.5); p1.font.color.rgb = txt_col
        p2 = cell2.text_frame.paragraphs[0]; p2.text = c2; p2.font.bold = is_head; p2.font.size = Pt(8); p2.font.color.rgb = C_SLATE_700

    # Top-Right: Build an Agentic System Stepper (6 cards)
    tx_stepper = s4.shapes.add_textbox(Inches(5.0), Inches(1.3), Inches(7.8), Inches(0.35))
    p_stp = tx_stepper.text_frame.paragraphs[0]
    p_stp.text = "Build an Agentic System"
    p_stp.font.name = "Arial"
    p_stp.font.size = Pt(10.5)
    p_stp.font.bold = True
    p_stp.font.color.rgb = C_DARK_TEXT
    p_stp.alignment = PP_ALIGN.CENTER

    stepper_nodes = [
        ("Establish Foundation", "OpenCode harness + llama.cpp + one local model", C_BLUE_700, C_BLUE_100),
        ("Implement Model Layer", "Hardware profiler + task router + second model", C_CYAN_700, C_CYAN_100),
        ("Apply Guardrails", "Permission interceptor and action toggle", C_GREEN_700, C_GREEN_100),
        ("Develop Agentic Core", "Memory, skills and multi-step execution", C_YELLOW_500, C_YELLOW_100),
        ("Integrate Multimodal", "OCR/vision and local document grounding", C_ORANGE_600, C_ORANGE_100),
        ("Perform Hardening", "Deliverable generation + offline proof", C_PINK_500, C_PINK_100),
    ]
    card_w = Inches(1.23)
    for idx, (title, desc, col, bg_col) in enumerate(stepper_nodes):
        bx = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0 + idx * 1.3), Inches(1.7), card_w, Inches(1.25))
        bx.fill.solid(); bx.fill.fore_color.rgb = bg_col
        bx.line.color.rgb = col
        tfx = bx.text_frame
        tfx.word_wrap = True
        p1 = tfx.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(7.5); p1.font.color.rgb = col; p1.alignment = PP_ALIGN.CENTER
        p2 = tfx.add_paragraph(); p2.text = desc; p2.font.size = Pt(6.5); p2.font.color.rgb = C_SLATE_700; p2.alignment = PP_ALIGN.CENTER

    # Bottom-Center: Risk and Mitigation Evidence Table
    tx_rme = s4.shapes.add_textbox(Inches(5.0), Inches(3.05), Inches(4.8), Inches(0.35))
    p_rme = tx_rme.text_frame.paragraphs[0]
    p_rme.text = "Risk and Mitigation Evidence"
    p_rme.font.name = "Arial"
    p_rme.font.size = Pt(10)
    p_rme.font.bold = True
    p_rme.font.color.rgb = C_DARK_TEXT
    p_rme.alignment = PP_ALIGN.CENTER

    t_risk_shape = s4.shapes.add_table(5, 3, Inches(5.0), Inches(3.45), Inches(4.8), Inches(3.35))
    t_rk = t_risk_shape.table
    t_rk.columns[0].width = Inches(1.6)
    t_rk.columns[1].width = Inches(1.6)
    t_rk.columns[2].width = Inches(1.6)

    risk_data = [
        ("Risk", "Control", "Demo evidence", C_SLATE_100, C_DARK_TEXT),
        ("Model too large", "Hardware profiler + quantized fallback", "Two installed models route live", C_YELLOW_100, C_YELLOW_500),
        ("Unsafe tool action", "Permission gate + sandbox", "Prompt shown before write/execute", C_ORANGE_100, C_ORANGE_600),
        ("OCR error", "Preview, confidence check, human review", "Source page and extracted text", C_PINK_100, C_PINK_500),
        ("Unexpected outbound call", "Offline config + firewall + monitor", "Traffic log remains empty", C_PINK_50, C_PINK_500),
    ]
    for r_idx, (c1, c2, c3, bg_col, txt_col) in enumerate(risk_data):
        cell1, cell2, cell3 = t_rk.cell(r_idx, 0), t_rk.cell(r_idx, 1), t_rk.cell(r_idx, 2)
        cell1.fill.solid(); cell1.fill.fore_color.rgb = bg_col
        cell2.fill.solid(); cell2.fill.fore_color.rgb = C_SLATE_50
        cell3.fill.solid(); cell3.fill.fore_color.rgb = C_WHITE
        p1 = cell1.text_frame.paragraphs[0]; p1.text = c1; p1.font.bold = True; p1.font.size = Pt(7.5); p1.font.color.rgb = txt_col
        p2 = cell2.text_frame.paragraphs[0]; p2.text = c2; p2.font.size = Pt(7); p2.font.color.rgb = C_SLATE_700
        p3 = cell3.text_frame.paragraphs[0]; p3.text = c3; p3.font.size = Pt(7); p3.font.color.rgb = C_SLATE_700

    # Bottom-Right: Validation Gates (Concentric Rings / List)
    tx_vg = s4.shapes.add_textbox(Inches(10.1), Inches(3.05), Inches(2.8), Inches(0.35))
    p_vg = tx_vg.text_frame.paragraphs[0]
    p_vg.text = "Validation Gates"
    p_vg.font.name = "Arial"
    p_vg.font.size = Pt(10)
    p_vg.font.bold = True
    p_vg.font.color.rgb = C_DARK_TEXT
    p_vg.alignment = PP_ALIGN.CENTER

    vg_items = [
        ("Exportable Artifact & Proof", "Final deliverable for review", C_GREEN_600, C_GREEN_100),
        ("End-to-End Agent Task", "Comprehensive task completion", C_CYAN_700, C_CYAN_100),
        ("Approved Action", "Decision-making and implementation", C_BLUE_700, C_BLUE_100),
        ("Demonstrated Routing", "Navigation and connectivity", C_BLUE_500, C_BLUE_50),
        ("Basic Local Answer", "Initial functionality", C_YELLOW_500, C_YELLOW_100),
    ]
    for idx, (title, desc, col, bg_col) in enumerate(vg_items):
        v_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.1), Inches(3.45 + idx * 0.68), Inches(2.8), Inches(0.6))
        v_box.fill.solid(); v_box.fill.fore_color.rgb = bg_col
        v_box.line.color.rgb = col
        tf_vb = v_box.text_frame
        p1 = tf_vb.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(7.5); p1.font.color.rgb = col
        p2 = tf_vb.add_paragraph(); p2.text = desc; p2.font.size = Pt(6.5); p2.font.color.rgb = C_SLATE_700

    # =========================================================================
    # SLIDE 5: Impact and Benefits
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s5, "IMPACT AND BENEFITS", 5)

    # Left: Workflow Execution Process (5 Gears)
    tx_wep = s5.shapes.add_textbox(Inches(0.6), Inches(1.4), Inches(5.4), Inches(0.4))
    p_wep = tx_wep.text_frame.paragraphs[0]
    p_wep.text = "Workflow Execution Process"
    p_wep.font.name = "Arial"
    p_wep.font.size = Pt(12)
    p_wep.font.bold = True
    p_wep.font.color.rgb = C_DARK_TEXT
    p_wep.alignment = PP_ALIGN.CENTER

    gears_nodes = [
        (Inches(0.6), Inches(2.3), Inches(2.4), Inches(1.1), "Route task types", "Two task types use different installed models.", C_BLUE_700, C_BLUE_100),
        (Inches(2.0), Inches(1.9), Inches(2.4), Inches(1.1), "Act in sandbox", "One coding task executes in a sandbox after permission approval.", C_CYAN_700, C_CYAN_100),
        (Inches(1.0), Inches(4.3), Inches(2.4), Inches(1.1), "Read report data", "One scanned report or drawing produces verified extracted findings.", C_CYAN_500, C_CYAN_500),
        (Inches(2.6), Inches(5.1), Inches(2.4), Inches(1.1), "Deliver approval note", "One Word approval note or equivalent artifact is exported locally.", C_YELLOW_500, C_YELLOW_100),
        (Inches(3.8), Inches(3.2), Inches(2.4), Inches(1.1), "Prove network security", "Network log or monitor shows no external call in strict mode.", C_PURPLE_600, C_PURPLE_100),
    ]
    for left, top, width, height, title, desc, col, bg_col in gears_nodes:
        gbx = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        gbx.fill.solid(); gbx.fill.fore_color.rgb = bg_col if bg_col != C_CYAN_500 else C_CYAN_100
        gbx.line.color.rgb = col
        gbx.line.width = Pt(1.5)
        tf_g = gbx.text_frame
        tf_g.word_wrap = True
        p1 = tf_g.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(9); p1.font.color.rgb = col; p1.alignment = PP_ALIGN.CENTER
        p2 = tf_g.add_paragraph(); p2.text = desc; p2.font.size = Pt(7.5); p2.font.color.rgb = C_SLATE_700; p2.alignment = PP_ALIGN.CENTER

    # Right: Possible Use Case Scenario & Why It Matters
    tx_right_imp = s5.shapes.add_textbox(Inches(6.6), Inches(1.4), Inches(6.2), Inches(5.4))
    tf_ri = tx_right_imp.text_frame
    tf_ri.word_wrap = True

    p_sc_h = tf_ri.paragraphs[0]
    p_sc_h.text = "POSSIBLE USE CASE SCENARIO"
    p_sc_h.font.name = "Arial"
    p_sc_h.font.size = Pt(13)
    p_sc_h.font.bold = True
    p_sc_h.font.color.rgb = C_DARK_TEXT
    p_sc_h.space_after = Pt(6)

    scenarios = [
        ("• Engineers: ", "Manual handling → Local scan understanding and evidence-backed drafts."),
        ("• Approvers: ", "Slow document reviews → Structured, traceable approval notes."),
        ("• IT / Security: ", "Cloud AI policy exposure → On-prem models with permissions and network evidence."),
        ("• Management: ", "Limited confidential-work automation → Reusable skills, artifacts, and auditability."),
    ]
    for lbl, desc in scenarios:
        p = tf_ri.add_paragraph()
        p.text = lbl + desc
        p.font.name = "Arial"
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_SLATE_800
        p.space_after = Pt(4)

    p_w_h = tf_ri.add_paragraph()
    p_w_h.text = "\nWHY IT MATTERS"
    p_w_h.font.name = "Arial"
    p_w_h.font.size = Pt(13)
    p_w_h.font.bold = True
    p_w_h.font.color.rgb = C_DARK_TEXT
    p_w_h.space_after = Pt(6)

    why_points = [
        ("• Confidentiality: ", "Protects proprietary drawings, financials, negotiations, correspondence, and designs by keeping prompts, models, and indexes on premises."),
        ("• Productivity: ", "Transforms a document review into an agent-assisted flow that extracts, retrieves, drafts, requests approval, and writes a usable deliverable."),
        ("• Control: ", "Keeps people in charge of consequential actions and gives security teams a visible trail of permissions, local tools, and network events."),
        ("• Adaptability: ", "Adds new open-weight models without redesigning the workbench; the router can select among models already installed on the device."),
    ]
    for lbl, desc in why_points:
        p = tf_ri.add_paragraph()
        p.text = lbl + desc
        p.font.name = "Arial"
        p.font.size = Pt(9)
        p.font.color.rgb = C_SLATE_800
        p.space_after = Pt(4)

    # =========================================================================
    # SLIDE 6: Research, References & Implementation Strategy
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_common_header_footer(s6, "RESEARCH AND REFERENCES", 6)

    # Left: Demonstration Checklist
    tx_dcl = s6.shapes.add_textbox(Inches(0.4), Inches(1.4), Inches(4.5), Inches(0.4))
    p_dcl = tx_dcl.text_frame.paragraphs[0]
    p_dcl.text = "Demonstration Checklist"
    p_dcl.font.name = "Arial"
    p_dcl.font.size = Pt(12)
    p_dcl.font.bold = True
    p_dcl.font.color.rgb = C_DARK_TEXT

    chk_items = [
        "Hardware recommendation matches the demo machine",
        "Sandboxed coding task runs after a permission decision",
        "Model routing switches between at least two installed models",
        "Strict-mode log / monitor records no external calls",
        "Agent reads a scan, consults local knowledge and exports an approval note"
    ]
    for idx, item in enumerate(chk_items):
        # Green tick circle
        tk = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.4), Inches(2.0 + idx * 0.95), Inches(0.35), Inches(0.35))
        tk.fill.solid(); tk.fill.fore_color.rgb = C_GREEN_600
        tk.line.color.rgb = C_GREEN_700
        p_tk = tk.text_frame.paragraphs[0]; p_tk.text = "✓"; p_tk.font.bold = True; p_tk.font.size = Pt(9); p_tk.font.color.rgb = C_WHITE; p_tk.alignment = PP_ALIGN.CENTER

        # Item card text
        tx_item = s6.shapes.add_textbox(Inches(0.85), Inches(1.95 + idx * 0.95), Inches(3.8), Inches(0.85))
        tf_it = tx_item.text_frame
        tf_it.word_wrap = True
        p_it = tf_it.paragraphs[0]
        p_it.text = item
        p_it.font.name = "Arial"
        p_it.font.size = Pt(9.5)
        p_it.font.color.rgb = C_SLATE_800

    # Right: Local Workbench Implementation Strategy (2 Rows of 5 Cards)
    tx_lwis = s6.shapes.add_textbox(Inches(5.0), Inches(1.4), Inches(7.8), Inches(0.4))
    p_lwis = tx_lwis.text_frame.paragraphs[0]
    p_lwis.text = "Local Workbench Implementation Strategy"
    p_lwis.font.name = "Arial"
    p_lwis.font.size = Pt(12)
    p_lwis.font.bold = True
    p_lwis.font.color.rgb = C_DARK_TEXT
    p_lwis.alignment = PP_ALIGN.CENTER

    row1_cards = [
        ("SIH 2026 PS 26117MRPL", "Define official scope for air-gapped, multi-model, agentic, multimodal local workbench.", C_CYAN_700, C_CYAN_100),
        ("Project Context & Design", "Define feature set, architecture, phased methodology, and target users.", C_BLUE_700, C_BLUE_100),
        ("llama.cpp / ggml-org", "Implement local LLM/VLM inference with quantization and acceleration.", C_BLUE_500, C_BLUE_50),
        ("PaddleOCR / Paddle", "Integrate open-source OCR for document processing of PDFs & images.", C_CYAN_700, C_CYAN_100),
        ("Open WebUI Offline", "Ensure offline operation via preloaded models & strict egress blocking.", C_BLUE_700, C_BLUE_100),
    ]

    row2_cards = [
        ("Build Success Criteria", "Build against exact success criteria rather than generic chatbot models.", C_CYAN_700, C_CYAN_100),
        ("Extensible Workbench", "Use an extensible local workbench instead of hardcoded flows.", C_BLUE_700, C_BLUE_100),
        ("Local Model Registry", "Serve open-weight GGUF models through a local registry/router.", C_BLUE_500, C_BLUE_50),
        ("Local Vision Reasoning", "Run OCR/vision locally before agent reasoning on scans.", C_CYAN_700, C_CYAN_100),
        ("Strict-Mode Proof", "Preload assets, block egress, and log traffic for proof.", C_BLUE_700, C_BLUE_100),
    ]

    c_width = Inches(1.48)
    for idx, (title, desc, col, bg_col) in enumerate(row1_cards):
        bx = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0 + idx * 1.55), Inches(2.0), c_width, Inches(1.7))
        bx.fill.solid(); bx.fill.fore_color.rgb = bg_col; bx.line.color.rgb = col
        tfx = bx.text_frame; tfx.word_wrap = True
        p1 = tfx.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(8); p1.font.color.rgb = col; p1.alignment = PP_ALIGN.CENTER
        p2 = tfx.add_paragraph(); p2.text = desc; p2.font.size = Pt(7); p2.font.color.rgb = C_SLATE_700; p2.alignment = PP_ALIGN.CENTER

    for idx, (title, desc, col, bg_col) in enumerate(row2_cards):
        bx = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0 + idx * 1.55), Inches(3.9), c_width, Inches(1.7))
        bx.fill.solid(); bx.fill.fore_color.rgb = bg_col; bx.line.color.rgb = col
        tfx = bx.text_frame; tfx.word_wrap = True
        p1 = tfx.paragraphs[0]; p1.text = title; p1.font.bold = True; p1.font.size = Pt(8); p1.font.color.rgb = col; p1.alignment = PP_ALIGN.CENTER
        p2 = tfx.add_paragraph(); p2.text = desc; p2.font.size = Pt(7); p2.font.color.rgb = C_SLATE_700; p2.alignment = PP_ALIGN.CENTER

    # Bottom URLs Line
    tx_urls = s6.shapes.add_textbox(Inches(0.4), Inches(6.6), Inches(12.5), Inches(0.4))
    p_u = tx_urls.text_frame.paragraphs[0]
    p_u.text = "URLs: sih.gov.in/sih2026PS | github.com/ggml-org/llama.cpp | github.com/PaddlePaddle/PaddleOCR | github.com/open-webui/docs/blob/main/docs/tutorials/maintenance/offline-mode.mdx"
    p_u.font.name = "Arial"
    p_u.font.size = Pt(8)
    p_u.font.color.rgb = C_SLATE_500
    p_u.alignment = PP_ALIGN.CENTER

    # Save
    prs.save(output_path)
    print(f"Exact SIH PowerPoint Presentation successfully created at: {output_path}")

if __name__ == '__main__':
    root_dir = os.path.dirname(os.path.abspath(__file__))
    out_file = os.path.join(root_dir, "SIH2026_SIH26117_Exact_Format_Presentation.pptx")
    build_sih_exact_presentation(out_file)
