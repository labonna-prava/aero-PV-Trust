"""
Script to generate an enhanced, presentation-ready 10-slide PowerPoint presentation (.pptx)
designed for simultaneous live UI sharing and defense presentation.
Author: Labonna Prava
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = Presentation()
    # 16:9 widescreen layout
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette
    C_NAVY_DARK = RGBColor(15, 23, 42)      # #0f172a
    C_NAVY_MED  = RGBColor(30, 41, 59)      # #1e293b
    C_BLUE_PRI  = RGBColor(37, 99, 235)     # #2563eb
    C_BLUE_LIGHT= RGBColor(96, 165, 250)    # #60a5fa
    C_BG_LIGHT  = RGBColor(248, 250, 252)   # #f8fafc
    C_CARD_BG   = RGBColor(255, 255, 255)   # #ffffff
    C_CARD_BORDER= RGBColor(226, 232, 240)  # #e2e8f0
    C_TEXT_DARK = RGBColor(15, 23, 42)      # #0f172a
    C_TEXT_MUTED= RGBColor(100, 116, 139)   # #64748b
    C_RED       = RGBColor(220, 38, 38)     # #dc2626
    C_GREEN     = RGBColor(5, 150, 105)     # #059669
    C_PURPLE    = RGBColor(124, 58, 237)    # #7c3aed
    C_AMBER     = RGBColor(217, 119, 6)     # #d97706

    def add_bg(slide, color):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, category_text="AEROPV-TRUST FRAMEWORK"):
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.7), Inches(0.3))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = C_BLUE_PRI

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.7), Inches(0.65))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(21)
        p_title.font.bold = True
        p_title.font.color.rgb = C_TEXT_DARK

    def add_card(slide, left, top, width, height, title="", border_color=C_BLUE_PRI):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = C_CARD_BG
        card.line.color.rgb = C_CARD_BORDER
        card.line.width = Pt(1)

        # Left accent stripe
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(0.12), Inches(height))
        bar.fill.solid()
        bar.fill.fore_color.rgb = border_color
        bar.line.fill.background()

        if title:
            tbox = slide.shapes.add_textbox(Inches(left + 0.28), Inches(top + 0.12), Inches(width - 0.4), Inches(0.38))
            tf = tbox.text_frame
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = C_TEXT_DARK

        content_box = slide.shapes.add_textbox(Inches(left + 0.28), Inches(top + (0.48 if title else 0.15)), Inches(width - 0.45), Inches(height - (0.58 if title else 0.25)))
        content_tf = content_box.text_frame
        content_tf.word_wrap = True
        return content_tf

    def add_ui_sync_banner(slide, left, top, width, height, cue_text):
        banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        banner.fill.solid()
        banner.fill.fore_color.rgb = RGBColor(238, 242, 255)
        banner.line.color.rgb = RGBColor(199, 210, 254)
        banner.line.width = Pt(1)

        tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left + 0.18), Inches(top + 0.12), Inches(1.3), Inches(0.32))
        tag.fill.solid()
        tag.fill.fore_color.rgb = C_BLUE_PRI
        tag.line.fill.background()
        tf_tag = tag.text_frame
        p_t = tf_tag.paragraphs[0]
        p_t.text = "LIVE UI SYNC"
        p_t.font.size = Pt(9)
        p_t.font.bold = True
        p_t.font.color.rgb = RGBColor(255, 255, 255)
        p_t.alignment = PP_ALIGN.CENTER

        cbox = slide.shapes.add_textbox(Inches(left + 1.6), Inches(top + 0.08), Inches(width - 1.8), Inches(height - 0.15))
        tf_c = cbox.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.text = cue_text
        p_c.font.size = Pt(10.5)
        p_c.font.color.rgb = RGBColor(30, 58, 138)
        p_c.font.bold = True

    # ==========================================
    # SLIDE 1: TITLE SLIDE (Interactive Presentation Mode)
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1, C_NAVY_DARK)

    # Accent line
    dec = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.5), Inches(1.2), Inches(0.08))
    dec.fill.solid()
    dec.fill.fore_color.rgb = C_BLUE_LIGHT
    dec.line.fill.background()

    t_box = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(2.2))
    tf1 = t_box.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "AeroPV-Trust"
    p1.font.size = Pt(44)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(255, 255, 255)

    p2 = tf1.add_paragraph()
    p2.text = "Multimodal UAV Photovoltaic Anomaly Inspection & Trustworthy XAI Framework"
    p2.font.size = Pt(20)
    p2.font.color.rgb = C_BLUE_LIGHT
    p2.space_before = Pt(8)

    meta_box = s1.shapes.add_textbox(Inches(1.0), Inches(4.3), Inches(11.3), Inches(2.4))
    tf_meta = meta_box.text_frame
    p_author = tf_meta.paragraphs[0]
    p_author.text = "Author & Presenter: Labonna Prava"
    p_author.font.size = Pt(17)
    p_author.font.bold = True
    p_author.font.color.rgb = RGBColor(255, 255, 255)

    p_info = tf_meta.add_paragraph()
    p_info.text = "Dual-Screen Interactive Walkthrough (Slides + Live Streamlit UI Dashboard at http://localhost:8501)"
    p_info.font.size = Pt(13)
    p_info.font.color.rgb = C_BLUE_LIGHT
    p_info.space_before = Pt(6)

    p_core = tf_meta.add_paragraph()
    p_core.text = "Key Pillars: Radiometric Physics | Computer Vision Segmentation | Late Fusion Neural Net | 4-Tier Trustworthy XAI"
    p_core.font.size = Pt(12)
    p_core.font.color.rgb = C_TEXT_MUTED
    p_core.space_before = Pt(4)

    # ==========================================
    # SLIDE 2: PROBLEM STATEMENT & MOTIVATION
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2, C_BG_LIGHT)
    add_header(s2, "Problem Statement: The Aerial Solar Inspection Dilemma")

    c1 = add_card(s2, 0.8, 1.45, 3.6, 4.4, "The O&M Challenge", C_BLUE_PRI)
    c1_pts = [
        "Massive Scale: Modern solar plants span hundreds of hectares containing over 100,000 PV modules.",
        "Severe Yield Degradation: Defective cells, bypass diode short-circuits, and localized hot spots lead to massive power loss and fire risk.",
        "Manual Bottlenecks: Foot patrols with handheld thermal cameras are slow, hazardous, and lack objective verification."
    ]
    for b in c1_pts:
        p = c1.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(10.5)
        p.space_before = Pt(4)

    c2 = add_card(s2, 4.8, 1.45, 3.6, 4.4, "Single-Sensor Blindspots", C_RED)
    c2_pts = [
        "Visual RGB Alone: Detects surface dust, glass fractures, and soiling, but is completely blind to internal electrical failures.",
        "Thermal IR Alone: Highlights thermal hot spots, but cannot distinguish between true internal defects and harmless surface shadows or bird droppings.",
        "The Dilemma: Neither sensor alone can provide high-confidence autonomous defect classification."
    ]
    for b in c2_pts:
        p = c2.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(10.5)
        p.space_before = Pt(4)

    c3 = add_card(s2, 8.8, 1.45, 3.7, 4.4, "The AeroPV-Trust Solution", C_GREEN)
    c3_pts = [
        "Synchronized Multimodal Fusion: Unites visual RGB texture with calibrated 16-bit radiometric thermal radiance.",
        "Physics-Grounded Preprocessing: Rigorous Kelvin-to-Celsius conversion and robust scene-adaptive MAD scoring.",
        "Dual-Branch Late Fusion: Modality-specific encoders preserving distinct visual and thermal distributions.",
        "Multi-Level XAI: Provides verifiable evidence (Grad-CAM heatmaps) so operators know exactly why a panel is flagged."
    ]
    for b in c3_pts:
        p = c3.add_paragraph()
        p.text = "• " + b
        p.font.size = Pt(10.5)
        p.space_before = Pt(4)

    add_ui_sync_banner(s2, 0.8, 6.05, 11.7, 0.95,
        "Speaker Cue: Point to the Executive Header on the live UI. Explain that AeroPV-Trust solves the single-sensor blindness dilemma by uniting drone RGB with radiometric thermal imagery into a unified trustworthy framework.")

    # ==========================================
    # SLIDE 3: SYSTEM ARCHITECTURE & UI FLOWCHART
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3, C_BG_LIGHT)
    add_header(s3, "System Architecture & Interactive UI Operational Flowchart")

    c3_flow = add_card(s3, 0.8, 1.45, 7.8, 4.4, "End-to-End Operational Pipeline Flow", C_BLUE_PRI)
    flow_steps = [
        "1. Drone Sensor Ingestion: High-Res Visual RGB (1280×720) + Radiometric 16-bit Thermal TIFF (640×512).",
        "2. Radiometric Thermal Conversion: Raw microbolometer sensor counts -> physical Celsius (°C).",
        "3. Coordinate Alignment: Calibrated spatial bounding crops and aspect-ratio interpolation.",
        "4. Autonomous Computer Vision Segmentation: Multi-channel HSV Gaussian Otsu + morphological opening.",
        "5. Robust MAD Anomaly Scoring: Module T95 temperature compared against scene median absolute deviation.",
        "6. Dual-Branch Late Fusion Classification: Modality-specific ResNet-18 encoders -> 512-d fused latent vector.",
        "7. Trustworthy Multi-Level XAI: Dual Grad-CAM overlays + Modality ablation + Prototype retrieval + CCME.",
        "8. Operator Delivery: Interactive Streamlit console with telemetry cards, filters, and one-click CSV export."
    ]
    for fs in flow_steps:
        p = c3_flow.add_paragraph()
        p.text = fs
        p.font.size = Pt(10)
        p.space_before = Pt(3)

    c3_ui = add_card(s3, 8.8, 1.45, 3.7, 4.4, "Live UI Navigation Map", C_PURPLE)
    ui_map = [
        "Sidebar Controls: Choose 'Preset Drone Dataset' or 'Upload Custom Capture'. Set anomaly confidence threshold (0.1-0.9).",
        "Tab 1 (Inspection Console): Runs alignment, module extraction, and displays side-by-side overlays + Grad-CAM heatmaps.",
        "Tab 2 (Attribution & Diagnostics): Explains Modality Ablation, Prototype Retrieval, and CCME counterfactuals.",
        "Tab 3 (Model Benchmarks): Live comparative performance table.",
        "Tab 4 (System Architecture): Neural topology schematic."
    ]
    for um in ui_map:
        p = c3_ui.add_paragraph()
        p.text = "• " + um
        p.font.size = Pt(9.8)
        p.space_before = Pt(4)

    add_ui_sync_banner(s3, 0.8, 6.05, 11.7, 0.95,
        "Speaker Cue: Switch to Tab 4 (System Architecture) on the UI. Show the architectural block diagram and explain how the pipeline takes raw drone flights all the way to decision evidence.")

    # ==========================================
    # SLIDE 4: IMAGE CAPTURE & SPATIAL ALIGNMENT
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4, C_BG_LIGHT)
    add_header(s4, "Aerial Image Acquisition & Spatial Alignment Pipeline")

    c4_1 = add_card(s4, 0.8, 1.45, 5.7, 4.4, "Dual-Sensor Ingestion Specifications", C_BLUE_PRI)
    p = c4_1.paragraphs[0]
    p.text = "Sensor Specifications & Data Modalities:"
    p.font.bold = True
    p.font.size = Pt(12)
    s_pts = [
        "Visual Sensor (RGB): 1280 × 720, 24-bit color. Captures micro-textures, module frames, busbars, glass cracks, and surface soiling.",
        "Thermal Infrared (IR): 640 × 512 radiometric uncooled microbolometer. Captures radiant flux as 16-bit digital sensor counts.",
        "Data Ingestion Flexibility: Supports the Kaggle Photovoltaic System O&M Inspection schema (single-row and double-row flights) as well as custom user uploads via UI."
    ]
    for sp in s_pts:
        p = c4_1.add_paragraph()
        p.text = "• " + sp
        p.font.size = Pt(10.5)
        p.space_before = Pt(4)

    c4_2 = add_card(s4, 6.8, 1.45, 5.7, 4.4, "Spatial Perspective Alignment", C_BLUE_LIGHT)
    p = c4_2.paragraphs[0]
    p.text = "Resolving Optical FOV Discrepancies:"
    p.font.bold = True
    p.font.size = Pt(12)
    align_pts = [
        "Focal Discrepancy: Dual drone cameras have distinct focal lengths and fields of view (FOV).",
        "Calibrated Crop Boundaries: Defined calibrated bounding boxes: IR_CROP (0, 75, 640, 490) and RGB_CROP (120, 0, 1150, 720).",
        "Spatial Interpolation: RGB crop is resized to match thermal spatial resolution using cv2.INTER_AREA interpolation, ensuring 1:1 pixel coordinate alignment.",
        "Thermal Colormapping: 8-bit normalized thermal representations are mapped to the Inferno colormap for operator inspection."
    ]
    for ap in align_pts:
        p = c4_2.add_paragraph()
        p.text = "• " + ap
        p.font.size = Pt(10.5)
        p.space_before = Pt(4)

    add_ui_sync_banner(s4, 0.8, 6.05, 11.7, 0.95,
        "Speaker Cue: On UI Tab 1 (Inspection Console), show the side-by-side 'Visual Inspection (RGB)' and 'Thermographic Inspection (IR Inferno)' panels. Point out how the bounding boxes perfectly align across both modalities.")

    # ==========================================
    # SLIDE 5: CORE WORKING MATHEMATICAL FORMULAS
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5, C_BG_LIGHT)
    add_header(s5, "Core Mathematical & Radiometric Formulations")

    c5_1 = add_card(s5, 0.8, 1.45, 3.6, 4.4, "1. Radiometric Celsius Calibration", C_BLUE_PRI)
    p = c5_1.paragraphs[0]
    p.text = "Kelvin-to-Celsius Physics Formula:"
    p.font.bold = True
    p.font.size = Pt(11)

    p_eq1 = c5_1.add_paragraph()
    p_eq1.text = "T(°C) = (Raw Counts × 0.04) - 273.15"
    p_eq1.font.bold = True
    p_eq1.font.size = Pt(11.5)
    p_eq1.font.color.rgb = C_BLUE_PRI
    p_eq1.space_before = Pt(6)

    f1_pts = [
        "Raw Counts: 16-bit digital sensor count values.",
        "0.04: Radiometric sensor scale factor (0.04 K per count).",
        "-273.15: Conversion offset from absolute Kelvin to Celsius.",
        "Outputs true calibrated surface temperatures for every pixel."
    ]
    for fp in f1_pts:
        p = c5_1.add_paragraph()
        p.text = "• " + fp
        p.font.size = Pt(10)
        p.space_before = Pt(3)

    c5_2 = add_card(s5, 4.8, 1.45, 3.6, 4.4, "2. Robust MAD Anomaly Score", C_GREEN)
    p = c5_2.paragraphs[0]
    p.text = "Outlier-Resilient Statistical Metric:"
    p.font.bold = True
    p.font.size = Pt(11)

    p_eq2 = c5_2.add_paragraph()
    p_eq2.text = "Score = (T95(mod) - Med(scene)) / (MAD + ε)"
    p_eq2.font.bold = True
    p_eq2.font.size = Pt(11)
    p_eq2.font.color.rgb = C_GREEN
    p_eq2.space_before = Pt(4)

    p_sub = c5_2.add_paragraph()
    p_sub.text = "where MAD = Median( | Ti - Median(scene) | )"
    p_sub.font.italic = True
    p_sub.font.size = Pt(9.5)
    p_sub.font.color.rgb = C_TEXT_MUTED

    f2_pts = [
        "T95(module): 95th-percentile temperature in module.",
        "Median & MAD: Resilient baseline unaffected by extreme defective outliers.",
        "Experimental Threshold: Score > 2.2931 flags candidate anomalies for neural evaluation."
    ]
    for fp in f2_pts:
        p = c5_2.add_paragraph()
        p.text = "• " + fp
        p.font.size = Pt(10)
        p.space_before = Pt(3)

    c5_3 = add_card(s5, 8.8, 1.45, 3.7, 4.4, "3. Weighted Cross-Entropy Loss", C_PURPLE)
    p = c5_3.paragraphs[0]
    p.text = "Class-Imbalance Loss Function:"
    p.font.bold = True
    p.font.size = Pt(11)

    p_eq3 = c5_3.add_paragraph()
    p_eq3.text = "L = - Σ wc · yc · log( ŷc )"
    p_eq3.font.bold = True
    p_eq3.font.size = Pt(12)
    p_eq3.font.color.rgb = C_PURPLE
    p_eq3.space_before = Pt(4)

    p_sub3 = c5_3.add_paragraph()
    p_sub3.text = "wc = (1 / Nc) / Σ(1 / Nk)"
    p_sub3.font.italic = True
    p_sub3.font.size = Pt(10)
    p_sub3.font.color.rgb = C_TEXT_MUTED

    f3_pts = [
        "Severe 10:1 Imbalance: Defect modules represent only ~10% of field modules.",
        "Inverse Frequency Weighting: Heavily penalizes false negatives.",
        "Eliminates majority-class prediction collapse."
    ]
    for fp in f3_pts:
        p = c5_3.add_paragraph()
        p.text = "• " + fp
        p.font.size = Pt(10)
        p.space_before = Pt(3)

    add_ui_sync_banner(s5, 0.8, 6.05, 11.7, 0.95,
        "Speaker Cue: Point to the extracted module cards on the UI. Highlight that each module displays both its 'Confidence' (from formula 3) and its 'MAD score' (from formula 2).")

    # ==========================================
    # SLIDE 6: DUAL-BRANCH LATE FUSION DEEP LEARNING MODEL
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_bg(s6, C_BG_LIGHT)
    add_header(s6, "Deep Learning Architecture: Dual-Branch Late Fusion")

    c6_1 = add_card(s6, 0.8, 1.45, 5.7, 4.4, "Why Late Fusion Outperforms Early Fusion", C_BLUE_PRI)
    p = c6_1.paragraphs[0]
    p.text = "Modality Independence & Separation:"
    p.font.bold = True
    p.font.size = Pt(12)
    lf_pts = [
        "The Early Fusion Failure: Concatenating RGB + Thermal into 4 channels forces early conv filters to mix visual textures with thermal flux, causing channel interference.",
        "Dual ResNet-18 Backbones: Modality-specific encoders allow the RGB branch to extract visual edge features (cracks, dust) and the Thermal branch to isolate heat gradients.",
        "Non-Linear Projections: Each 512-d backbone output is projected through Linear(512, 256) -> ReLU -> Dropout(0.2).",
        "Fused Multimodal Vector: Concatenates into a 512-d representation [z_rgb || z_thermal] before the shared classification head."
    ]
    for lp in lf_pts:
        p = c6_1.add_paragraph()
        p.text = "• " + lp
        p.font.size = Pt(10)
        p.space_before = Pt(4)

    c6_2 = add_card(s6, 6.8, 1.45, 5.7, 4.4, "Architectural Topology & Training", C_BLUE_LIGHT)
    p = c6_2.paragraphs[0]
    p.text = "Network Parameters & Hyperparameters:"
    p.font.bold = True
    p.font.size = Pt(12)
    train_pts = [
        "RGB Encoder: 3-channel ResNet-18 initialized with ImageNet representations.",
        "Thermal Encoder: 1-channel ResNet-18 adapted for single-channel radiometry.",
        "Batch Size & Optimizer: Adam optimizer (lr=1e-3, weight_decay=1e-4, batch_size=16/32).",
        "Data Augmentation: Random horizontal flipping applied to both modalities simultaneously during training.",
        "Latency & Throughput: Lightweight 512-d representation executes in ~45 ms per frame, ideal for embedded drone hardware."
    ]
    for tp in train_pts:
        p = c6_2.add_paragraph()
        p.text = "• " + tp
        p.font.size = Pt(10)
        p.space_before = Pt(4)

    add_ui_sync_banner(s6, 0.8, 6.05, 11.7, 0.95,
        "Speaker Cue: On UI Tab 1, point to the top status badge 'Late Fusion Architecture • ResNet-18 Dual Encoders • Grad-CAM Attribution' to show that this exact network is serving live inferences.")

    # ==========================================
    # SLIDE 7: EMPIRICAL BENCHMARK & RESULT COMPARISON
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_bg(s7, C_BG_LIGHT)
    add_header(s7, "Empirical Benchmark: Comparative Performance Results")

    # Benchmarks Table
    rows, cols = 5, 6
    table_shape = s7.shapes.add_table(rows, cols, Inches(0.8), Inches(1.45), Inches(11.7), Inches(2.2))
    table = table_shape.table

    headers = ["Model Architecture", "Input Modality", "ROC-AUC", "Average Precision (AP)", "F1-Score", "Balanced Accuracy"]
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_DARK
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(10.5)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.alignment = PP_ALIGN.CENTER

    data = [
        ["RGB-only", "3-Channel RGB", "0.967", "0.873", "0.810", "0.940"],
        ["Thermal-only", "1-Channel Thermal", "0.990", "0.939", "0.636", "0.915"],
        ["Early Fusion", "4-Channel Concatenated", "0.984", "0.952", "0.686", "0.938"],
        ["Late Fusion (Ours)", "Dual-Branch ResNet-18", "0.996", "0.971", "0.783", "0.958"]
    ]

    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = val
            cell.fill.solid()
            if i == 3:
                cell.fill.fore_color.rgb = RGBColor(238, 242, 255)
            else:
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255) if i % 2 == 0 else RGBColor(248, 250, 252)

            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10.5)
                p.font.color.rgb = C_BLUE_PRI if (i == 3 and j in [2, 3, 5]) else C_TEXT_DARK
                p.font.bold = True if (i == 3) else False
                p.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT

    c7_1 = add_card(s7, 0.8, 3.8, 5.7, 2.1, "Discrimination Superiority", C_BLUE_PRI)
    b_pts1 = [
        "Late Fusion achieves highest ROC-AUC (0.996) and Average Precision (0.971).",
        "Eliminates early convolution feature mixing, boosting positive defect identification.",
        "Maintains high balanced accuracy (0.958) on rare defect samples."
    ]
    for bp in b_pts1:
        p = c7_1.add_paragraph()
        p.text = "• " + bp
        p.font.size = Pt(9.8)
        p.space_before = Pt(2)

    c7_2 = add_card(s7, 6.8, 3.8, 5.7, 2.1, "Confidence Reliability (Brier Score)", C_GREEN)
    b_pts2 = [
        "Lowest Brier Score (0.0401): Measures mean squared probability error.",
        "Autonomous UAV Safety: High probability flags genuinely represent ground-truth failures.",
        "Prevents costly false alarms and unnecessary field technician truck rolls."
    ]
    for bp in b_pts2:
        p = c7_2.add_paragraph()
        p.text = "• " + bp
        p.font.size = Pt(9.8)
        p.space_before = Pt(2)

    add_ui_sync_banner(s7, 0.8, 6.05, 11.7, 0.95,
        "Speaker Cue: Switch to Tab 3 (Model Benchmarks) on the live UI. Show that this exact benchmark table is integrated into the dashboard, confirming the superiority of Late Fusion.")

    # ==========================================
    # SLIDE 8: DEFECT DECISION CRITERIA
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_bg(s8, C_BG_LIGHT)
    add_header(s8, "Defect Decision Criteria: How We Prove a Panel is Defective")

    c8_1 = add_card(s8, 0.8, 1.45, 3.6, 4.4, "1. Statistical Evidence", C_AMBER)
    p = c8_1.paragraphs[0]
    p.text = "Scene Thermal MAD Score:"
    p.font.bold = True
    p.font.size = Pt(11.5)
    e1_pts = [
        "First Line of Defense: The module's 95th-percentile temperature is scaled against the entire array's median absolute deviation.",
        "Criterion: Thermal MAD Score > 2.2931 (representing the top 10% hottest modules).",
        "Example (Module 04): Attains a MAD score of 9.92, showing extreme localized heating."
    ]
    for ep in e1_pts:
        p = c8_1.add_paragraph()
        p.text = "• " + ep
        p.font.size = Pt(10)
        p.space_before = Pt(4)

    c8_2 = add_card(s8, 4.8, 1.45, 3.6, 4.4, "2. Neural Confidence Evidence", C_RED)
    p = c8_2.paragraphs[0]
    p.text = "Late Fusion Probability:"
    p.font.bold = True
    p.font.size = Pt(11.5)
    e2_pts = [
        "Multimodal Classification: Both visual and thermal features are evaluated jointly through the trained late-fusion head.",
        "Criterion: Anomaly Probability P(Defect) >= Threshold (default 0.50).",
        "Example (Module 04): Reaches 0.9991 (99.91% neural confidence), while normal modules remain at 0.0000."
    ]
    for ep in e2_pts:
        p = c8_2.add_paragraph()
        p.text = "• " + ep
        p.font.size = Pt(10)
        p.space_before = Pt(4)

    c8_3 = add_card(s8, 8.8, 1.45, 3.7, 4.4, "3. Spatial Verification (Grad-CAM)", C_GREEN)
    p = c8_3.paragraphs[0]
    p.text = "Visual Heatmap Attribution:"
    p.font.bold = True
    p.font.size = Pt(11.5)
    e3_pts = [
        "Explainable Spatial Grounding: Computes gradient backpropagation into final conv layers.",
        "Criterion: Heatmap activation must concentrate specifically on the internal solar cell or diode.",
        "Eliminates False Positives: Verifies the model did not trigger on external shadows, ground grass, or mounting racks."
    ]
    for ep in e3_pts:
        p = c8_3.add_paragraph()
        p.text = "• " + ep
        p.font.size = Pt(10)
        p.space_before = Pt(4)

    add_ui_sync_banner(s8, 0.8, 6.05, 11.7, 0.95,
        "Speaker Cue: On UI Tab 1, scroll directly to Module 04. Show the red 'DEFECT' badge, the 0.999 confidence, the 9.92 MAD score, and the Grad-CAM red heatmap highlighting the exact hot-spot cell.")

    # ==========================================
    # SLIDE 9: TRUSTWORTHY XAI & DIAGNOSTICS
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_bg(s9, C_BG_LIGHT)
    add_header(s9, "Trustworthy XAI: Modality Ablation & Prototype Retrieval")

    c9_1 = add_card(s9, 0.8, 1.45, 3.6, 4.4, "Modality Ablation Analysis", C_BLUE_PRI)
    p = c9_1.paragraphs[0]
    p.text = "Marginal Contribution Drops:"
    p.font.bold = True
    p.font.size = Pt(11.5)
    x1_pts = [
        "Modality Masking: Evaluates model with full input, then with RGB zeroed out, then with Thermal zeroed out.",
        "Thermal Drop (ΔP_th = 0.157): Removing thermal drops defect probability by 15.7%, proving thermal radiance dominates for electrical hot spots.",
        "RGB Drop (ΔP_rgb): Isolates reliance on visual glass cracking and soiling."
    ]
    for xp in x1_pts:
        p = c9_1.add_paragraph()
        p.text = "• " + xp
        p.font.size = Pt(10)
        p.space_before = Pt(4)

    c9_2 = add_card(s9, 4.8, 1.45, 3.6, 4.4, "Prototype Exemplar Retrieval", C_PURPLE)
    p = c9_2.paragraphs[0]
    p.text = "Case-Based Reasoning:"
    p.font.bold = True
    p.font.size = Pt(11.5)
    x2_pts = [
        "512-d Fused Unit Hypersphere: L2-normalizes the learned representation z_norm = z / ||z||2.",
        "Cosine Similarity: Measures angular proximity between test module query and historical training bank.",
        "Top-k Retrieval: Instantly surfaces matching historical exemplars so engineers can compare current defects to known cases."
    ]
    for xp in x2_pts:
        p = c9_2.add_paragraph()
        p.text = "• " + xp
        p.font.size = Pt(10)
        p.space_before = Pt(4)

    c9_3 = add_card(s9, 8.8, 1.45, 3.7, 4.4, "Counterfactual CCME", C_GREEN)
    p = c9_3.paragraphs[0]
    p.text = "Latent Boundary Exploration:"
    p.font.bold = True
    p.font.size = Pt(11.5)
    x3_pts = [
        "Latent Manifold Interpolation: z(α) = α·z_sample + (1-α)·z_base.",
        "Binary Search for Transition Factor: Finds minimum interpolation strength α needed to flip prediction from normal to defective.",
        "Mathematical Verification: Proves the exact modality threshold driving the model's decision."
    ]
    for xp in x3_pts:
        p = c9_3.add_paragraph()
        p.text = "• " + xp
        p.font.size = Pt(10)
        p.space_before = Pt(4)

    add_ui_sync_banner(s9, 0.8, 6.05, 11.7, 0.95,
        "Speaker Cue: Switch to Tab 2 (Attribution & Diagnostics) on the UI. Explain how field engineers rely on Modality Sensitivity and Prototype Retrieval to verify autonomous drone detections.")

    # ==========================================
    # SLIDE 10: LIVE DEMONSTRATION & CONCLUSION
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_bg(s10, C_NAVY_DARK)

    cat_b = s10.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
    cat_b.text_frame.paragraphs[0].text = "LIVE DEMONSTRATION CUES & SUMMARY"
    cat_b.text_frame.paragraphs[0].font.size = Pt(10)
    cat_b.text_frame.paragraphs[0].font.bold = True
    cat_b.text_frame.paragraphs[0].font.color.rgb = C_BLUE_LIGHT

    t_b = s10.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.7))
    t_b.text_frame.paragraphs[0].text = "Live Demonstration Guide & Conclusion"
    t_b.text_frame.paragraphs[0].font.size = Pt(22)
    t_b.text_frame.paragraphs[0].font.bold = True
    t_b.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    # Left Card
    c10_1 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.7), Inches(4.3))
    c10_1.fill.solid()
    c10_1.fill.fore_color.rgb = C_NAVY_MED
    c10_1.line.color.rgb = RGBColor(51, 65, 85)
    tb1 = s10.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(5.3), Inches(4.0))
    tf_c10_1 = tb1.text_frame
    tf_c10_1.word_wrap = True
    p = tf_c10_1.paragraphs[0]
    p.text = "Live UI Interactive Demonstration Script:"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_BLUE_LIGHT
    demo_steps = [
        "1. Select Capture: Choose preset 'DJI_0000_RGB' or drag-and-drop custom drone imagery.",
        "2. Click 'Run Inspection': Observe end-to-end alignment and inference executing in ~45 ms.",
        "3. Review Telemetry Strip: Point out 12 Modules Detected, 1 Anomaly Flagged (8.3% Defect Rate).",
        "4. Showcase Defective Module 04: Point to the Grad-CAM localized red heat cluster and 0.9991 confidence.",
        "5. Toggle Display Filter: Switch from 'All Modules' to 'Defects Only' to show instant operational triage.",
        "6. Click 'Export CSV Report': Download structured CSV report containing bounding boxes, MAD scores, and probabilities."
    ]
    for ds in demo_steps:
        p = tf_c10_1.add_paragraph()
        p.text = ds
        p.font.size = Pt(9.8)
        p.font.color.rgb = RGBColor(241, 245, 249)
        p.space_before = Pt(3)

    # Right Card
    c10_2 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.3))
    c10_2.fill.solid()
    c10_2.fill.fore_color.rgb = C_NAVY_MED
    c10_2.line.color.rgb = RGBColor(51, 65, 85)
    tb2 = s10.shapes.add_textbox(Inches(7.0), Inches(1.65), Inches(5.3), Inches(4.0))
    tf_c10_2 = tb2.text_frame
    tf_c10_2.word_wrap = True
    p2 = tf_c10_2.paragraphs[0]
    p2.text = "Conclusion & Industrial Impact:"
    p2.font.size = Pt(13)
    p2.font.bold = True
    p2.font.color.rgb = C_GREEN
    impact_pts = [
        "Multimodal Synergy: Solves single-modality blindness by fusing high-res RGB with radiometric thermal radiance.",
        "State-of-the-Art Precision: 0.996 ROC-AUC and 0.0401 Brier calibration score guarantee high operational confidence.",
        "Trustworthy AI: Eliminates black-box barriers with spatial Grad-CAM, prototype retrieval, and counterfactuals.",
        "Open Framework: Authored by Labonna Prava | https://github.com/labonna-prava/AeroPV-Trust"
    ]
    for ip in impact_pts:
        p = tf_c10_2.add_paragraph()
        p.text = "• " + ip
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(241, 245, 249)
        p.space_before = Pt(4)

    # Banner at bottom
    add_ui_sync_banner(s10, 0.8, 6.0, 11.7, 1.0,
        "Speaker Closing: 'Thank you for your time. The live dashboard at http://localhost:8501 is running and available for interactive testing. I am now open to your questions.'")

    # Save presentation
    scripts_dir = Path(__file__).resolve().parent
    root_dir = scripts_dir.parent

    out_file_scripts = scripts_dir / "AeroPV-Trust_Comprehensive_Presentation.pptx"
    out_file_root = root_dir / "AeroPV-Trust_Comprehensive_Presentation.pptx"

    prs.save(str(out_file_scripts))
    prs.save(str(out_file_root))
    print(f"[AeroPV-Trust] PowerPoint Presentation updated in scripts: {out_file_scripts}")
    print(f"[AeroPV-Trust] PowerPoint Presentation updated in root: {out_file_root}")

if __name__ == "__main__":
    create_deck()
