import streamlit as st
import google.generativeai as genai
from PIL import Image
import fitz
import io, os, re, pickle
import cv2, numpy as np
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Inscribe 2.0 – Hebrew OCR",
    page_icon="✍️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #F0F2F5; }
    [data-testid="stSidebar"] { background-color: #10243E !important; }
    [data-testid="stSidebar"] * { color: #C5D4E3 !important; }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 { color: #FFFFFF !important; font-size: 1.4rem !important; }
    .gold { color: #C79A3B; font-weight: 700; }
    .page-title { font-size: 1.8rem; font-weight: 800; color: #10243E; margin-bottom: 0; }
    .page-sub { font-size: 0.95rem; color: #667482; margin-top: 2px; margin-bottom: 20px; }
    .divider { border: none; border-top: 1px solid #DDE3EA; margin: 16px 0 24px 0; }
    .card { background: #FFFFFF; border: 1px solid #DDE3EA; border-radius: 10px;
            padding: 24px 28px; margin-bottom: 20px; }
    .rtl-box { direction: rtl; text-align: right;
               font-family: "David Libre","Times New Roman",serif;
               font-size: 1.15rem; line-height: 1.9;
               background: #FAFBFC; border: 1px solid #DDE3EA;
               border-radius: 8px; padding: 20px 24px;
               min-height: 200px; white-space: pre-wrap; color: #1C2733; }
    .col-label { font-size: 0.78rem; font-weight: 700; text-transform: uppercase;
                 letter-spacing: .06em; color: #667482; margin-bottom: 6px; }
    .glyph-box { background:#F0F2F5; border:1px solid #DDE3EA; border-radius:8px;
                 padding:10px; text-align:center; margin-bottom:8px; }
    .stButton > button { background-color: #C79A3B !important; color: white !important;
                         font-weight: 700 !important; border: none !important;
                         border-radius: 7px !important; padding: 10px 24px !important; }
    .stButton > button:hover { background-color: #AD812B !important; }
    [data-testid="stFileUploader"] { border: 2px dashed #C79A3B !important;
        border-radius: 10px !important; background: #FFFDF6 !important; padding: 16px !important; }
    footer { visibility: hidden; }
    .stDeployButton { display: none; }
</style>
""", unsafe_allow_html=True)

# ── Word export ───────────────────────────────────────────────────────────────
def make_rtl_para(p, size=12):
    pPr = p._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is None: pPr.append(OxmlElement("w:bidi"))
    p.alignment = 2
    for run in p.runs:
        run.font.name = "David Libre"; run.font.size = Pt(size)
        rPr = run._r.get_or_add_rPr()
        if rPr.find(qn("w:rtl")) is None: rPr.append(OxmlElement("w:rtl"))

def build_word(text):
    doc = Document()
    for line in text.splitlines():
        p = doc.add_paragraph(); p.add_run(line or ""); make_rtl_para(p)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

# ── API key helpers ───────────────────────────────────────────────────────────
def configured_api_key():
    try: secret = st.secrets.get("GEMINI_API_KEY","")
    except Exception: secret = ""
    return (secret or st.session_state.get("api_key","")).strip()

def get_model(model_name="gemini-2.5-flash"):
    key = configured_api_key()
    if not key: return None
    genai.configure(api_key=key)
    try: name = st.secrets.get("GEMINI_MODEL", model_name)
    except Exception: name = model_name
    return genai.GenerativeModel(name)

# ── Glyph segmentation (same as desktop) ─────────────────────────────────────
def segment_glyphs(pil_img):
    gray = np.array(pil_img.convert("L"))
    bw = cv2.threshold(gray,0,255,cv2.THRESH_BINARY_INV+cv2.THRESH_OTSU)[1]
    n,labels,stats,_ = cv2.connectedComponentsWithStats(bw,8)
    glyphs = []
    for x,y,w,h,area in stats[1:]:
        if area>18 and h>8 and w>2:
            crop = gray[max(0,y-3):y+h+3, max(0,x-3):x+w+3]
            glyphs.append((Image.fromarray(crop),(x,y,w,h)))
    glyphs.sort(key=lambda g:(g[1][1]//20,-g[1][0]))
    return glyphs

def glyph_feat(pil_img):
    a = cv2.resize(np.array(pil_img.convert("L")),(28,28)).astype(float)/255.0
    return a.flatten()

# ── Writer library profile helpers ───────────────────────────────────────────
def serialize_profile(name, samples):
    """Serialize profile to bytes for download."""
    buf = io.BytesIO()
    pickle.dump({"name":name,"samples":samples,"version":2}, buf)
    return buf.getvalue()

def load_profile_bytes(b):
    """Load profile from uploaded bytes."""
    return pickle.loads(b)

def samples_to_pil_map(samples):
    """Build {letter: [PIL images]} from samples list."""
    mapping = {}
    for feat, letter in samples:
        arr = (feat.reshape(28,28)*255).astype(np.uint8)
        img = Image.fromarray(arr).resize((56,56), Image.NEAREST)
        mapping.setdefault(letter,[]).append(img)
    return mapping

# ── Gemini OCR with optional writer context ───────────────────────────────────
PRINT_PROMPT = """You are a precise Hebrew OCR engine.
Extract ALL Hebrew (and any English) text from this image exactly as it appears.
Preserve line breaks. Do not translate or summarize. Output only the extracted text."""

def build_hand_prompt(writer_samples=None):
    base = (
        "You are a specialist in Hebrew handwriting transcription.\n"
        "Carefully read the handwritten Hebrew text and transcribe it exactly as written.\n"
        "- Preserve line breaks as they appear.\n"
        "- If a word is unclear, mark it with (?) after the word.\n"
        "- Do not translate or explain. Output only the transcribed Hebrew text."
    )
    if writer_samples:
        letter_list = ", ".join(sorted(set(writer_samples.keys())))
        base += (
            f"\n\nWRITER REFERENCE: The following letters have been calibrated for this writer: "
            f"{letter_list}. "
            f"Use the provided letter sample images to recognize this writer's unique style. "
            f"Pay close attention to their specific letter forms when transcribing."
        )
    return base

def ocr_image_gemini(pil_img, mode, writer_samples=None):
    model = get_model()
    if not model: return "⚠️ No API key set."
    if mode == "printed":
        content = [PRINT_PROMPT, pil_img]
    else:
        prompt = build_hand_prompt(writer_samples)
        content = [prompt]
        # Add up to 2 sample images per letter as visual reference
        if writer_samples:
            content.append("\n\nWRITER LETTER SAMPLES (use these to recognize this writer's style):\n")
            for letter, imgs in sorted(writer_samples.items()):
                content.append(f"Letter '{letter}':")
                content.extend(imgs[:2])  # max 2 examples per letter
            content.append("\n\nNOW TRANSCRIBE THIS HANDWRITTEN PAGE:")
        content.append(pil_img)
    try:
        resp = model.generate_content(content)
        return resp.text.strip()
    except Exception as e:
        return f"❌ Error: {e}"

def ocr_pdf_gemini(pdf_bytes, mode, writer_samples=None):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results = []
    prog = st.progress(0, text="Processing pages…")
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(200/72,200/72), alpha=False)
        img = Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
        text = ocr_image_gemini(img, mode, writer_samples)
        results.append((i+1, text))
        prog.progress((i+1)/len(doc), text=f"Page {i+1} of {len(doc)}…")
    prog.empty()
    return results

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ✍️ Inscribe 2.0")
    st.markdown('<span class="gold">Hebrew OCR — Test Version</span>', unsafe_allow_html=True)
    st.markdown("---")

    page = st.radio("Navigation", [
        "📄  Printed OCR",
        "✍️  Handwriting OCR",
        "◎  Writer Library",
        "🔬  Side-by-Side Test",
        "ℹ️  About"
    ], label_visibility="collapsed")

    st.markdown("---")

    if not configured_api_key():
        st.markdown("**🔑 Gemini API Key**")
        api_input = st.text_input("", value=st.session_state.get("api_key",""),
            type="password", placeholder="AIza...", label_visibility="collapsed")
        if api_input:
            st.session_state["api_key"] = api_input
            st.success("Key saved ✓", icon="✅")
    else:
        st.success("AI connection ready", icon="✅")

    # Show loaded writer profile in sidebar
    if st.session_state.get("loaded_profile"):
        prof = st.session_state["loaded_profile"]
        st.markdown("---")
        st.markdown(f"**✎ Writer Profile loaded:**")
        st.info(f"👤 {prof['name']}\n\n{len(prof['samples'])} letter samples")

    st.markdown("---")
    st.markdown(
        '<div style="font-size:0.78rem;color:#7A94A8;line-height:1.6">'
        'Institute for Preserving Jewish Literature<br>'
        'inscribe.ipjl@gmail.com<br>929 282 2635</div>',
        unsafe_allow_html=True)

def api_key_notice():
    if not configured_api_key():
        with st.expander("🔑 How to get a free Gemini API key (2 min)", expanded=False):
            st.markdown("""
1. Go to **[aistudio.google.com](https://aistudio.google.com)**
2. Sign in with your Google account
3. Click **"Get API key"** → **"Create API key"**
4. Copy the key and paste it in the sidebar
            """)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: Printed OCR
# ─────────────────────────────────────────────────────────────────────────────
if "Printed" in page:
    st.markdown('<div class="page-title">📄 Printed Hebrew OCR</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub">Extract Hebrew text from printed PDF or image using AI vision</div>', unsafe_allow_html=True)
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    api_key_notice()

    uploaded = st.file_uploader("Upload PDF or image",
        type=["pdf","png","jpg","jpeg","tif","tiff"])

    if uploaded and st.button("▶  Run OCR"):
        if not configured_api_key():
            st.error("Please add your Gemini API key in the sidebar first.")
        else:
            with st.spinner("Running AI OCR…"):
                if uploaded.type == "application/pdf":
                    pages = ocr_pdf_gemini(uploaded.read(),"printed")
                    text = "\n\n".join(f"── Page {n} ──\n{t}" for n,t in pages)
                else:
                    text = ocr_image_gemini(Image.open(uploaded),"printed")
            st.session_state["printed_result"] = text
            st.success("✓ OCR complete!")

    if st.session_state.get("printed_result"):
        text = st.session_state["printed_result"]
        st.markdown("**Result:**")
        st.markdown(f'<div class="rtl-box">{text}</div>', unsafe_allow_html=True)
        edited = st.text_area("✏️ Edit here:", value=text, height=260, label_visibility="visible")
        st.download_button("⬇  Download Word", data=build_word(edited),
            file_name="Inscribe_OCR.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: Handwriting OCR
# ─────────────────────────────────────────────────────────────────────────────
elif "Handwriting" in page:
    st.markdown('<div class="page-title">✍️ Handwriting OCR</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub">Transcribe handwritten Hebrew — AI-powered with optional writer profile</div>', unsafe_allow_html=True)
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    api_key_notice()

    # Writer profile section
    prof = st.session_state.get("loaded_profile")
    if prof:
        st.success(f"✅ Writer profile active: **{prof['name']}** ({len(prof['samples'])} samples) — AI will use this writer's letter style as a reference.")
        if st.button("✕ Remove profile", key="remove_prof"):
            del st.session_state["loaded_profile"]
            st.rerun()
    else:
        st.info("💡 **No writer profile loaded.** Load one from the Writer Library page for better accuracy on a specific writer's handwriting.")

    uploaded = st.file_uploader("Upload handwritten image or PDF",
        type=["pdf","png","jpg","jpeg","tif","tiff"])

    if uploaded and st.button("▶  Transcribe Handwriting"):
        if not configured_api_key():
            st.error("Please add your Gemini API key in the sidebar first.")
        else:
            writer_samples = None
            if prof:
                writer_samples = samples_to_pil_map(prof["samples"])

            with st.spinner("AI is reading the handwriting…"):
                if uploaded.type == "application/pdf":
                    pages = ocr_pdf_gemini(uploaded.read(),"handwriting", writer_samples)
                    text = "\n\n".join(f"── Page {n} ──\n{t}" for n,t in pages)
                else:
                    text = ocr_image_gemini(Image.open(uploaded),"handwriting", writer_samples)
            st.session_state["hand_result"] = text
            uncertain = len(re.findall(r'\(\?\)', text))
            if uncertain:
                st.warning(f"⚠️ {uncertain} word(s) marked as uncertain — please review.")
            else:
                st.success("✓ Transcription complete!")

    if st.session_state.get("hand_result"):
        text = st.session_state["hand_result"]
        st.markdown("**Transcription:**")
        st.markdown(f'<div class="rtl-box">{text}</div>', unsafe_allow_html=True)
        edited = st.text_area("✏️ Correct here:", value=text, height=300, label_visibility="visible")
        st.download_button("⬇  Download Word", data=build_word(edited),
            file_name="Inscribe_Handwriting.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: Writer Library
# ─────────────────────────────────────────────────────────────────────────────
elif "Writer" in page:
    st.markdown('<div class="page-title">◎ Writer Library</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub">Create a writer profile — the AI will use it to better recognize that writer\'s handwriting style</div>', unsafe_allow_html=True)
    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["➕ Create New Profile", "📂 Load Existing Profile"])

    # ── TAB 1: Create profile ─────────────────────────────────────────────────
    with tab1:
        st.markdown("""
**How it works:**
1. Upload a calibration image — write each Hebrew letter separately (one per isolated shape)
2. The app detects each letter shape automatically
3. You label each shape with the correct Hebrew letter
4. Save and download the profile file
5. Upload the profile on the Handwriting OCR page — the AI will use your letter samples as a style reference
        """)

        writer_name = st.text_input("Writer's name / ID", placeholder="e.g.  Rabbi Moshe — Notebook 1")
        calib_img = st.file_uploader("Upload calibration image",
            type=["png","jpg","jpeg","tif","tiff"], key="calib_up")

        if calib_img and writer_name.strip():
            img = Image.open(calib_img)
            glyphs = segment_glyphs(img)
            st.info(f"✓ Detected **{len(glyphs)}** letter shapes in the image.")

            if not glyphs:
                st.error("No shapes found. Try a cleaner image with darker ink on white paper.")
            else:
                # Show all glyphs in a grid for labeling
                st.markdown("### Label each shape")
                st.markdown("Type the matching Hebrew letter for each shape below:")

                if "calib_labels" not in st.session_state:
                    st.session_state["calib_labels"] = {}
                if "calib_glyphs" not in st.session_state or \
                   st.session_state.get("calib_img_name") != calib_img.name:
                    st.session_state["calib_glyphs"] = glyphs
                    st.session_state["calib_img_name"] = calib_img.name
                    st.session_state["calib_labels"] = {}

                # 6-column grid
                cols_per_row = 6
                rows = [glyphs[i:i+cols_per_row] for i in range(0,len(glyphs),cols_per_row)]
                for row_idx, row in enumerate(rows):
                    cols = st.columns(cols_per_row)
                    for col_idx, (glyph_img, box) in enumerate(row):
                        glyph_idx = row_idx*cols_per_row + col_idx
                        with cols[col_idx]:
                            disp = glyph_img.copy()
                            disp.thumbnail((80,80))
                            st.image(disp, use_container_width=False, width=80)
                            label = st.text_input("",
                                key=f"glyph_{glyph_idx}",
                                max_chars=1,
                                placeholder="א",
                                label_visibility="collapsed")
                            if label.strip():
                                st.session_state["calib_labels"][glyph_idx] = label.strip()

                labeled = {i:l for i,l in st.session_state["calib_labels"].items() if l}
                st.markdown(f"**{len(labeled)} / {len(glyphs)} labeled**")

                if labeled and st.button("💾 Save Profile"):
                    samples = []
                    for idx, letter in labeled.items():
                        feat = glyph_feat(glyphs[idx][0])
                        samples.append((feat, letter))
                    name = writer_name.strip()
                    prof_bytes = serialize_profile(name, samples)
                    st.success(f"✅ Profile '{name}' ready with {len(samples)} letter samples!")
                    st.download_button(
                        f"⬇  Download '{name}.inscribe'",
                        data=prof_bytes,
                        file_name=f"{name}.inscribe",
                        mime="application/octet-stream"
                    )
                    st.info("Save this file — upload it on the Handwriting OCR page to activate this writer's profile.")

    # ── TAB 2: Load profile ───────────────────────────────────────────────────
    with tab2:
        st.markdown("Upload a previously saved `.inscribe` profile file to activate it:")
        prof_file = st.file_uploader("Upload profile", type=["inscribe","pkl"], key="prof_up")
        if prof_file:
            try:
                data = load_profile_bytes(prof_file.read())
                st.success(f"✅ Profile loaded: **{data['name']}** — {len(data['samples'])} letter samples")
                if st.button("✅ Use this profile for Handwriting OCR"):
                    st.session_state["loaded_profile"] = data
                    st.success(f"Profile '{data['name']}' is now active! Go to the Handwriting OCR page.")
            except Exception as e:
                st.error(f"Could not load profile: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: Side-by-Side Test
# ─────────────────────────────────────────────────────────────────────────────
elif "Side" in page:
    st.markdown('<div class="page-title">🔬 Side-by-Side Comparison</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub">Compare AI transcription vs. your manual correction — test accuracy on real pages</div>', unsafe_allow_html=True)
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    api_key_notice()

    prof = st.session_state.get("loaded_profile")
    if prof:
        st.success(f"✅ Writer profile active: **{prof['name']}**")

    uploaded = st.file_uploader("Upload a handwritten image",
        type=["png","jpg","jpeg","tif","tiff"])

    if uploaded:
        img = Image.open(uploaded)
        col_img, col_res = st.columns(2)
        with col_img:
            st.markdown('<div class="col-label">Original Image</div>', unsafe_allow_html=True)
            st.image(img, use_container_width=True)
        with col_res:
            st.markdown('<div class="col-label">AI Transcription</div>', unsafe_allow_html=True)
            if st.button("▶  Run AI Transcription"):
                if not configured_api_key():
                    st.error("Add API key in sidebar.")
                else:
                    writer_samples = samples_to_pil_map(prof["samples"]) if prof else None
                    with st.spinner("AI reading handwriting…"):
                        text = ocr_image_gemini(img,"handwriting", writer_samples)
                    st.session_state["compare_result"] = text

            if st.session_state.get("compare_result"):
                text = st.session_state["compare_result"]
                uncertain = len(re.findall(r'\(\?\)', text))
                st.markdown(f'<div class="rtl-box">{text}</div>', unsafe_allow_html=True)
                if uncertain:
                    st.warning(f"{uncertain} uncertain word(s) marked with (?)")

        st.markdown("---")
        st.markdown("**📝 Your manual correction:**")
        ground = st.text_area("Type what it actually says:", height=200,
            label_visibility="collapsed", key="ground_truth",
            placeholder="Type the correct transcription here…")

        if ground and st.session_state.get("compare_result"):
            ai = re.sub(r'\(\?\)','',st.session_state["compare_result"]).strip()
            ai_words = set(ai.split())
            gt_words = set(ground.strip().split())
            if gt_words:
                match = len(ai_words & gt_words) / len(gt_words) * 100
                st.metric("Approximate word match", f"{match:.0f}%",
                    help="Rough estimate — counts matching words regardless of order")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: About
# ─────────────────────────────────────────────────────────────────────────────
elif "About" in page:
    st.markdown('<div class="page-title">ℹ️ About Inscribe 2.0</div>', unsafe_allow_html=True)
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown("""
<div class="card">
<h3 style="color:#10243E">Institute for Preserving Jewish Literature</h3>
<p style="color:#444;line-height:1.8">
Some sefarim are too important to disappear.<br><br>
Across the generations, countless treasured sefarim have become difficult—or nearly 
impossible—to find. Their pages hold the Torah of earlier generations, the insight of great 
talmidei chachamim, and a legacy that deserves to remain accessible.<br><br>
Our organization is dedicated to bringing these rare and out-of-print sefarim back to the 
ציבור through careful reprinting and distribution.
</p>
<p style="color:#C79A3B;font-weight:700">
inscribe.ipjl@gmail.com &nbsp;|&nbsp; 929 282 2635
</p>
</div>
<div class="card">
<h3 style="color:#10243E">How the Writer Library works</h3>
<p style="color:#444;line-height:1.8">
Unlike the desktop version (which does basic shape-matching), the online writer library works differently:<br><br>
✦ You upload a calibration image and label each letter shape<br>
✦ The app saves your labeled letter images as a profile<br>
✦ When transcribing, those letter images are sent to Gemini AI as <strong>visual examples</strong><br>
✦ The AI uses them to understand that specific writer's unique letter forms<br>
✦ This gives the AI a personal style reference — not just shape matching<br><br>
This is experimental — accuracy improves with more labeled letter samples.
</p>
</div>
    """, unsafe_allow_html=True)
