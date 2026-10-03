import streamlit as st
import streamlit.components.v1 as components
import pypdf
import pdfplumber
import docx
import json
import re
import io
import base64
import time
import zipfile

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Sách Nói AI - App Đọc Sách Bằng Giọng Nói",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# SAMPLE BOOK DATA (DÙNG KHI CHƯA UPLOAD TẬP TIN)
# -----------------------------------------------------------------------------
SAMPLE_BOOK = {
    "title": "Dế Mèn Phiêu Lưu Ký",
    "author": "Tô Hoài",
    "chapters": [
        {
            "title": "Chương 1: Tôi sống độc lập từ bé - Một tai họa đầu đời",
            "content": """Bởi tôi ăn uống điều độ và làm việc có chừng mực nên tôi chóng lớn lắm. Chẳng bao lâu tôi đã trở thành một chàng dế mên cường tráng. Đôi càng tôi mẫm bóng. Những cái vuốt ở chân, ở khoeo cứ cứng dần và nhọn hắt. Thỉnh thoảng, muốn thử sự lợi hại của những chiếc vuốt, tôi co cẳng lên, đạp phanh phách vào các ngọn cỏ. Những ngọn cỏ gãy rạp, y như có nhát dao vừa liềm qua.

Đôi cánh tôi, trước kia ngắn hủn hoẳn, bây giờ thành cái áo dài kín xuống tận đuôi. Mỗi khi tôi vũ lên, đã nghe tiếng phành phạch giòn giã. Lúc tôi đi bách bộ thì cả người tôi rung rinh một màu nâu bóng mỡ soi gương được và rất xao động. Đầu tôi to ra và nổi từng tảng, rất bướng. Hai cái răng đen nhánh lúc nào cũng hùm hụp như hai lưỡi liềm máy làm việc.

Sợi râu tôi dài và uốn cong một vẻ rất hùng dũng. Tôi lấy làm hãnh diện với bà con về cặp râu ấy lắm. Cứ mỗi lần đứng trước cửa hang, tôi lại trịnh trọng và khoan thai đưa hai chân trước lên vuốt râu. Tôi tớ con nhà giàu, tính nết hay kiêu ngạo và xốc xếch.

Tôi đi đứng oai vệ. Tôi tớ coi ai ra gì. Cà khịa với tất cả mọi người trong xóm. Khi tôi gáy, tôi quát mấy anh Dế Trũi nghèo khổ, trêu chị Cốc, chọc anh Xiển Tóc. Nhưng rồi một tai họa bất ngờ đã đến với tôi vì tính hung hăng bồng bột ấy."""
        },
        {
            "title": "Chương 2: Cảnh ngộ đáng thương của Dế Trũi và chuyến du hành đầu tiên",
            "content": """Một ngày kia, tôi rời khỏi khu đầm lầy quen thuộc để lên đường khám phá thế giới rộng lớn. Trên đường đi, tôi gặp Dế Trũi - một gã dế chân tay lอะ lอะ, người ngợm thô giáp nhưng tính tình vô cùng thật thà và kiên cường.

Hai chúng tôi kết làm anh em bái mạng. Dế Trũi tuy không đẹp đẽ nhưng lại là một người bạn đường vô cùng trung thành và dũng cảm. Chúng tôi cùng nhau vượt qua những dòng sông chảy xiết, đi qua những đồng cỏ ngút ngàn và đối mặt với vô số hiểm nguy.

Qua những trải nghiệm ấy, tôi mới hiểu rằng tình bạn chân chính và lòng khiêm tốn mới là thứ vũ khí mạnh mẽ nhất trên đời, chứ không phải là vẻ bề ngoài oai phong hay tính tình kiêu ngạo."""
        },
        {
            "title": "Chương 3: Bị bắt làm đồ chơi cho trẻ con và bài học về tự do",
            "content": """Thế rồi một hôm, trong lúc mải mê kiếm ăn ở bãi cỏ xanh ven làng, cả hai chúng tôi đã bị mấy cậu bé bắt sống. Họ nhốt chúng tôi vào một chiếc lồng gõ nhỏ xíu, xung quanh bọc lưới đồng.

Ở trong lồng, dù được cho ăn đầy đủ cỏ tươi và đường ngọt, tôi vẫn cảm thấy ngột ngạt và đau đớn vô cùng. Tôi nhớ đồng cỏ bao la, nhớ làn gió mát lành ban đêm và bầu trời sao rộng lớn. Tự do là điều quý giá nhất mà không thức ăn ngon ngọt nào có thể thay thế được.

Chúng tôi đã cùng nhau lên kế hoạch trốn thoát. Nhờ sự kiên trì đào bới và sự giúp đỡ của những người bạn bên ngoài, cuối cùng cửa lồng đã mở. Chúng tôi trở lại với thiên nhiên tự do, lòng thầm hứa sẽ luôn đấu tranh cho lẽ phải và tình đoàn kết của muôn loài."""
        }
    ]
}

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS: FILE EXTRACTOR & TEXT PROCESSING
# -----------------------------------------------------------------------------
def extract_text_from_pdf(uploaded_file):
    text = ""
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n\n"
    except Exception:
        uploaded_file.seek(0)
        reader = pypdf.PdfReader(uploaded_file)
        for page in reader.pages:
            text += (page.extract_text() or "") + "\n\n"
    return text

def extract_text_from_docx(uploaded_file):
    doc = docx.Document(uploaded_file)
    return "\n\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def extract_text_from_txt(uploaded_file):
    content = uploaded_file.read()
    for encoding in ['utf-8', 'utf-16', 'latin-1', 'cp1252']:
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return content.decode('utf-8', errors='ignore')

def parse_book_chapters(full_text):
    # Tách chương thông qua các từ khóa phổ biến
    chapter_pattern = r'(?=(?:Chương|CHƯƠNG|Chương\s+\d+|Phần\s+\d+|Chapter\s+\d+))'
    parts = re.split(chapter_pattern, full_text)
    
    chapters = []
    for idx, part in enumerate(parts):
        clean_part = part.strip()
        if not clean_part:
            continue
        lines = clean_part.split('\n')
        title = lines[0].strip() if lines else f"Mục {idx + 1}"
        content = "\n".join(lines[1:]).strip() if len(lines) > 1 else clean_part
        chapters.append({"title": title, "content": content})
    
    if not chapters:
        chapters = [{"title": "Toàn bộ nội dung sách", "content": full_text}]
    return chapters

def split_into_sentences(text):
    # Tách câu tiếng Việt chuẩn xác
    sentences = re.split(r'(?<=[.!?])\s+', text)
    return [s.strip() for s in sentences if s.strip()]

# -----------------------------------------------------------------------------
# INITIALIZE SESSION STATES
# -----------------------------------------------------------------------------
if 'current_book' not in st.session_state:
    st.session_state.current_book = SAMPLE_BOOK
if 'chapter_idx' not in st.session_state:
    st.session_state.chapter_idx = 0
if 'font_size' not in st.session_state:
    st.session_state.font_size = 18
if 'theme' not in st.session_state:
    st.session_state.theme = 'Sepia'
if 'bookmarks' not in st.session_state:
    st.session_state.bookmarks = []
if 'reading_history' not in st.session_state:
    st.session_state.reading_history = {}

# -----------------------------------------------------------------------------
# DYNAMIC THEMING CSS
# -----------------------------------------------------------------------------
THEMES = {
    "Sáng (Light)": {
        "bg": "#ffffff",
        "text": "#222222",
        "card_bg": "#f8f9fa",
        "border": "#e9ecef"
    },
    "Sepia (Cổ điển)": {
        "bg": "#fbf0d9",
        "text": "#3f301d",
        "card_bg": "#f4e4c1",
        "border": "#e4d3b0"
    },
    "Tối (Dark)": {
        "bg": "#1e1e1e",
        "text": "#e0e0e0",
        "card_bg": "#2d2d2d",
        "border": "#3d3d3d"
    }
}

active_theme = THEMES.get(st.session_state.theme, THEMES["Sepia (Cổ điển)"])

custom_css = f"""
<style>
    .stApp {{
        background-color: {active_theme['bg']} !important;
        color: {active_theme['text']} !important;
    }}
    .book-reader-container {{
        background-color: {active_theme['card_bg']};
        border: 1px solid {active_theme['border']};
        border-radius: 12px;
        padding: 30px;
        font-size: {st.session_state.font_size}px;
        line-height: 1.8;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        color: {active_theme['text']} !important;
        font-family: 'Merriweather', 'Georgia', serif;
    }}
    .chapter-title {{
        font-family: 'Roboto', sans-serif;
        color: #d9534f;
        font-weight: 700;
        margin-bottom: 20px;
        border-bottom: 2px solid {active_theme['border']};
        padding-bottom: 10px;
    }}
    .stButton>button {{
        border-radius: 8px;
        font-weight: 600;
    }}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# SIDEBAR CONTROL PANEL
# -----------------------------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/color/96/open-book--v1.png", width=64)
    st.title("Sách Nói AI")
    st.caption("Ứng dụng Đọc sách & Chuyển Giọng nói Tiếng Việt")
    
    st.markdown("---")
    st.subheader("📚 Quản lý Sách")
    
    upload_option = st.radio("Nguồn sách:", ["Dùng sách mẫu", "Tải file của bạn (PDF/EPUB/DOCX/TXT)"])
    
    if upload_option == "Tải file của bạn (PDF/EPUB/DOCX/TXT)":
        uploaded_file = st.file_uploader("Chọn tập tin sách:", type=['pdf', 'docx', 'txt'])
        if uploaded_file is not None:
            file_type = uploaded_file.name.split('.')[-1].lower()
            with st.spinner("Đang phân tích và xử lý tập tin..."):
                if file_type == 'pdf':
                    raw_text = extract_text_from_pdf(uploaded_file)
                elif file_type == 'docx':
                    raw_text = extract_text_from_docx(uploaded_file)
                else:
                    raw_text = extract_text_from_txt(uploaded_file)
                
                chapters = parse_book_chapters(raw_text)
                st.session_state.current_book = {
                    "title": uploaded_file.name.rsplit('.', 1)[0],
                    "author": "Người dùng tải lên",
                    "chapters": chapters
                }
                st.success(f"Đã tải thành công {len(chapters)} chương!")
    else:
        st.session_state.current_book = SAMPLE_BOOK

    st.markdown("---")
    st.subheader("⚙️ Cấu hình Giao diện & Giọng đọc")
    
    st.session_state.theme = st.selectbox("Giao diện đọc sách:", list(THEMES.keys()), index=1)
    st.session_state.font_size = st.slider("Kích thước chữ (px):", 14, 32, st.session_state.font_size)
    
    tts_engine = st.selectbox(
        "Động cơ Giọng nói (TTS Engine):",
        ["Trình duyệt Web (Offline - Miễn phí)", "FPT.AI Voice API", "Zalo AI TTS API", "Viettel AI TTS API"]
    )
    
    if tts_engine != "Trình duyệt Web (Offline - Miễn phí)":
        api_key = st.text_input(f"Nhập API Key {tts_engine.split()[0]}:", type="password")
        voice_gender = st.selectbox("Giọng đọc vùng miền:", ["Bắc - Nữ (Ban Mai)", "Bắc - Nam (Lê Minh)", "Nam - Nữ (Lan Nhi)", "Trung - Nam (Quang Tiến)"])

# -----------------------------------------------------------------------------
# MAIN CONTENT AREA
# -----------------------------------------------------------------------------
book = st.session_state.current_book
chapters = book["chapters"]

st.title(f"📖 {book['title']}")
st.caption(f"Tác giả: {book.get('author', 'Chưa rõ')} | Tổng số chương: {len(chapters)}")

# Dynamic Progress Bar
progress_val = (st.session_state.chapter_idx + 1) / len(chapters)
st.progress(progress_val)

# Chapter Navigator Tabs
cols = st.columns([1, 4, 1])
with cols[0]:
    if st.button("⬅️ Chương trước") and st.session_state.chapter_idx > 0:
        st.session_state.chapter_idx -= 1
        st.rerun()

with cols[1]:
    chapter_titles = [f"{i+1}. {ch['title']}" for i, ch in enumerate(chapters)]
    selected_ch_idx = st.selectbox(
        "Chọn chương đọc:",
        options=list(range(len(chapters))),
        format_func=lambda x: chapter_titles[x],
        index=st.session_state.chapter_idx
    )
    if selected_ch_idx != st.session_state.chapter_idx:
        st.session_state.chapter_idx = selected_ch_idx
        st.rerun()

with cols[2]:
    if st.button("Chương sau ➡️") and st.session_state.chapter_idx < len(chapters) - 1:
        st.session_state.chapter_idx += 1
        st.rerun()

current_chapter = chapters[st.session_state.chapter_idx]

# -----------------------------------------------------------------------------
# BROWSER-NATIVE INTERACTIVE TTS PLAYER (HTML5 SpeechSynthesis)
# -----------------------------------------------------------------------------
st.subheader("🔊 Trình phát Âm thanh & Đọc tự động")

# Prepare text for Javascript TTS
text_to_read = f"{current_chapter['title']}. {current_chapter['content']}"
clean_json_text = json.dumps(text_to_read, ensure_ascii=False)

tts_component_code = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: transparent;
            margin: 0;
            padding: 10px;
        }}
        .controls-card {{
            background: #ffffff;
            border-radius: 12px;
            padding: 16px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            align-items: center;
        }}
        button {{
            background-color: #2e7d32;
            color: white;
            border: none;
            padding: 10px 18px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }}
        button:hover {{
            background-color: #1b5e20;
            transform: translateY(-1px);
        }}
        button.stop {{
            background-color: #c62828;
        }}
        button.stop:hover {{
            background-color: #8e0000;
        }}
        .setting-group {{
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 13px;
            color: #333;
        }}
        select, input[type=range] {{
            padding: 4px 8px;
            border-radius: 4px;
            border: 1px solid #ccc;
        }}
        #status-badge {{
            font-size: 12px;
            padding: 4px 10px;
            border-radius: 12px;
            background: #e0e0e0;
            color: #424242;
            font-weight: bold;
        }}
    </style>
</head>
<body>
    <div class="controls-card">
        <button onclick="playTTS()">▶️ Đọc sách</button>
        <button onclick="pauseTTS()">⏸️ Tạm dừng</button>
        <button onclick="resumeTTS()">▶️ Tiếp tục</button>
        <button class="stop" onclick="stopTTS()">⏹️ Dừng đọc</button>
        
        <div class="setting-group">
            <label>⚡ Tốc độ:</label>
            <input type="range" id="rate" min="0.5" max="2.0" value="1.0" step="0.1" onchange="updateRate()">
            <span id="rate-val">1.0x</span>
        </div>

        <div class="setting-group">
            <label>🗣️ Giọng đọc:</label>
            <select id="voice-select"></select>
        </div>

        <span id="status-badge">Sẵn sàng</span>
    </div>

    <script>
        const textContent = {clean_json_text};
        const synth = window.speechSynthesis;
        let utterance = null;
        let voices = [];

        function initVoices() {{
            voices = synth.getVoices();
            const voiceSelect = document.getElementById('voice-select');
            voiceSelect.innerHTML = '';
            
            // Lọc ưu tiên giọng Tiếng Việt
            let viVoices = voices.filter(v => v.lang.includes('vi') || v.lang.includes('VI'));
            let displayVoices = viVoices.length > 0 ? viVoices : voices;

            displayVoices.forEach((voice, index) => {{
                const option = document.createElement('option');
                option.textContent = voice.name + ' (' + voice.lang + ')';
                option.setAttribute('data-name', voice.name);
                option.setAttribute('data-lang', voice.lang);
                voiceSelect.appendChild(option);
            }});
        }}

        if (speechSynthesis.onvoiceschanged !== undefined) {{
            speechSynthesis.onvoiceschanged = initVoices;
        }}
        initVoices();

        function playTTS() {{
            synth.cancel();
            utterance = new SpeechSynthesisUtterance(textContent);
            
            const selectedVoiceName = document.getElementById('voice-select').selectedOptions[0]?.getAttribute('data-name');
            const selectedVoice = voices.find(v => v.name === selectedVoiceName);
            if (selectedVoice) {{
                utterance.voice = selectedVoice;
            }}
            
            utterance.rate = parseFloat(document.getElementById('rate').value);
            
            utterance.onstart = function() {{
                document.getElementById('status-badge').innerText = '🔊 Đang đọc...';
                document.getElementById('status-badge').style.background = '#c8e6c9';
                document.getElementById('status-badge').style.color = '#2e7d32';
            }};

            utterance.onend = function() {{
                document.getElementById('status-badge').innerText = '✅ Hoàn thành';
                document.getElementById('status-badge').style.background = '#e0e0e0';
                document.getElementById('status-badge').style.color = '#424242';
            }};

            utterance.onerror = function(event) {{
                document.getElementById('status-badge').innerText = '❌ Lỗi đọc';
            }};

            synth.speak(utterance);
        }}

        function pauseTTS() {{
            if (synth.speaking) {{
                synth.pause();
                document.getElementById('status-badge').innerText = '⏸️ Đã tạm dừng';
            }}
        }}

        function resumeTTS() {{
            if (synth.paused) {{
                synth.resume();
                document.getElementById('status-badge').innerText = '🔊 Đang đọc...';
            }}
        }}

        function stopTTS() {{
            synth.cancel();
            document.getElementById('status-badge').innerText = '⏹️ Đã dừng';
        }}

        function updateRate() {{
            const val = document.getElementById('rate').value;
            document.getElementById('rate-val').innerText = val + 'x';
            if (synth.speaking) {{
                playTTS();
            }}
        }}
    </script>
</body>
</html>
"""

components.html(tts_component_code, height=90)

# -----------------------------------------------------------------------------
# READING DISPLAY AREA
# -----------------------------------------------------------------------------
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    f"""
    <div class="book-reader-container">
        <div class="chapter-title">{current_chapter['title']}</div>
        <div>{"<br><br>".join(current_chapter['content'].split('\n'))}</div>
    </div>
    """,
    unsafe_allow_html=True
)

# -----------------------------------------------------------------------------
# BOOKMARKS & FOOTER
# -----------------------------------------------------------------------------
st.markdown("<br>", unsafe_allow_html=True)
b_col1, b_col2 = st.columns([2, 1])

with b_col1:
    if st.button("🔖 Đánh dấu trang (Bookmark chương này)"):
        bookmark_entry = f"{book['title']} - {current_chapter['title']}"
        if bookmark_entry not in st.session_state.bookmarks:
            st.session_state.bookmarks.append(bookmark_entry)
            st.success("Đã lưu dấu trang thành công!")
        else:
            st.info("Chương này đã có trong danh sách đánh dấu.")

with b_col2:
    if st.session_state.bookmarks:
        with st.expander("Danh sách Đánh dấu (Bookmarks)"):
            for bm in st.session_state.bookmarks:
                st.write(f"• {bm}")

st.markdown("---")
st.caption("🚀 Phát triển bởi Gemini Notebook | Ứng dụng Đọc sách & Chuyển giọng nói AI Tiếng Việt")
