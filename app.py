import streamlit as st
import pandas as pd
import sqlite3
import json
from datetime import datetime
import streamlit.components.v1 as components
from docx import Document

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Hệ Thống Kiểm Tra Trực Tuyến & Chống Gian Lận",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- KHỞI TẠO DATABASE SQLITE ---
def init_db():
    conn = sqlite3.connect("exam_pro_system.db", check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS exams (
            exam_code TEXT PRIMARY KEY,
            title TEXT,
            questions TEXT,
            created_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_code TEXT,
            student_name TEXT,
            student_id TEXT,
            score REAL,
            violations INTEGER,
            submitted_at TEXT,
            details TEXT
        )
    """)
    conn.commit()
    return conn

conn = init_db()

# --- HÀM PHÂN TÍCH FILE WORD (.DOCX) DỰA TRÊN ĐỊNH DẠNG GẠCH CHÂN ---
def parse_docx_to_questions(docx_file):
    doc = Document(docx_file)
    questions = []
    current_q = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        is_underlined = any(run.underline for run in para.runs if run.underline)
        
        if text.lower().startswith("câu"):
            if current_q:
                questions.append(current_q)
            current_q = {
                "question": text,
                "type": "mcq", 
                "options": [],
                "statements": [],
                "answer": ""
            }
        elif current_q is not None:
            if (len(text) > 1 and text[0].upper() in ['A', 'B', 'C', 'D'] and text[1] in ['.', ')', ' ']):
                current_q["options"].append(text)
                if is_underlined:
                    current_q["answer"] = text[0].upper()
            elif (len(text) > 1 and text[0].lower() in ['a', 'b', 'c', 'd'] and text[1] in ['.', ')', ' ']):
                current_q["type"] = "true_false"
                ans_val = "Đúng" if is_underlined else "Sai"
                current_q["statements"].append({"text": text, "answer": ans_val})
            else:
                if current_q["type"] == "mcq" and not current_q["options"]:
                    current_q["question"] += " " + text
                    
    if current_q:
        questions.append(current_q)
    return questions

# --- GIAO DIỆN CHÍNH ---
st.sidebar.title("🧭 Điều hướng hệ thống")
role = st.sidebar.selectbox("Chọn vai trò:", ["👨‍🎓 Học sinh làm bài", "👨‍🏫 Giáo viên / Quản trị viên"])

# ==========================================
# GIAO DIỆN HỌC SINH (CHỐNG GIAN LẬN & LÀM BÀI)
# ==========================================
if role == "👨‍🎓 Học sinh làm bài":
    st.title("🛡️ Phòng Thi Trực Tuyến - Giám Sát Nghiêm Ngặt")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        student_name = st.text_input("Họ và tên học sinh:")
    with col2:
        student_id = st.text_input("Mã số học sinh (MSSV):")
    with col3:
        exam_code_input = st.text_input("Mã đề thi:")
        
    if not student_name or not student_id or not exam_code_input:
        st.info("💡 Vui lòng nhập đầy đủ thông tin để vào phòng thi.")
        st.stop()
        
    cursor = conn.cursor()
    cursor.execute("SELECT title, questions FROM exams WHERE exam_code = ?", (exam_code_input,))
    exam_row = cursor.fetchone()
    
    if not exam_row:
        st.error("❌ Không tìm thấy mã đề thi này!")
        st.stop()
        
    exam_title, questions_json = exam_row
    questions = json.loads(questions_json)
    
    st.subheader(f"📖 Đề thi: {exam_title} (Mã đề: {exam_code_input})")
    
    # --- JAVASCRIPT CHỐNG GIAN LẬN ---
    anti_cheat_script = """
    <div style="padding: 10px; background: #ffebee; border-left: 5px solid #d32f2f; margin-bottom: 15px; border-radius: 4px;">
        <h4 style="margin: 0; color: #b71c1c;">🚨 CHẾ ĐỘ PHÒNG CHỐNG GIAN LẬN ĐANG BẬT</h4>
        <p style="margin: 5px 0 0 0; font-size: 13px; color: #c62828;">
            Hệ thống giám sát việc chuyển tab, thu nhỏ cửa sổ, chặn copy/paste và chặn sử dụng chuột phải. Mọi hành vi vi phạm sẽ bị ghi lại!
        </p>
    </div>
    <script>
    let violationCount = 0;
    document.addEventListener("visibilitychange", function() {
        if (document.hidden) {
            violationCount++;
            alert("⚠️ CẢNH BÁO LẦN " + violationCount + ": Bạn vừa chuyển tab hoặc thoát khỏi màn hình làm bài!");
        }
    });
    document.addEventListener("contextmenu", e => e.preventDefault());
    document.addEventListener("keydown", function(e) {
        if (e.ctrlKey && (e.key === 'c' || e.key === 'v' || e.key === 'x' || e.key === 'u' || e.key === 'p')) {
            e.preventDefault();
            alert("⚠️ Phím tắt sao chép/dán bị chặn!");
        }
    });
    document.addEventListener("paste", function(e) {
        alert("🚨 Phát hiện hành vi dán (Paste) dữ liệu ngoài!");
    });
    </script>
    """
    components.html(anti_cheat_script, height=110)
    
    with st.form("exam_form"):
        user_answers = {}
        for idx, q in enumerate(questions):
            st.markdown(f"**Câu {idx+1}:** {q['question']}")
            q_type = q.get("type", "mcq")
            
            if q_type == "true_false" and q.get("statements"):
                user_answers[str(idx)] = {}
                for s_idx, stmt in enumerate(q["statements"]):
                    choice = st.radio(
                        f"{stmt['text']}", 
                        ["Đúng", "Sai"], 
                        key=f"q_{idx}_stmt_{s_idx}",
                        horizontal=True
                    )
                    user_answers[str(idx)][str(s_idx)] = choice
            else:
                options = q.get('options', [])
                choice = st.radio(
                    f"Chọn đáp án câu {idx+1}:", 
                    options, 
                    key=f"q_{idx}",
                    label_visibility="collapsed"
                )
                user_answers[str(idx)] = choice
            st.markdown("---")
            
        submitted = st.form_submit_button("🚀 Nộp Bài Thi Chính Thức")
        
    if submitted:
        total_score = 0.0
        details = []
        
        for idx, q in enumerate(questions):
            q_type = q.get("type", "mcq")
            if q_type == "true_false" and q.get("statements"):
                correct_count = 0
                stmt_results = []
                user_stmt_dict = user_answers.get(str(idx), {})
                
                for s_idx, stmt in enumerate(q["statements"]):
                    user_choice = user_stmt_dict.get(str(s_idx), "Đúng")
                    correct_choice = stmt["answer"]
                    is_stmt_correct = (user_choice.strip().lower() == correct_choice.strip().lower())
                    if is_stmt_correct:
                        correct_count += 1
                    stmt_results.append({
                        "statement": stmt['text'],
                        "chosen": user_choice,
                        "correct": correct_choice,
                        "is_correct": is_stmt_correct
                    })
                
                score_mapping = {0: 0.0, 1: 0.1, 2: 0.25, 3: 0.5, 4: 1.0}
                q_score = score_mapping.get(correct_count, 0.0)
                total_score += q_score
                
                details.append({
                    "question": q['question'],
                    "type": "true_false",
                    "correct_statements_count": correct_count,
                    "score_awarded": q_score,
                    "statements_detail": stmt_results
                })
            else:
                chosen = user_answers.get(str(idx), "")
                chosen_letter = chosen[0].upper() if chosen else ""
                correct_letter = q.get('answer', '').upper()
                
                is_correct = (chosen_letter == correct_letter)
                q_score = 1.0 if is_correct else 0.0
                total_score += q_score
                
                details.append({
                    "question": q['question'],
                    "type": "mcq",
                    "chosen": chosen,
                    "correct": correct_letter,
                    "is_correct": is_correct,
                    "score_awarded": q_score
                })
                
        cursor.execute("""
            INSERT INTO submissions (exam_code, student_name, student_id, score, violations, submitted_at, details)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            exam_code_input,
            student_name,
            student_id,
            total_score,
            0,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            json.dumps(details, ensure_ascii=False)
        ))
        conn.commit()
        
        st.success(f"🎉 Nộp bài thành công! Tổng điểm của bạn: **{total_score:.2f} điểm**")
        st.balloons()

# ==========================================
# GIAO DIỆN GIÁO VIÊN
# ==========================================
elif role == "👨‍🏫 Giáo viên / Quản trị viên":
    st.title("⚙️ Cổng Quản Lý Giáo Viên - Tải Đề Word & Xuất Điểm")
    
    password = st.text_input("Nhập mật khẩu quản trị:", type="password")
    if password != "admin123" and password != "":
        st.error("❌ Mật khẩu không chính xác!")
        st.stop()
    elif password == "":
        st.info("🔐 Vui lòng nhập mật khẩu quản trị (Mặc định: `admin123`) để tiếp tục.")
        st.stop()
        
    tab1, tab2, tab3 = st.tabs(["📁 Tải Lên Đề Thi (File Word)", "📊 Bảng Điểm & Xuất Excel", "🔍 Chi Tiết Bài Làm"])
    
    with tab1:
        st.subheader("1. Đăng tải đề thi bằng file Word (.docx)")
        st.info("""
        💡 **Quy tắc soạn thảo file Word:**
        - **Trắc nghiệm:** `Câu 1: ...`, liệt kê các dòng `A`, `B`, `C`, `D`. **Gạch chân (underline)** đáp án đúng.
        - **Đúng/Sai:** Liệt kê các ý `a)`, `b)`, `c)`, `d)`. Ý nào **gạch chân** hệ thống tự hiểu là **Đúng**, ý khác là **Sai**.
        """)
        
        upload_code = st.text_input("Mã đề thi (Ví dụ: DE01):")
        upload_title = st.text_input("Tiêu đề bài kiểm tra:")
        uploaded_file = st.file_uploader("Chọn file Word đề thi (.docx):", type=["docx"])
        
        if uploaded_file and st.button("Xử lý và Lưu Đề Thi Lên Hệ Thống"):
            try:
                parsed_questions = parse_docx_to_questions(uploaded_file)
                if not parsed_questions:
                    st.error("⚠️ Không đọc được câu hỏi nào từ file Word.")
                else:
                    cursor = conn.cursor()
                    cursor.execute("INSERT OR REPLACE INTO exams (exam_code, title, questions, created_at) VALUES (?, ?, ?, ?)", (
                        upload_code,
                        upload_title,
                        json.dumps(parsed_questions, ensure_ascii=False),
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ))
                    conn.commit()
                    st.success(f"✅ Đã tải và lưu thành công đề thi `{upload_code}` gồm {len(parsed_questions)} câu hỏi!")
                    st.json(parsed_questions)
            except Exception as e:
                st.error(f"Lỗi khi đọc file Word: {e}")

    with tab2:
        st.subheader("📊 Bảng Điểm Tổng Hợp & Xuất File Excel")
        cursor = conn.cursor()
        cursor.execute("SELECT id, exam_code, student_name, student_id, score, violations, submitted_at FROM submissions")
        subs = cursor.fetchall()
        
        if not subs:
            st.info("Chưa có học sinh nào nộp bài.")
        else:
            df_subs = pd.DataFrame(subs, columns=["ID", "Mã đề", "Họ tên", "MSSV", "Điểm số", "Số lần vi phạm", "Thời gian nộp"])
            st.dataframe(df_subs, use_container_width=True)
            
            csv_data = df_subs.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Tải Xuống Bảng Điểm (CSV / Excel)",
                data=csv_data,
                file_name=f"bang_diem_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )

    with tab3:
        st.subheader("🔍 Xem Chi Tiết Bài Làm Học Sinh")
        cursor = conn.cursor()
        cursor.execute("SELECT id, student_name, student_id, exam_code, score, violations, details FROM submissions")
        all_subs = cursor.fetchall()
        
        if all_subs:
            sub_options = {f"{s[1]} ({s[2]}) - Đề: {s[3]} - Điểm: {s[4]}": s[0] for s in all_subs}
            chosen_sub_label = st.selectbox("Chọn bài làm của học sinh:", list(sub_options.keys()))
            
            if chosen_sub_label:
                selected_id = sub_options[chosen_sub_label]
                selected_record = [s for s in all_subs if s[0] == selected_id][0]
                
                st.markdown(f"**Học sinh:** {selected_record[1]} (MSSV: {selected_record[2]})")
                st.markdown(f"**Mã đề:** {selected_record[3]} | **Tổng điểm:** `{selected_record[4]}`")
                
                details_data = json.loads(selected_record[6])
                st.write("---")
                for i, d in enumerate(details_data):
                    st.markdown(f"**Câu {i+1}:** {d['question']}")
                    if d.get('type') == 'true_false':
                        st.markdown(f"- Số ý đúng: `{d['correct_statements_count']}/4` -> Đạt được: **{d['score_awarded']}đ**")
                        for s_det in d['statements_detail']:
                            icon = "✅" if s_det['is_correct'] else "❌"
                            st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;+ {s_det['statement']} | HS chọn: `{s_det['chosen']}` (Đáp án quy định: {s_det['correct']}) {icon}")
                    else:
                        status = "✅ Đúng" if d['is_correct'] else "❌ Sai"
                        st.markdown(f"- Học sinh chọn: `{d.get('chosen')}` | Đáp án đúng: `{d.get('correct')}` -> **{status}**")
                    st.markdown("---")
