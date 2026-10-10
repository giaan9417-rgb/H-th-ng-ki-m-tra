import streamlit as st
import pandas as pd
import sqlite3
import json
from datetime import datetime
import streamlit.components.v1 as components
from docx import Document

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Hệ Thống Kiểm Tra & Thi Trực Tuyến Thông Minh",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- KHỞI TẠO DATABASE SQLITE ---
def init_db():
    conn = sqlite3.connect("exam_system_v3.db", check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            email TEXT PRIMARY KEY,
            role TEXT,
            name TEXT,
            student_id TEXT,
            stt TEXT,
            class_name TEXT,
            updated_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS exams (
            exam_code TEXT PRIMARY KEY,
            title TEXT,
            class_target TEXT,
            duration INTEGER,
            max_attempts INTEGER,
            questions TEXT,
            created_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS classes (
            class_name TEXT PRIMARY KEY,
            created_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_code TEXT,
            student_email TEXT,
            student_name TEXT,
            student_id TEXT,
            stt TEXT,
            class_name TEXT,
            score REAL,
            violations INTEGER,
            submitted_at TEXT,
            details TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    cursor.execute("SELECT value FROM settings WHERE key = 'admin_password'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO settings (key, value) VALUES ('admin_password', 'admin123')")
    conn.commit()
    return conn

conn = init_db()
cursor = conn.cursor()  # Khởi tạo cursor toàn cục để dùng chung cho toàn bộ ứng dụng

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

# --- HỆ THỐNG XÁC THỰC GOOGLE & ĐĂNG NHẬP THÔNG MINH ---
st.sidebar.title("🔐 Đăng Nhập Hệ Thống")

if "user" not in st.session_state:
    st.session_state.user = None

if st.session_state.user is None:
    if "register_email" not in st.session_state:
        st.sidebar.info("💡 Nhập email Google của bạn để tiếp tục.")
        login_email = st.sidebar.text_input("Email Google (@gmail.com):")
        
        if st.sidebar.button("Tiếp tục"):
            if not login_email or "@" not in login_email:
                st.sidebar.error("❌ Vui lòng nhập địa chỉ email hợp lệ!")
            else:
                cursor.execute("SELECT role, name, student_id, stt, class_name FROM users WHERE email = ?", (login_email,))
                user_row = cursor.fetchone()
                
                if user_row:
                    st.session_state.user = {
                        "email": login_email,
                        "role": user_row[0],
                        "name": user_row[1],
                        "student_id": user_row[2],
                        "stt": user_row[3],
                        "class_name": user_row[4]
                    }
                    st.rerun()
                else:
                    st.session_state.register_email = login_email
                    st.rerun()
    else:
        st.sidebar.subheader(f"Tạo tài khoản mới cho:\n`{st.session_state.register_email}`")
        selected_role = st.sidebar.selectbox("Bạn là ai?", ["Học sinh", "Giáo viên / Quản trị viên"])
        
        admin_pwd_input = ""
        if selected_role == "Giáo viên / Quản trị viên":
            admin_pwd_input = st.sidebar.text_input("Nhập mật khẩu Quản trị (Admin):", type="password")
            
        col_a, col_b = st.sidebar.columns(2)
        with col_a:
            if st.button("Xác nhận"):
                if selected_role == "Giáo viên / Quản trị viên":
                    cursor.execute("SELECT value FROM settings WHERE key = 'admin_password'")
                    correct_pwd = cursor.fetchone()[0]
                    if admin_pwd_input != correct_pwd:
                        st.sidebar.error("❌ Mật khẩu quản trị không đúng!")
                        st.stop()
                        
                cursor.execute("""
                    INSERT INTO users (email, role, name, student_id, stt, class_name, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (st.session_state.register_email, selected_role, "", "", "", "", datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                conn.commit()
                
                st.session_state.user = {
                    "email": st.session_state.register_email,
                    "role": selected_role,
                    "name": "",
                    "student_id": "",
                    "stt": "",
                    "class_name": ""
                }
                del st.session_state.register_email
                st.rerun()
        with col_b:
            if st.button("Quay lại"):
                del st.session_state.register_email
                st.rerun()
    st.stop()
else:
    user = st.session_state.user
    st.sidebar.success(f"👤 Đang đăng nhập:\n**{user['email']}**\nVai trò: **{user['role']}**")
    if st.sidebar.button("Đăng xuất"):
        st.session_state.user = None
        if "register_email" in st.session_state:
            del st.session_state.register_email
        st.rerun()

# ==========================================
# GIAO DIỆN HỌC SINH
# ==========================================
if user["role"] == "Học sinh":
    st.title("🛡️ Cổng Thi Trực Tuyến - Giám Sát Nghiêm Ngặt")
    
    cursor.execute("SELECT class_name FROM classes")
    all_classes = [c[0] for c in cursor.fetchall()]
    
    with st.form("student_profile_form"):
        st.subheader("📝 Xác nhận thông tin cá nhân")
        s_name = st.text_input("Họ và tên học sinh:", value=user.get("name", ""))
        s_id = st.text_input("Mã số học sinh (MSSV):", value=user.get("student_id", ""))
        s_stt = st.text_input("Số thứ tự (STT):", value=user.get("stt", ""))
        
        if all_classes:
            current_cls = user.get("class_name", all_classes[0])
            idx = all_classes.index(current_cls) if current_cls in all_classes else 0
            s_class = st.selectbox("Chọn lớp của bạn:", all_classes, index=idx)
        else:
            s_class = st.text_input("Nhập tên lớp của bạn:", value=user.get("class_name", ""))
            
        update_profile = st.form_submit_button("💾 Lưu thông tin & Vào phòng thi")
        
        if update_profile:
            if not s_name or not s_id or not s_stt or not s_class:
                st.warning("⚠️ Vui lòng điền đầy đủ thông tin!")
            else:
                cursor.execute("""
                    UPDATE users SET name = ?, student_id = ?, stt = ?, class_name = ?, updated_at = ?
                    WHERE email = ?
                """, (s_name, s_id, s_stt, s_class, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user["email"]))
                conn.commit()
                user["name"] = s_name
                user["student_id"] = s_id
                user["stt"] = s_stt
                user["class_name"] = s_class
                st.success("✅ Đã lưu thông tin!")
                st.rerun()
                
    if not user["name"] or not user["class_name"]:
        st.info("💡 Vui lòng hoàn tất thông tin cá nhân ở trên để tiếp tục.")
        st.stop()
        
    st.markdown("---")
    st.subheader(f"📚 Xin chào {user['name']} (Lớp: {user['class_name']} - STT: {user['stt']})")
    
    cursor.execute("SELECT exam_code, title, duration, max_attempts FROM exams WHERE class_target = ? OR class_target = 'Tất cả'", (user["class_name"],))
    available_exams = cursor.fetchall()
    
    if not available_exams:
        st.info("📭 Hiện chưa có bài kiểm tra nào được giao cho lớp của bạn.")
        st.stop()
        
    exam_options = {f"{e[1]} (Mã đề: {e[0]} - Thời gian: {e[2]} phút)": e for e in available_exams}
    chosen_exam_label = st.selectbox("Chọn bài kiểm tra:", list(exam_options.keys()))
    exam_info = exam_options[chosen_exam_label]
    exam_code_input, exam_title, exam_duration, max_attempts = exam_info[0], exam_info[1], exam_info[2], exam_info[3]
    
    cursor.execute("SELECT COUNT(*) FROM submissions WHERE exam_code = ? AND student_email = ?", (exam_code_input, user["email"]))
    attempts_made = cursor.fetchone()[0]
    
    st.markdown(f"**Số lần đã làm:** `{attempts_made} / {max_attempts}`")
    if attempts_made >= max_attempts:
        st.error("❌ Bạn đã hết số lần làm bài cho phép đối với đề thi này!")
        st.stop()
        
    if st.button("🚀 Bắt Đầu Làm Bài Ngay"):
        st.session_state.active_exam = exam_code_input
        st.rerun()
        
    if "active_exam" in st.session_state and st.session_state.active_exam == exam_code_input:
        cursor.execute("SELECT questions FROM exams WHERE exam_code = ?", (exam_code_input,))
        q_row = cursor.fetchone()
        questions = json.loads(q_row[0])
        
        fullscreen_script = f"""
        <div style="padding: 15px; background: #ffebee; border-left: 6px solid #d32f2f; margin-bottom: 15px; border-radius: 4px;">
            <h3 style="margin: 0; color: #b71c1c;">🚨 PHÒNG THI AN NINH CAO</h3>
            <p style="margin: 5px 0 0 0; font-size: 14px; color: #c62828;">
                Hệ thống đang ở chế độ <b>Toàn màn hình</b>. Nghiêm cấm chuyển tab hoặc rời khỏi trang!
            </p>
        </div>
        <script>
        function triggerFullscreen() {{
            let elem = window.parent.document.documentElement;
            if (elem.requestFullscreen) {{ elem.requestFullscreen(); }}
        }}
        let violationCount = 0;
        document.addEventListener("visibilitychange", function() {{
            if (document.hidden) {{
                violationCount++;
                alert("⚠️ CẢNH BÁO LẦN " + violationCount + ": Bạn vừa chuyển tab!");
            }}
        }});
        document.addEventListener("contextmenu", e => e.preventDefault());
        document.addEventListener("keydown", function(e) {{
            if (e.ctrlKey && (e.key === 'c' || e.key === 'v' || e.key === 'x' || e.key === 'u' || e.key === 'p')) {{
                e.preventDefault();
                alert("⚠️ Phím tắt bị chặn!");
            }}
        }});
        </script>
        <button onclick="triggerFullscreen()" style="padding: 10px 20px; background: #1976d2; color: white; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; margin-bottom: 20px;">
            🔲 Bật Toàn Màn Hình Thi Cử
        </button>
        """
        components.html(fullscreen_script, height=165)
        
        st.subheader(f"📖 Đề thi: {exam_title} (Thời gian: {exam_duration} phút)")
        
        with st.form("exam_submission_form"):
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
                
            submitted_exam = st.form_submit_button("🚀 Nộp Bài Thi Chính Thức")
            
        if submitted_exam:
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
                INSERT INTO submissions (exam_code, student_email, student_name, student_id, stt, class_name, score, violations, submitted_at, details)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                exam_code_input,
                user["email"],
                user["name"],
                user["student_id"],
                user["stt"],
                user["class_name"],
                total_score,
                0,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                json.dumps(details, ensure_ascii=False)
            ))
            conn.commit()
            del st.session_state.active_exam
            st.success(f"🎉 Nộp bài thành công! Tổng điểm của bạn: **{total_score:.2f} điểm**")
            st.balloons()

# ==========================================
# GIAO DIỆN GIÁO VIÊN / QUẢN TRỊ
# ==========================================
elif user["role"] == "Giáo viên / Quản trị viên":
    st.title("⚙️ Cổng Quản Lý Giáo Viên & Kho Đề Thi")
    
    t_tab1, t_tab2, t_tab3, t_tab4 = st.tabs([
        "📁 Kho Đề Thi & Tải Lên Word", 
        "🏫 Quản Lý Lớp Học & Giao Đề", 
        "📊 Bảng Điểm & Xuất Excel", 
        "🔐 Cài Đặt Mật Khẩu Admin"
    ])
    
    with t_tab1:
        st.subheader("1. Tải Lên Đề Thi Vào Kho Đề")
        cursor.execute("SELECT class_name FROM classes")
        cls_list = [c[0] for c in cursor.fetchall()]
        cls_choices = ["Tất cả"] + cls_list
        
        up_code = st.text_input("Mã đề thi (VD: LS11_01):")
        up_title = st.text_input("Tiêu đề bài kiểm tra:")
        up_class = st.selectbox("Giao cho lớp:", cls_choices)
        up_duration = st.number_input("Thời gian làm bài (phút):", min_value=5, max_value=180, value=45)
        up_max_attempts = st.number_input("Số lần làm bài tối đa:", min_value=1, max_value=10, value=1)
        up_file = st.file_uploader("Chọn file Word đề thi (.docx):", type=["docx"])
        
        if up_file and st.button("Lưu Đề Thi Vào Kho"):
            try:
                parsed_q = parse_docx_to_questions(up_file)
                if not parsed_q:
                    st.error("⚠️ Không đọc được câu hỏi từ file Word.")
                else:
                    cursor.execute("""
                        INSERT OR REPLACE INTO exams (exam_code, title, class_target, duration, max_attempts, questions, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (
                        up_code, up_title, up_class, up_duration, up_max_attempts,
                        json.dumps(parsed_q, ensure_ascii=False),
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ))
                    conn.commit()
                    st.success(f"✅ Đã thêm đề thi `{up_code}` vào kho đề thành công!")
            except Exception as e:
                st.error(f"Lỗi xử lý file: {e}")
                
        st.markdown("---")
        st.subheader("📚 Kho Đề Thi Hiện Có")
        cursor.execute("SELECT exam_code, title, class_target, duration, max_attempts, created_at FROM exams")
        exams_db = cursor.fetchall()
        if exams_db:
            df_exams = pd.DataFrame(exams_db, columns=["Mã đề", "Tiêu đề", "Lớp áp dụng", "Thời gian (phút)", "Số lần thi", "Ngày tạo"])
            st.dataframe(df_exams, use_container_width=True)
        else:
            st.info("Kho đề thi trống.")

    with t_tab2:
        st.subheader("2. Quản Lý Lớp Học")
        new_class_name = st.text_input("Tên lớp mới (VD: 11A1, 11A2):")
        if st.button("Thêm Lớp Học"):
            if new_class_name:
                try:
                    cursor.execute("INSERT INTO classes (class_name, created_at) VALUES (?, ?)", (new_class_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    conn.commit()
                    st.success(f"✅ Đã thêm lớp `{new_class_name}` thành công!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("❌ Lớp học này đã tồn tại.")
                    
        cursor.execute("SELECT class_name, created_at FROM classes")
        cls_db = cursor.fetchall()
        if cls_db:
            st.write("**Danh sách các lớp:**")
            for c in cls_db:
                st.markdown(f"- Lớp: **{c[0]}** (Tạo lúc: {c[1]})")

    with t_tab3:
        st.subheader("3. Bảng Điểm & Xuất File Excel")
        cursor.execute("SELECT id, exam_code, student_name, student_id, stt, class_name, score, violations, submitted_at FROM submissions")
        subs = cursor.fetchall()
        if subs:
            df_subs = pd.DataFrame(subs, columns=["ID", "Mã đề", "Họ tên", "MSSV", "STT", "Lớp", "Điểm số", "Vi phạm", "Thời gian nộp"])
            st.dataframe(df_subs, use_container_width=True)
            csv_bytes = df_subs.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Tải Xuống Bảng Điểm (CSV/Excel)", data=csv_bytes, file_name="bang_diem_chi_tiet.csv", mime="text/csv")
        else:
            st.info("Chưa có học sinh nào nộp bài.")

    with t_tab4:
        st.subheader("4. Đổi Mật Khẩu Quản Trị (Admin)")
        new_pwd = st.text_input("Mật khẩu admin mới:", type="password")
        if st.button("Cập Nhật Mật Khẩu"):
            if new_pwd:
                cursor.execute("UPDATE settings SET value = ? WHERE key = 'admin_password'", (new_pwd,))
                conn.commit()
                st.success("✅ Đã cập nhật mật khẩu admin thành công! Hệ thống sẽ ghi nhớ cho lần sau.")
