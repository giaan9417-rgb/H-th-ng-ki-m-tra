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

# Custom CSS chống bôi đen và quét OCR cơ bản
st.markdown("""
<style>
    body, div, p, span, h1, h2, h3, label {
        -webkit-user-select: none !important;
        -moz-user-select: none !important;
        -ms-user-select: none !important;
        user-select: none !important;
    }
</style>
""", unsafe_allow_html=True)

# --- KHỞI TẠO DATABASE SQLITE (WAL MODE CHO 35+ HỌC SINH ĐỒNG THỜI) ---
def init_db():
    conn = sqlite3.connect("exam_system_v9.db", check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("PRAGMA journal_mode=WAL;")
    
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
        CREATE TABLE IF NOT EXISTS classes (
            class_name TEXT PRIMARY KEY,
            created_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS exams (
            exam_code TEXT PRIMARY KEY,
            title TEXT,
            class_target TEXT,
            duration INTEGER,
            max_attempts INTEGER,
            show_score INTEGER,
            questions TEXT,
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
            status TEXT,
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
cursor = conn.cursor()

# --- HÀM PHÂN TÍCH WORD (PHÂN BIỆT A,B,C,D VÀ a,b,c,d) ---
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
                "answer": "",
                "points": 1.0
            }
        elif current_q is not None:
            if (len(text) > 1 and text[0] in ['A', 'B', 'C', 'D'] and text[1] in ['.', ')', ' ']):
                current_q["options"].append(text)
                if is_underlined:
                    current_q["answer"] = text[0]
            elif (len(text) > 1 and text[0] in ['a', 'b', 'c', 'd'] and text[1] in ['.', ')', ' ']):
                current_q["type"] = "true_false"
                ans_val = "Đúng" if is_underlined else "Sai"
                current_q["statements"].append({"text": text, "answer": ans_val})
            else:
                if current_q["type"] == "mcq" and not current_q["options"]:
                    current_q["question"] += " " + text
                    
    if current_q:
        questions.append(current_q)
    return questions

# --- XỬ LÝ KHI HỌC SINH VI PHẠM ---
if st.query_params.get("violated") == "true":
    if "user" in st.session_state and st.session_state.user and st.session_state.user["role"] == "Học sinh":
        u = st.session_state.user
        e_code = st.session_state.get("active_exam", "UNKNOWN")
        cursor.execute("""
            INSERT INTO submissions (exam_code, student_email, student_name, student_id, stt, class_name, score, violations, status, submitted_at, details)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            e_code, u["email"], u["name"], u["student_id"], u["stt"], u["class_name"],
            0.0, 1, "⚠️ Vi phạm (Rời tab/Chia đôi màn hình/AI)", datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            json.dumps([{"error": "Học sinh rời khỏi tab, chia đôi màn hình hoặc mở cửa sổ nổi AI"}], ensure_ascii=False)
        ))
        conn.commit()
        for key in ["active_exam", "active_exam_title", "active_exam_duration", "active_exam_questions"]:
            if key in st.session_state:
                del st.session_state[key]
    st.error("🚨 BẠN ĐÃ VI PHẠM QUY CHẾ THI! Hành vi rời tab, chia đôi màn hình hoặc mở trợ lý AI đã bị ghi nhận. Bài làm đã bị khóa và báo cáo về cho giáo viên.")
    if st.button("Quay lại màn hình chính"):
        st.query_params.clear()
        st.rerun()
    st.stop()

# --- XÁC THỰC & ĐĂNG NHẬP ---
st.sidebar.title("🔐 Đăng Nhập Hệ Thống")

if "user" not in st.session_state:
    st.session_state.user = None

if st.session_state.user is None:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); padding: 30px; border-radius: 12px; color: white; text-align: center; margin-bottom: 25px;">
        <h1 style="color: white; margin-bottom: 10px;">🎓 HỆ THỐNG KIỂM TRA TRỰC TUYẾN THÔNG MINH</h1>
        <p style="font-size: 16px; opacity: 0.9;">Bảo mật tuyệt đối - Khóa bài tức thì khi gian lận - Quét đề tự động</p>
    </div>
    """, unsafe_allow_html=True)
    
    if "register_email" not in st.session_state:
        st.sidebar.info("💡 Nhập Email Google của bạn để tiếp tục.")
        login_email = st.sidebar.text_input("Email Google (@gmail.com):")
        
        if st.sidebar.button("Đăng Nhập / Tiếp Tục"):
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
    st.sidebar.success(f"👤 Tài khoản:\n**{user['email']}**\nVai trò: **{user['role']}**")
    if st.sidebar.button("Đăng xuất"):
        st.session_state.user = None
        if "register_email" in st.session_state:
            del st.session_state.register_email
        st.rerun()

# ==========================================
# GIAO DIỆN HỌC SINH
# ==========================================
if user["role"] == "Học sinh":
    st.title("🛡️ Cổng Thi Trực Tuyến - Giám Sát An Ninh")
    
    cursor.execute("SELECT class_name FROM classes")
    all_classes = [c[0] for c in cursor.fetchall()]
    
    with st.form("student_profile_form"):
        st.subheader("📝 Xác nhận thông tin học sinh")
        s_name = st.text_input("Họ và tên học sinh:", value=user.get("name", ""))
        s_id = st.text_input("Mã số học sinh (MSSV):", value=user.get("student_id", ""))
        s_stt = st.text_input("Số thứ tự (STT):", value=user.get("stt", ""))
        
        if all_classes:
            current_cls = user.get("class_name", all_classes[0])
            idx = all_classes.index(current_cls) if current_cls in all_classes else 0
            s_class = st.selectbox("Chọn lớp của bạn:", all_classes, index=idx)
        else:
            s_class = st.text_input("Nhập tên lớp của bạn:", value=user.get("class_name", ""))
            
        update_profile = st.form_submit_button("💾 Lưu thông tin & Tiếp tục")
        
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
        st.info("💡 Vui lòng hoàn tất thông tin cá nhân ở trên để xem bài thi.")
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
        cursor.execute("SELECT questions FROM exams WHERE exam_code = ?", (exam_code_input,))
        q_row = cursor.fetchone()
        if q_row:
            st.session_state.active_exam = exam_code_input
            st.session_state.active_exam_title = exam_title
            st.session_state.active_exam_duration = exam_duration
            st.session_state.active_exam_questions = json.loads(q_row[0])
            st.rerun()
        else:
            st.error("❌ Không tìm thấy dữ liệu đề thi này trong hệ thống!")
        
    if "active_exam" in st.session_state:
        exam_title = st.session_state.active_exam_title
        exam_duration = st.session_state.active_exam_duration
        questions = st.session_state.active_exam_questions
        
        # --- BẢO MẬT NÂNG CAO: PHÁT ÂM THÀNH CẢNH BÁO + KHÓA MÀN HÌNH TỨC THÌ KHI GIAN LẬN ---
        security_script = """
        <div style="padding: 12px; background: #ffebee; border-left: 6px solid #d32f2f; margin-bottom: 15px; border-radius: 4px;">
            <h4 style="margin: 0; color: #b71c1c;">🚨 PHÒNG THI AN NINH CAO - CHỐNG ĐA NHIỆM & AI</h4>
            <p style="margin: 5px 0 0 0; font-size: 13px; color: #c62828;">
                Bắt buộc Toàn màn hình. Cấm chia đôi màn hình, mở cửa sổ nổi AI (Gemini) hoặc rời tab. Vi phạm sẽ hú còi cảnh báo và KHÓA BÀI NGAY!
            </p>
        </div>
        <script>
        function forceFullscreen() {
            let elem = window.parent.document.documentElement;
            if (elem.requestFullscreen) {
                elem.requestFullscreen().catch(err => { console.log(err); });
            }
        }
        window.onload = forceFullscreen;
        document.addEventListener("click", forceFullscreen);

        let isLocked = false;

        // Hàm phát âm thanh cảnh báo (Web Audio API)
        function playAlarmSound() {
            try {
                let ctx = new (window.AudioContext || window.webkitAudioContext)();
                let osc = ctx.createOscillator();
                let gain = ctx.createGain();
                osc.type = 'sawtooth';
                osc.frequency.setValueAtTime(800, ctx.currentTime);
                osc.frequency.exponentialRampToValueAtTime(300, ctx.currentTime + 1.0);
                gain.gain.setValueAtTime(1, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 1.0);
                osc.connect(gain);
                gain.connect(ctx.destination);
                osc.start();
                osc.stop(ctx.currentTime + 1.0);
            } catch(e) {}
        }

        // Hàm khóa màn hình tức thì
        function triggerLockdown() {
            if (isLocked) return;
            isLocked = true;
            playAlarmSound();

            // Tạo khung đỏ phủ kín toàn bộ trang web ngay lập tức
            let overlay = window.parent.document.createElement('div');
            overlay.style.position = 'fixed';
            overlay.style.top = '0';
            overlay.style.left = '0';
            overlay.style.width = '100vw';
            overlay.style.height = '100vh';
            overlay.style.backgroundColor = '#b71c1c';
            overlay.style.color = 'white';
            overlay.style.zIndex = '9999999';
            overlay.style.display = 'flex';
            overlay.style.flexDirection = 'column';
            overlay.style.justifyContent = 'center';
            overlay.style.alignItems = 'center';
            overlay.style.fontFamily = 'sans-serif';
            overlay.innerHTML = '<h1 style="font-size: 2.5rem; margin-bottom: 20px; text-align: center;">🚨 PHÁT HIỆN GIAN LẬN / RỜI BÀI THI!</h1><p style="font-size: 1.3rem; text-align: center; padding: 0 20px;">Hệ thống đã khóa bài thi và ghi nhận vi phạm về cho giáo viên.</p>';
            window.parent.document.body.appendChild(overlay);

            setTimeout(function() {
                let currentUrl = window.parent.location.href.split('?')[0];
                window.parent.location.href = currentUrl + "?violated=true";
            }, 1800);
        }

        // 1. Phát hiện mất tiêu điểm (blur) hoặc chuyển tab (visibilitychange)
        window.addEventListener("blur", function() {
            triggerLockdown();
        });

        document.addEventListener("visibilitychange", function() {
            if (document.hidden) {
                triggerLockdown();
            }
        });

        // 2. Phát hiện thay đổi kích thước khung hình đột ngột (chia đôi màn hình / mở app nổi AI)
        let initialWidth = window.innerWidth;
        window.addEventListener("resize", function() {
            if (Math.abs(window.innerWidth - initialWidth) > 150) {
                triggerLockdown();
            }
        });

        document.addEventListener("contextmenu", e => e.preventDefault());
        document.addEventListener("keydown", function(e) {
            if (e.ctrlKey || e.metaKey || e.key === 'F12') {
                e.preventDefault();
            }
        });
        </script>
        """
        components.html(security_script, height=130)
        
        st.subheader(f"📖 Đề thi: {exam_title} (Thời gian: {exam_duration} phút)")
        
        with st.form("exam_submission_form"):
            user_answers = {}
            for idx, q in enumerate(questions):
                st.markdown(f"**Câu {idx+1}** ({q.get('points', 1.0)} điểm): {q['question']}")
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
                q_pts = q.get("points", 1.0)
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
                    fraction_mapping = {0: 0.0, 1: 0.1, 2: 0.25, 3: 0.5, 4: 1.0}
                    q_score = fraction_mapping.get(correct_count, 0.0) * q_pts
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
                    q_score = q_pts if is_correct else 0.0
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
                INSERT INTO submissions (exam_code, student_email, student_name, student_id, stt, class_name, score, violations, status, submitted_at, details)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                st.session_state.active_exam,
                user["email"],
                user["name"],
                user["student_id"],
                user["stt"],
                user["class_name"],
                total_score,
                0,
                "Hoàn thành",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                json.dumps(details, ensure_ascii=False)
            ))
            conn.commit()
            
            cursor.execute("SELECT show_score FROM exams WHERE exam_code = ?", (st.session_state.active_exam,))
            show_row = cursor.fetchone()
            show_score_setting = show_row[0] if show_row else 1
            
            for key in ["active_exam", "active_exam_title", "active_exam_duration", "active_exam_questions"]:
                if key in st.session_state:
                    del st.session_state[key]
            
            exit_script = "<script>if (document.fullscreenElement) { document.exitFullscreen(); }</script>"
            components.html(exit_script, height=0)
            
            if show_score_setting == 1:
                st.success(f"🎉 Nộp bài thành công! Điểm của bạn: **{total_score:.2f} điểm**")
            else:
                st.success("🎉 Nộp bài thành công! Bài làm đã được lưu.")
            st.balloons()
            st.rerun()

# ==========================================
# GIAO DIỆN GIÁO VIÊN / QUẢN TRỊ
# ==========================================
elif user["role"] == "Giáo viên / Quản trị viên":
    st.title("⚙️ Cổng Quản Lý Giáo Viên & Kho Đề Thi")
    
    t_tab1, t_tab2, t_tab3, t_tab4, t_tab5 = st.tabs([
        "📁 Kho Đề Thi & Tải Lên", 
        "🏫 Quản Lý Lớp & Học Sinh", 
        "📊 Bảng Điểm & Xuất Excel", 
        "🚨 Thống Kê Vi Phạm", 
        "🔐 Đổi Mật Khẩu Admin"
    ])
    
    with t_tab1:
        st.subheader("1. Hướng Dẫn Định Dạng File Word (.docx)")
        with st.expander("📌 Xem cấu trúc mẫu file chuẩn để hệ thống quét đúng 100%"):
            st.markdown("""
            **- Trắc nghiệm nhiều lựa chọn (Sử dụng chữ HOA A, B, C, D):**
            > Câu 1: Nội dung câu hỏi...
            > A. Lựa chọn 1
            > B. Lựa chọn 2 *(Gạch chân đáp án đúng)*
            > C. Lựa chọn 3
            > D. Lựa chọn 4
            
            **- Đúng / Sai (Sử dụng chữ thường a, b, c, d):**
            > Câu 2: Nội dung câu hỏi...
            > a) Phát biểu ý a... *(Gạch chân nếu là Đúng)*
            > b) Phát biểu ý b...
            > c) Phát biểu ý c...
            > d) Phát biểu ý d...
            """)
            
        st.markdown("---")
        st.subheader("2. Tải Lên & Tự Động Quét Đề Thi")
        cursor.execute("SELECT class_name FROM classes")
        cls_list = [c[0] for c in cursor.fetchall()]
        cls_choices = ["Tất cả"] + cls_list
        
        up_code = st.text_input("Mã đề thi (VD: LS11_01):")
        up_title = st.text_input("Tiêu đề bài kiểm tra:")
        up_class = st.selectbox("Giao cho lớp:", cls_choices)
        up_duration = st.number_input("Thời gian làm bài (phút):", min_value=5, max_value=180, value=45)
        up_max_attempts = st.number_input("Số lần làm bài tối đa:", min_value=1, max_value=10, value=1)
        
        col_pts, col_show = st.columns(2)
        with col_pts:
            default_q_pts = st.number_input("Điểm mặc định mỗi câu:", min_value=0.1, max_value=10.0, value=1.0)
        with col_show:
            show_score_to_student = st.checkbox("Hiển thị điểm cho học sinh ngay sau khi nộp bài", value=True)
            
        up_file = st.file_uploader("Chọn file Word đề thi (.docx):", type=["docx"])
        
        if up_file and st.button("Quét và Lưu Đề Thi"):
            try:
                parsed_q = parse_docx_to_questions(up_file)
                if not parsed_q:
                    st.error("⚠️ Không đọc được câu hỏi từ file Word.")
                else:
                    mcq_count = sum(1 for q in parsed_q if q.get("type") == "mcq")
                    tf_count = sum(1 for q in parsed_q if q.get("type") == "true_false")
                    
                    st.success(f"🔍 Hệ thống đã quét thành công: **{len(parsed_q)} câu** (Trắc nghiệm: {mcq_count}, Đúng/Sai: {tf_count}).")
                    
                    for q in parsed_q:
                        q["points"] = default_q_pts
                        
                    cursor.execute("""
                        INSERT OR REPLACE INTO exams (exam_code, title, class_target, duration, max_attempts, show_score, questions, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        up_code, up_title, up_class, up_duration, up_max_attempts, 
                        1 if show_score_to_student else 0,
                        json.dumps(parsed_q, ensure_ascii=False),
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ))
                    conn.commit()
                    st.info(f"✅ Đã lưu đề thi `{up_code}` vào kho đề!")
            except Exception as e:
                st.error(f"Lỗi xử lý file: {e}")
                
        st.markdown("---")
        st.subheader("📚 Kho Đề Thi Hiện Có")
        cursor.execute("SELECT exam_code, title, class_target, duration, max_attempts, show_score, created_at FROM exams")
        exams_db = cursor.fetchall()
        if exams_db:
            df_exams = pd.DataFrame(exams_db, columns=["Mã đề", "Tiêu đề", "Lớp áp dụng", "Thời gian (phút)", "Số lần thi", "Hiện điểm", "Ngày tạo"])
            st.dataframe(df_exams, use_container_width=True)
        else:
            st.info("Kho đề thi trống.")

    with t_tab2:
        st.subheader("1. Thêm Lớp Học Mới")
        new_class_name = st.text_input("Tên lớp mới (VD: 11A1, 11A2):")
        if st.button("Thêm Lớp"):
            if new_class_name:
                try:
                    cursor.execute("INSERT INTO classes (class_name, created_at) VALUES (?, ?)", (new_class_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    conn.commit()
                    st.success(f"✅ Đã thêm lớp `{new_class_name}` thành công!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("❌ Lớp học này đã tồn tại.")
                    
        st.markdown("---")
        st.subheader("2. Quản Lý Danh Sách Học Sinh Theo Lớp")
        cursor.execute("SELECT class_name FROM classes")
        cls_db = [c[0] for c in cursor.fetchall()]
        
        if cls_db:
            selected_manage_class = st.selectbox("Chọn lớp cần xem danh sách học sinh:", cls_db, key="manage_cls")
            cursor.execute("SELECT stt, name, student_id, email FROM users WHERE class_name = ? AND role = 'Học sinh' ORDER BY stt", (selected_manage_class,))
            students_in_class = cursor.fetchall()
            
            if students_in_class:
                df_students = pd.DataFrame(students_in_class, columns=["STT", "Họ tên", "MSSV", "Email Google"])
                st.dataframe(df_students, use_container_width=True)
            else:
                st.info(f"Chưa có học sinh nào đăng ký vào lớp {selected_manage_class}.")
        else:
            st.info("Chưa có lớp học nào được tạo.")

    with t_tab3:
        st.subheader("📊 Quản Lý Bảng Điểm & Xuất File Excel Theo Lớp")
        cursor.execute("SELECT class_name FROM classes")
        cls_export_list = [c[0] for c in cursor.fetchall()]
        
        if cls_export_list:
            selected_export_class = st.selectbox("Chọn lớp cần xem và xuất điểm:", cls_export_list, key="export_cls")
            
            cursor.execute("""
                SELECT id, exam_code, student_name, student_id, stt, score, violations, status, submitted_at 
                FROM submissions WHERE class_name = ?
            """, (selected_export_class,))
            class_subs = cursor.fetchall()
            
            if class_subs:
                df_class_subs = pd.DataFrame(class_subs, columns=["ID", "Mã đề", "Họ tên", "MSSV", "STT", "Điểm số", "Vi phạm", "Trạng thái", "Thời gian nộp"])
                st.dataframe(df_class_subs, use_container_width=True)
                
                csv_bytes = df_class_subs.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Tải Xuống Bảng Điểm Lớp {selected_export_class} (CSV/Excel)",
                    data=csv_bytes,
                    file_name=f"bang_diem_lop_{selected_export_class}.csv",
                    mime="text/csv"
                )
            else:
                st.info(f"Chưa có bài nộp nào từ học sinh lớp {selected_export_class}.")
        else:
            st.info("Chưa có lớp học nào.")

    with t_tab4:
        st.subheader("🚨 Thống Kê Học Sinh Vi Phạm Quy Chế Thi")
        cursor.execute("""
            SELECT id, exam_code, student_name, student_id, stt, class_name, status, submitted_at 
            FROM submissions WHERE violations > 0 OR status LIKE '%vi phạm%' OR status LIKE '%khóa%'
        """)
        violations_db = cursor.fetchall()
        
        if violations_db:
            df_violations = pd.DataFrame(violations_db, columns=["ID", "Mã đề", "Họ tên", "MSSV", "STT", "Lớp", "Trạng thái vi phạm", "Thời gian ghi nhận"])
            st.dataframe(df_violations, use_container_width=True)
            st.warning("⚠️ Những học sinh có tên trong danh sách trên đã thực hiện hành vi rời khỏi tab, chia đôi màn hình hoặc mở trợ lý AI trong lúc làm bài!")
        else:
            st.success("🎉 Tuyệt vời! Chưa có học sinh nào vi phạm quy chế trong đợt kiểm tra này.")

    with t_tab5:
        st.subheader("🔐 Đổi Mật Khẩu Quản Trị (Admin)")
        new_pwd = st.text_input("Mật khẩu admin mới:", type="password")
        if st.button("Cập Nhật Mật Khẩu"):
            if new_pwd:
                cursor.execute("UPDATE settings SET value = ? WHERE key = 'admin_password'", (new_pwd,))
                conn.commit()
                st.success("✅ Đã cập nhật mật khẩu admin thành công!")
