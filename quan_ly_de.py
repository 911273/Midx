import os
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime
import io
import qrcode
from PIL import Image, ImageTk
import utils
import db

# Import từ tron_de để parse file Word
try:
    from tron_de import split_questions_from_docx
except ImportError:
    split_questions_from_docx = None

OUTPUT_FOLDER = "DE_TRON"

class ExamManagerTab:
    def __init__(self, parent):
        self.parent = parent
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        
        self.parent.rowconfigure(1, weight=1)
        self.parent.columnconfigure(0, weight=1)
        
        # --- Toolbar ---
        frm_toolbar = ttk.Frame(self.parent)
        frm_toolbar.grid(row=0, column=0, sticky="ew", padx=utils.PAD_M, pady=utils.PAD_S)
        
        ttk.Button(frm_toolbar, text="🔄 Làm mới", command=self.load_list).pack(side=tk.LEFT, padx=utils.PAD_XS)
        ttk.Button(frm_toolbar, text="📂 Xem thư mục", command=self.open_selected_folder).pack(side=tk.LEFT, padx=utils.PAD_XS)
        ttk.Button(frm_toolbar, text="🔍 Xem trước đề", command=self.preview_exam).pack(side=tk.LEFT, padx=utils.PAD_XS)
        ttk.Button(frm_toolbar, text="📝 Tổng hợp", command=self.open_tonghop_file).pack(side=tk.LEFT, padx=utils.PAD_XS)
        ttk.Button(frm_toolbar, text="📊 Đáp án", command=self.open_dapan_file).pack(side=tk.LEFT, padx=utils.PAD_XS)
        ttk.Button(frm_toolbar, text="📋 Chi tiết", command=self.view_db_detail).pack(side=tk.LEFT, padx=utils.PAD_XS)
        ttk.Button(frm_toolbar, text="🔗 QR Code", command=self.manage_qr_code).pack(side=tk.LEFT, padx=utils.PAD_XS)
        ttk.Button(frm_toolbar, text="📝 Google Script", command=self.generate_form_script, bootstyle="info").pack(side=tk.LEFT, padx=utils.PAD_XS)
        ttk.Button(frm_toolbar, text="📊 Xem điểm", command=self.view_scores, bootstyle="success").pack(side=tk.LEFT, padx=utils.PAD_XS)
        
        ttk.Button(frm_toolbar, text="❌ Xóa đợt này", command=self.delete_run, bootstyle="danger").pack(side=tk.RIGHT, padx=utils.PAD_XS)
        
        # --- Treeview ---
        frm_list = ttk.Frame(self.parent)
        frm_list.grid(row=1, column=0, sticky="nsew", padx=utils.PAD_M, pady=utils.PAD_XS)
        frm_list.rowconfigure(0, weight=1)
        frm_list.columnconfigure(0, weight=1)
        
        cols = ("foldername", "created", "files_count", "size", "has_aggregate", "db_status")
        self.tree = ttk.Treeview(frm_list, columns=cols, show="headings", bootstyle="info", selectmode="extended")
        self.tree.heading("foldername", text="Tên đợt trộn đề")
        self.tree.heading("created", text="Thời gian tạo")
        self.tree.heading("files_count", text="Số lượng file")
        self.tree.heading("size", text="Kích thước (MB)")
        self.tree.heading("has_aggregate", text="Tổng hợp & Đáp án")
        self.tree.heading("db_status", text="CSDL")
        
        self.tree.column("foldername", width=200, anchor=tk.W)
        self.tree.column("created", width=130, anchor=tk.CENTER)
        self.tree.column("files_count", width=80, anchor=tk.CENTER)
        self.tree.column("size", width=80, anchor=tk.E)
        self.tree.column("has_aggregate", width=130, anchor=tk.CENTER)
        self.tree.column("db_status", width=80, anchor=tk.CENTER)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        vsb = ttk.Scrollbar(frm_list, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.bind("<Double-1>", lambda e: self.open_selected_folder())
        
        self.load_list()

    def get_selected_path(self):
        """Trả về đường dẫn của mục được chọn đầu tiên."""
        selected = self.tree.selection()
        if not selected:
            return None
        item = self.tree.item(selected[0])
        fname = item['values'][0]
        return os.path.join(OUTPUT_FOLDER, fname)

    def get_selected_paths(self):
        """Trả về danh sách đường dẫn của tất cả các mục được chọn."""
        selected = self.tree.selection()
        if not selected:
            return []
        paths = []
        for sel in selected:
            item = self.tree.item(sel)
            fname = item['values'][0]
            paths.append(os.path.join(OUTPUT_FOLDER, fname))
        return paths

    def load_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if not os.path.exists(OUTPUT_FOLDER):
            return
        
        runs = [f for f in os.listdir(OUTPUT_FOLDER) if os.path.isdir(os.path.join(OUTPUT_FOLDER, f))]
        runs.sort(key=lambda x: os.path.getctime(os.path.join(OUTPUT_FOLDER, x)), reverse=True)
        
        for r in runs:
            try:
                path = os.path.join(OUTPUT_FOLDER, r)
                files = os.listdir(path)
                ctime = os.path.getctime(path)
                dt_str = datetime.fromtimestamp(ctime).strftime("%d/%m/%Y %H:%M")
                
                # Tính size thư mục
                total_size_bytes = sum(os.path.getsize(os.path.join(path, f)) for f in files if os.path.isfile(os.path.join(path, f)))
                size_mb = total_size_bytes / (1024 * 1024)
                
                # Check has dap_an and tong_hop
                has_dapan = "dap_an.xlsx" in files
                has_tonghop = "Tong_hop_de.docx" in files
                
                if has_dapan and has_tonghop:
                    status = "✅ Đầy đủ"
                elif has_dapan:
                    status = "⚠️ Thiếu tổng hợp"
                elif has_tonghop:
                    status = "⚠️ Thiếu đáp án"
                else:
                    status = "❌ Không có"
                
                # Kiểm tra DB
                db_stat = "⬜"
                try:
                    abs_path = os.path.abspath(path)
                    session = db.find_session_by_folder(abs_path)
                    if not session:
                        session = db.find_session_by_folder(path)
                    db_stat = "✅" if session else "⬜"
                except Exception:
                    pass
                    
                self.tree.insert("", "end", values=(r, dt_str, f"{len(files)} file", f"{size_mb:.2f}", status, db_stat))
            except (PermissionError, OSError):
                # Bỏ qua các thư mục đang bị khóa hoặc lỗi truy cập hệ thống
                continue

    def open_selected_folder(self):
        path = self.get_selected_path()
        if not path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 đợt trộn đề!")
            return
        try:
            os.startfile(os.path.abspath(path))
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def open_tonghop_file(self):
        path = self.get_selected_path()
        if not path: return
        file_path = os.path.join(path, "Tong_hop_de.docx")
        if not os.path.exists(file_path):
            messagebox.showwarning("Không tồn tại", "Đợt trộn này không có file Tổng hợp đề.")
            return
        try:
            os.startfile(os.path.abspath(file_path))
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def open_dapan_file(self):
        path = self.get_selected_path()
        if not path: return
        file_path = os.path.join(path, "dap_an.xlsx")
        if not os.path.exists(file_path):
            messagebox.showwarning("Không tồn tại", "Đợt trộn này không có file Đáp án.")
            return
        try:
            os.startfile(os.path.abspath(file_path))
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def delete_run(self):
        paths = self.get_selected_paths()
        if not paths:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất 1 đợt trộn đề để xóa!")
            return
            
        n = len(paths)
        if n == 1:
            run_name = os.path.basename(paths[0])
            confirm_msg = f"Bạn có chắc chắn muốn xóa TOÀN BỘ thư mục trộn đề:\n{run_name} không?"
        else:
            confirm_msg = f"Bạn có chắc chắn muốn xóa {n} đợt trộn đề đã chọn không?"
            
        if messagebox.askyesno("Xác nhận", confirm_msg):
            success_count = 0
            for path in paths:
                try:
                    # Xóa bản ghi DB tương ứng
                    try:
                        abs_path = os.path.abspath(path)
                        db.delete_session_by_folder(abs_path)
                        db.delete_session_by_folder(path)
                    except Exception:
                        pass
                    
                    if os.path.exists(path):
                        shutil.rmtree(path)
                    success_count += 1
                except Exception as e:
                    print(f"Lỗi khi xóa {path}: {e}")
            
            self.load_list()
            if n == 1:
                messagebox.showinfo("Thành công", "Đã xóa đợt trộn đề thành công.")
            else:
                messagebox.showinfo("Thành công", f"Đã xóa {success_count}/{n} đợt trộn đề thành công.")

    def preview_exam(self):
        """Xem trước nội dung một mã đề trong đợt trộn."""
        path = self.get_selected_path()
        if not path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 đợt trộn đề!")
            return
        
        if not split_questions_from_docx:
            messagebox.showerror("Lỗi", "Không tìm thấy hàm phân tích đề thi (split_questions_from_docx).")
            return

        # Liệt kê các file .docx trong thư mục (trừ Tong_hop_de.docx)
        files = [f for f in os.listdir(path) if f.lower().endswith('.docx') 
                 and f != "Tong_hop_de.docx" and not f.startswith('~')]
        
        if not files:
            messagebox.showwarning("Thông báo", "Không tìm thấy file mã đề (.docx) trong thư mục này.")
            return

        selected_file = None
        if len(files) == 1:
            selected_file = files[0]
        else:
            # Hiện cửa sổ chọn mã đề
            dlg_pick = tk.Toplevel(self.parent)
            utils.setup_dialog(dlg_pick, width_pct=0.3, height_pct=0.5, title="Chọn mã đề", parent=self.parent)
            
            lbl = ttk.Label(dlg_pick, text="Chọn một mã đề cụ thể:", padding=10)
            lbl.pack()
            
            lb = tk.Listbox(dlg_pick)
            lb.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
            for f in files: lb.insert(tk.END, f)
            
            def on_pick():
                nonlocal selected_file
                sel = lb.curselection()
                if sel:
                    selected_file = lb.get(sel[0])
                    dlg_pick.destroy()
            
            ttk.Button(dlg_pick, text="Xem", command=on_pick).pack(pady=10)
            self.parent.wait_window(dlg_pick)

        if not selected_file: return
        
        full_path = os.path.join(path, selected_file)
        self._show_preview_window(full_path)

    def _show_preview_window(self, file_path):
        """Hiển thị cửa sổ xem nội dung file đề bằng QuestionPreviewDialog chung."""
        try:
            qs, _, _ = split_questions_from_docx(file_path)
            # Sử dụng Dialog chuẩn từ utils
            utils.QuestionPreviewDialog(self.parent, qs, title=f"Xem trước đề: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xem trước đề thi: {e}")

    def view_db_detail(self):
        """Xem chi tiết đợt trộn đề từ cơ sở dữ liệu."""
        path = self.get_selected_path()
        if not path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 đợt trộn đề!")
            return
        
        abs_path = os.path.abspath(path)
        session = db.find_session_by_folder(abs_path)
        if not session:
            session = db.find_session_by_folder(path)
        
        if not session:
            messagebox.showinfo("Thông báo", 
                "Đợt trộn đề này chưa được lưu trong CSDL.\n"
                "Chỉ các đợt trộn mới (sau khi bổ sung tính năng DB) mới có dữ liệu.")
            return
        
        dlg = tk.Toplevel(self.parent)
        utils.setup_dialog(dlg, width_pct=0.5, height_pct=0.6, title="Chi tiết đợt trộn đề (DB)", parent=self.parent)
        
        txt = tk.Text(dlg, wrap="word", font=("Consolas", 10))
        txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        msg = "📋 THÔNG TIN ĐỢT TRỘN ĐỀ (TỪ CSDL)\n"
        msg += "=" * 45 + "\n\n"
        msg += f"🏫 Trường:        {session.get('school', '')}\n"
        msg += f"🏛️ Khoa:          {session.get('faculty', '')}\n"
        msg += f"📚 Môn học:       {session.get('subject_name', '')}\n"
        msg += f"📅 Năm học:       {session.get('school_year', '')}\n"
        msg += f"🕒 Học kỳ:        {session.get('semester', '')}\n"
        msg += f"👥 Lớp:           {session.get('class_name', '')}\n"
        msg += f"📝 Bài kiểm tra:  {session.get('exam_title', '')}\n"
        msg += f"⏱️ Thời gian:     {session.get('duration', '')}\n\n"
        msg += f"📊 Số mã đề:      {session.get('num_variants', 0)}\n"
        msg += f"❓ Số câu/đề:     {session.get('num_questions', 0)}\n"
        msg += f"🔀 Chiến lược:    {session.get('strategy', '')}\n"
        msg += f"🔄 Đảo đáp án:    {'Có' if session.get('shuffle_answers') else 'Không'}\n\n"
        msg += f"📁 Thư mục:       {session.get('folder_path', '')}\n"
        msg += f"🕐 Ngày tạo:      {session.get('created_at', '')}\n"
        msg += f"🆔 Session ID:    {session.get('id', '')}\n"
        
        txt.insert(tk.END, msg)
        txt.config(state="disabled")

    def generate_form_script(self):
        """Tạo mã kịch bản Google Apps Script từ đợt thi đang chọn."""
        path = self.get_selected_path()
        if not path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 đợt trộn đề!")
            return
        
        abs_path = os.path.abspath(path)
        session = db.find_session_by_folder(abs_path)
        if not session:
            session = db.find_session_by_folder(path)
            
        # --- Dự phòng: Đọc từ metadata.json nếu không có trong DB ---
        if not session:
            meta_path = os.path.join(path, "metadata.json")
            if os.path.exists(meta_path):
                try:
                    import json
                    with open(meta_path, "r", encoding="utf-8") as f:
                        session = json.load(f)
                        # Map keys if needed (metadata.json might have slightly different names)
                        if "subject" in session and "subject_name" not in session:
                            session["subject_name"] = session["subject"]
                except Exception as e:
                    print(f"Lỗi đọc metadata.json: {e}")

        if not session:
            messagebox.showinfo("Thông báo", 
                "Đợt thi này chưa có trong CSDL và không tìm thấy file metadata.json dự phòng.\n"
                "Không thể lấy thông tin để tạo kịch bản.")
            return

        import re
        from datetime import datetime

        # Lấy dữ liệu
        school   = session.get("school", "ĐẠI HỌC ĐIỆN LỰC")
        faculty  = session.get("faculty", "KHOA NĂNG LƯỢNG MỚI")
        ex_title = session.get("exam_title", "")
        year_txt = session.get("school_year", "")
        semester = session.get("semester", "")
        subject  = session.get("subject_name", "")
        class_n  = session.get("class_name", "")
        n_qs     = session.get("num_questions", 50)
        
        # Dialog confirm
        dlg = tk.Toplevel(self.parent)
        utils.setup_dialog(dlg, width_pct=0.4, height_pct=0.6, title="Tạo Google Form Script", parent=self.parent)
        
        frame = ttk.Frame(dlg, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        r = 0
        ttk.Label(frame, text="Nhập các thông tin bổ sung cho Form:").grid(row=r, column=0, columnspan=2, pady=(0, 10))
        r += 1

        ttk.Label(frame, text="Tên Form:").grid(row=r, column=0, sticky="w", pady=5)
        ent_title = ttk.Entry(frame, width=40)
        ent_title.insert(0, f"BÀI THI: {ex_title or subject}")
        ent_title.grid(row=r, column=1, sticky="ew", pady=5)
        r += 1

        ttk.Label(frame, text="Số câu hỏi:").grid(row=r, column=0, sticky="w", pady=5)
        ent_n = ttk.Entry(frame, width=15)
        ent_n.insert(0, str(n_qs))
        ent_n.grid(row=r, column=1, sticky="w", pady=5)
        r += 1
        
        def js_esc(s):
            return str(s).replace("'", "\\'").replace("\n", " ")

        def on_confirm():
            n = ent_n.get()
            esc_form_title = js_esc(ent_title.get())
            esc_school = js_esc(school)
            esc_faculty = js_esc(faculty)
            esc_subject = js_esc(subject)
            esc_class = js_esc(class_n)
            esc_title = js_esc(ex_title)
            esc_year = js_esc(year_txt)
            esc_semester = js_esc(semester)

            script_content = f"""function taoFormThi() {{
  var form = FormApp.create('{esc_form_title}');
  form.setDescription('Thông tin bài thi:\\n- Trường: {esc_school}\\n- Khoa: {esc_faculty}\\n- Môn học: {esc_subject}\\n- Lớp: {esc_class}\\n- Kỳ thi: {esc_title}\\n- Bài số: {semester}\\n- Năm học: {esc_year} (Học kỳ: {esc_semester})\\n\\nLưu ý: Sinh viên điền đúng thông tin và chỉ chọn 1 đáp án cho mỗi câu.');
  
  form.addTextItem().setTitle('Họ và tên').setRequired(true);
  form.addTextItem().setTitle('Mã sinh viên').setRequired(true);
  form.addTextItem().setTitle('Lớp').setRequired(true);
  form.addTextItem().setTitle('Mã đề').setRequired(true);
  
  var numQuestions = {n};
  var choices = ['A', 'B', 'C', 'D'];
  var gridItem = form.addGridItem();
  gridItem.setTitle('Phần trả lời (Chọn đáp án chính xác)');
  
  var rows = [];
  for (var i = 1; i <= numQuestions; i++) {{
    rows.push('Câu ' + i);
  }}
  gridItem.setRows(rows);
  gridItem.setColumns(choices);
  gridItem.setRequired(true);
  
  Logger.log('Tạo form thành công!');
  Logger.log('URL Form: ' + form.getEditUrl());
}}"""
            try:
                out_path = os.path.join(path, "Tao_Google_Form_Script.txt")
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(script_content)
                
                dlg.destroy()
                msg = (f"Đã tạo mã kịch bản thành công tại:\n{out_path}\n\n"
                       "Hướng dẫn:\n"
                       "1. Truy cập: https://script.google.com và tạo Dự án mới.\n"
                       "2. Dán mã này vào và bấm 'Chạy'.\n")
                messagebox.showinfo("Thành công", msg)
                os.startfile(out_path)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể tạo file: {e}")

        ttk.Button(frame, text="Tạo mã script", command=on_confirm, bootstyle="primary").grid(row=r, column=0, columnspan=2, pady=20)

    def view_scores(self):
        """Xem kết quả chấm điểm của đợt thi đang chọn."""
        path = self.get_selected_path()
        if not path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 đợt trộn đề!")
            return
            
        abs_path = os.path.abspath(path)
        session = db.find_session_by_folder(abs_path)
        if not session:
            session = db.find_session_by_folder(path)
        
        if not session:
            messagebox.showinfo("Thông báo", "Đợt thi này chưa được lưu trong CSDL.")
            return

        # Gọi hàm hiển thị lịch sử từ cham_diem nhưng filter theo session_id
        try:
            from cham_diem import show_grading_history_for_session
            show_grading_history_for_session(self.parent, session)
        except ImportError:
            # Fallback nếu chưa kịp sửa cham_diem
            messagebox.showinfo("Thông báo", "Tính năng này đang được cập nhật.")

    def manage_qr_code(self):
        """Quản lý link Google Form và tạo mã QR."""
        path = self.get_selected_path()
        if not path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 đợt trộn đề!")
            return
        
        abs_path = os.path.abspath(path)
        session = db.find_session_by_folder(abs_path)
        if not session:
            session = db.find_session_by_folder(path)
            
        if not session:
            messagebox.showinfo("Thông báo", 
                "Đợt trộn đề này chưa được lưu trong CSDL nên không thể gán link Google Form.")
            return
            
        QRManagerDialog(self.parent, session, path)

class QRManagerDialog:
    def __init__(self, parent, session, folder_path):
        self.parent = parent
        self.session = session
        self.folder_path = folder_path
        self.session_id = session['id']
        
        self.win = tk.Toplevel(parent)
        utils.setup_dialog(self.win, width_pct=0.4, height_pct=0.6, title="Quản lý QR Code nộp bài", parent=parent)
        
        # UI Elements
        frm_input = ttk.Frame(self.win, padding=10)
        frm_input.pack(fill=tk.X)
        
        ttk.Label(frm_input, text="Link Google Form nộp bài:").pack(anchor=tk.W)
        self.ent_link = ttk.Entry(frm_input)
        self.ent_link.pack(fill=tk.X, pady=5)
        
        # Load existing link
        current_link = session.get('google_form_link', '')
        self.ent_link.insert(0, current_link)
        
        frm_buttons = ttk.Frame(frm_input)
        frm_buttons.pack(fill=tk.X, pady=5)
        
        self.btn_save = ttk.Button(frm_buttons, text="💾 Lưu & Tạo QR", command=self.generate_and_save, bootstyle="success")
        self.btn_save.pack(side=tk.LEFT, padx=5)
        
        self.btn_export = ttk.Button(frm_buttons, text="🖼️ Xuất ảnh QR", command=self.export_qr_image)
        self.btn_export.pack(side=tk.LEFT, padx=5)
        
        # QR Preview Area
        self.lbl_qr = ttk.Label(self.win, text="[Mã QR sẽ hiển thị ở đây]")
        self.lbl_qr.pack(expand=True, pady=10)
        
        self.qr_img = None
        if current_link:
            self.show_qr(current_link)
            
    def show_qr(self, data):
        if not data: return
        try:
            qr = qrcode.QRCode(version=1, box_size=10, border=5)
            qr.add_data(data)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            
            # Convert to PhotoImage for TK
            img_thumb = img.resize((250, 250), Image.LANCZOS)
            self.tk_qr = ImageTk.PhotoImage(img_thumb)
            self.lbl_qr.config(image=self.tk_qr, text="")
            self.qr_pil_image = img # Keep original for export
        except Exception as e:
            messagebox.showerror("Lỗi QR", str(e))

    def generate_and_save(self):
        link = self.ent_link.get().strip()
        if not link:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập link!")
            return
        
        # Update DB
        try:
            db.update_exam_session_link(self.session_id, link)
            self.show_qr(link)
            messagebox.showinfo("Thành công", "Đã lưu link vào CSDL.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu vào CSDL: {e}")

    def export_qr_image(self):
        link = self.ent_link.get().strip()
        if not link:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập link trước khi xuất ảnh!")
            return
            
        try:
            qr = qrcode.QRCode(version=1, box_size=10, border=5)
            qr.add_data(link)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            
            export_path = os.path.join(self.folder_path, "QR_Nop_Bai.png")
            img.save(export_path)
            
            messagebox.showinfo("Thành công", f"Đã lưu ảnh QR tại:\n{export_path}")
            os.startfile(os.path.abspath(self.folder_path))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất ảnh: {e}")
