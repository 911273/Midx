# -*- coding: utf-8 -*-
"""
Tab Chấm điểm – có điểm mỗi câu đúng/sai
- Nhập: dap_an.xlsx (cột A=Mã đề, các cột sau: Câu 1..N)
- Nhập: file trả lời (xlsx/csv) – Google Forms/Excel tự tạo
- Nhận diện mạnh:
  • Mã đề: ưu tiên "Mã đề", sau đó các biến thể "ma de", "ma de thi", "exam/code"
  • MSSV: "MSSV", "Mã sinh viên", "Mã số sinh viên", "MSV", "Student ID", "ID", …
  • Họ & Tên: "Họ và tên", "Họ tên", "Họ & tên", "Tên sinh viên", "Name", "Full name", …
  • Câu trả lời: "Đáp án [i]" / "Đáp án i" / "Đáp án câu i" / "Answer [i]" / "Ans i" / "Câu i" / "Q i" / "Question i"
    - Bỏ trùng: giữ cột đầu tiên bắt được cho cùng số câu (tránh “…].1”)
- Chuẩn hóa đáp án: lấy ký tự A–D đầu tiên; nếu không thấy → để trống
- Tùy chọn chấm điểm:
    * Điểm mỗi câu đúng (mặc định = 10 / số câu)
    * Điểm mỗi câu sai (mặc định = 0)
    * Tùy chọn "coi bỏ trống là sai"
- Xuất: Ket_qua_cham_diem_YYYYmmdd_HHMMSS.xlsx (ChiTiet / ThongKe / CauHinh)
  * Bổ sung các cột: "Mã sinh viên", "Số lượng câu đúng", "Số lượng câu sai", "Điểm"
"""

import os
import re
import math
import unicodedata
from datetime import datetime
from typing import List, Dict, Any, Optional
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Text, END
import utils
import db

try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False


# ===================== CHUẨN HÓA CHUỖI =====================
def _strip_accents(s: str) -> str:
    try:
        return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    except Exception:
        return s

def _norm_header(s: str) -> str:
    """Chuẩn hóa tiêu đề: bỏ ký tự ẩn, &→'va', bỏ dấu, lower, gom space."""
    s = str(s or "")
    s = s.replace('\u00A0', ' ')  # nbsp
    s = s.replace('\u200b', '')   # zero-width space
    s = s.replace('\u200c', '')
    s = s.replace('\u200d', '')
    s = s.replace('\ufeff', '')   # BOM
    s = s.replace('&', ' va ')
    s = s.strip()
    s = _strip_accents(s).lower()
    s = re.sub(r'\s+', ' ', s)
    return s

def _answer_letter(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ''
    text = str(s).strip()
    if not text:
        return ''
    m = re.search(r'([A-Da-d])', text)
    return m.group(1).upper() if m else ''

def _coerce_exam_code(x):
    if pd.isna(x):
        return None
    xs = str(x).strip()
    m = re.search(r'(\d+)', xs)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return xs
    try:
        f = float(xs)
        i = int(f)
        if abs(f - i) < 1e-9:
            return i
        return xs
    except Exception:
        return xs or None

# ===================== TÌM CỘT THEO BỘ TỪ KHÓA =====================
def _find_ma_de_column(cols):
    for i, c in enumerate(cols):
        if str(c).strip() == "Mã đề":
            return i
    for i, c in enumerate(cols):
        if _norm_header(c) == "ma de":
            return i
    return None

def _find_by_alias_exact(cols, alias_norm_list):
    ncols = [_norm_header(c) for c in cols]
    for i, n in enumerate(ncols):
        if n in alias_norm_list:
            return i
    return None

def _find_by_regex_score(cols, patterns_with_weight):
    best_i, best_score = None, -1
    for i, c in enumerate(cols):
        n = _norm_header(c)
        score = 0
        for pat, w in patterns_with_weight:
            if re.search(pat, n):
                score += w
        if score > best_score:
            best_score, best_i = score, i
    return best_i

# ===================== NHẬN DIỆN CỘT CÂU HỎI =====================
def _is_question_col(name: str) -> bool:
    n = _norm_header(name)
    if re.search(r'\bcau\s*\d+\b', n):                # "Câu 12"
        return True
    if re.search(r'\bq(uestion)?\s*\d+\b', n):        # "Q12", "Question 12"
        return True
    if re.search(r'\bdap\s*an\b.*?\[\s*\d+\s*\]', n): # "Đáp án [12]"
        return True
    if re.search(r'\bdap\s*an\s*(?:cau\s*)?\d+\b', n):# "Đáp án 12"/"Đáp án câu 12"
        return True
    if re.search(r'\bans(wer)?\b.*?\[\s*\d+\s*\]', n):# "Answer [12]"
        return True
    if re.search(r'\bans(wer)?\s*\d+\b', n):          # "Ans 12"
        return True
    return False

def _extract_question_number(name: str) -> Optional[int]:
    n = _norm_header(name)
    m = re.search(r'(\d+)', n)
    if m:
        return int(m.group(1))
    return None


# ===================== ĐỌC ĐÁP ÁN CHUẨN =====================
def load_answer_key(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError("Không tìm thấy file đáp án: " + path)
    df = pd.read_csv(path, dtype=object) if path.lower().endswith(".csv") else pd.read_excel(path, dtype=object)
    if df.empty:
        raise ValueError("File đáp án rỗng.")

    cols = list(df.columns)
    md_idx = _find_ma_de_column(cols)
    if md_idx is None:
        md_idx = _find_by_regex_score(cols, [
            (r'\bma\s*de\b', 5),
            (r'\bma\b.*\bde\b',4),
            (r'\bexam\b',3),
            (r'\bcode\b',2),
        ])
    if md_idx is None:
        md_idx = 0

    q_cols = []
    for j, c in enumerate(cols):
        if j == md_idx: continue
        m = re.search(r'\bcau\s*(\d+)\b', _norm_header(c))
        if m:
            q_cols.append((int(m.group(1)), j))
    if not q_cols:
        for j, c in enumerate(cols):
            if j == md_idx: continue
            m = re.search(r'(\d+)', _norm_header(c))
            if m:
                q_cols.append((int(m.group(1)), j))
    if not q_cols:
        raise ValueError("Không nhận diện được các cột câu hỏi trong đáp án.")

    q_cols.sort(key=lambda x: x[0])
    num_qs = len(q_cols)

    key = {}
    for _, row in df.iterrows():
        code = _coerce_exam_code(row.iloc[md_idx])
        if code is None: continue
        arr = []
        for _, j in q_cols:
            arr.append(_answer_letter(row.iloc[j]))
        key[code] = arr

    if not key:
        raise ValueError("Không đọc được dòng đáp án nào hợp lệ.")
    return key, num_qs

# ===================== ĐỌC FILE TRẢ LỜI =====================
def load_responses(path: str):
    """
    Trả về:
      df: DataFrame gốc
      mapping: {
        'exam_code_col': name,
        'student_id_col': name or None,
        'student_name_col': name or None,
        'question_cols': { qnum:int -> colname:str }  # đã bỏ trùng
      }
    """
    if not os.path.exists(path):
        raise FileNotFoundError("Không tìm thấy file trả lời: " + path)
    df = pd.read_csv(path, dtype=object) if path.lower().endswith(".csv") else pd.read_excel(path, dtype=object)
    if df.empty:
        raise ValueError("File trả lời rỗng.")
    cols = list(df.columns)

    # ---- auto detect trước ----
    try:
        md_idx = _find_ma_de_column(cols)
        if md_idx is None:
            md_idx = _find_by_alias_exact(cols, ["ma de", "ma de thi"])
        if md_idx is None:
            md_idx = _find_by_regex_score(cols, [
                (r'\bma\s*de\b', 6),(r'\bma\s*de\s*thi\b', 6),(r'\bexam\b', 3),(r'\bcode\b', 2),
            ])
        exam_code_col = cols[md_idx] if md_idx is not None else None

        id_idx = _find_by_alias_exact(cols, [
            "mssv","ma sinh vien","ma so sinh vien","ma sv","ma so sv",
            "student id","studentid","id","student number","roll no","roll number"
        ])
        if id_idx is None:
            id_idx = _find_by_regex_score(cols, [
                (r'\bmssv\b', 7),(r'\bma\s*sinh\s*vien\b',7),(r'\bma\s*so\s*sinh\s*vien\b',7),
                (r'\bma\s*sv\b',6),(r'\bstudent\s*id\b',6),(r'\bid\b',3),(r'\broll\s*(no|number)\b',4),
            ])
        student_id_col = cols[id_idx] if id_idx is not None else None

        name_idx = _find_by_alias_exact(cols, [
            "ho va ten","ho ten","ten sinh vien","ten hoc vien","full name","name","student name"
        ])
        if name_idx is None:
            name_idx = _find_by_regex_score(cols, [
                (r'\bho\s*(va|&)?\s*ten\b',7),(r'\bten\s*sinh\s*vien\b',6),
                (r'\bfull\s*name\b',4),(r'\bstudent\s*name\b',4),(r'\bname\b',2),
            ])
        student_name_col = cols[name_idx] if name_idx is not None else None

        q_map = {}
        debug_hits = []
        for c in cols:
            if _is_question_col(str(c)):
                qn = _extract_question_number(str(c))
                if qn is not None and qn not in q_map:
                    q_map[qn] = c; debug_hits.append((qn, c))
        if not q_map:
            for c in cols:
                n = _norm_header(c)
                m = re.search(r'\bcau\s*(\d+)\b', n)
                if m:
                    qn = int(m.group(1))
                    if qn not in q_map:
                        q_map[qn] = c; debug_hits.append((qn, c))

        if exam_code_col and q_map:
            try:
                print("MAP CỘT CÂU HỎI (auto):")
                for qn, colname in sorted(debug_hits, key=lambda x: x[0]):
                    print(f"  Câu {qn}  <-  {colname}")
            except Exception:
                pass
            return df, {
                'exam_code_col': exam_code_col,
                'student_id_col': student_id_col,
                'student_name_col': student_name_col,
                'question_cols': q_map
            }
    except Exception:
        pass

    # ---- Fallback theo bố cục bạn mô tả ----
    # Cột 1: Timestamp, 2: Họ tên, 3: MSSV, 4: Lớp, 5: Mã đề, 6..: Câu 1..N
    if len(cols) < 6:
        raise ValueError("File trả lời không đủ cột theo bố cục (cần ≥ 6 cột).")
    student_name_col = cols[1]
    student_id_col   = cols[2]
    exam_code_col    = cols[4]
    q_map = {}
    for j in range(5, len(cols)):
        qnum = j - 4
        q_map[qnum] = cols[j]
    try:
        print("FALLBACK (cố định vị trí):")
        print(f"  Họ & Tên  <- {student_name_col}")
        print(f"  MSSV      <- {student_id_col}")
        print(f"  Mã đề     <- {exam_code_col}")
        print("  Cột câu hỏi:")
        for qn in sorted(q_map.keys()):
            print(f"    Câu {qn} <- {q_map[qn]}")
    except Exception:
        pass

    return df, {
        'exam_code_col': exam_code_col,
        'student_id_col': student_id_col,
        'student_name_col': student_name_col,
        'question_cols': q_map
    }

# ===================== CHẤM ĐIỂM =====================
def grade_responses(answer_key: dict, num_qs_key: int, resp_df: pd.DataFrame, mapping: dict,
                    treat_blank_as_wrong=True,
                    point_per_correct: float = None,
                    point_per_wrong: float = 0.0):
    """
    Trả về (df_detail, df_stats, N_thuc_te)
    df_detail: từng bài (có Điểm theo tùy chọn)
    df_stats: thống kê từng câu
    """
    q_cols = mapping['question_cols']
    max_q = max(q_cols.keys())
    N = max(num_qs_key, max_q)

    # Mặc định điểm mỗi câu đúng = 10 / N nếu không truyền
    if point_per_correct is None:
        point_per_correct = 10.0 / max(N, 1)

    exam_code_col = mapping['exam_code_col']
    id_col = mapping.get('student_id_col')
    name_col = mapping.get('student_name_col')

    rows = []
    for _, row in resp_df.iterrows():
        code = _coerce_exam_code(row.get(exam_code_col))
        if code is None or code not in answer_key:
            continue
        key_arr = answer_key[code]
        if len(key_arr) < N:
            key_arr = key_arr + [''] * (N - len(key_arr))

        mssv = str(row.get(id_col)).strip() if id_col else ''
        if mssv.lower() == 'nan':
            mssv = ''
        hoten = str(row.get(name_col)).strip() if name_col else ''
        if hoten.lower() == 'nan':
            hoten = ''

        stu_ans = []
        correct = 0
        blanks = 0
        wrong_nonblank = 0

        for q in range(1, N+1):
            col = q_cols.get(q)
            ans = _answer_letter(row.get(col)) if col is not None else ''
            key = key_arr[q-1] if q-1 < len(key_arr) else ''

            if not ans:
                blanks += 1
                ok = (key == '') and (not treat_blank_as_wrong)
            else:
                ok = (ans == key)
                if not ok:
                    wrong_nonblank += 1

            if ok:
                correct += 1

            stu_ans.append(ans)

        # Số câu sai theo chế độ “bỏ trống tính sai / không”
        if treat_blank_as_wrong:
            wrong = N - correct
        else:
            wrong = wrong_nonblank

        # Tính điểm
        score = correct * point_per_correct + wrong * point_per_wrong

        out = {
            'Mã đề': code,
            'Mã sinh viên': mssv,     # bổ sung cột theo yêu cầu
            'MSSV': mssv,             # giữ cột MSSV cũ cho tương thích
            'Họ và tên': hoten,
            'Số lượng câu đúng': correct,
            'Số lượng câu sai': wrong,
            'Tổng số câu': N,
            'Điểm': round(score, 3),      # điểm theo cấu hình
            'Điểm (%)': round((correct / N * 100.0) if N else 0.0, 2)  # giữ thêm % để tham khảo
        }
        for q in range(1, N+1):
            out[f'Câu {q}'] = stu_ans[q-1]
        rows.append(out)

    if not rows:
        raise ValueError("Không có bản ghi hợp lệ sau khi đối chiếu Mã đề với đáp án.")

    df_detail = pd.DataFrame(rows)

    # ----- Thống kê theo câu -----
    ans_rows = []
    for _, r in df_detail.iterrows():
        code = r['Mã đề']
        key = answer_key[code]
        if len(key) < N:
            key = key + [''] * (N - len(key))
        ans_rows.append(key[:N])
    df_ans_by_row = pd.DataFrame(ans_rows, columns=[f'Câu {i}' for i in range(1, N+1)])

    stats = []
    for i in range(1, N+1):
        col = f'Câu {i}'
        s_stu = df_detail[col].astype(str).fillna('').str.strip().str.upper()
        s_key = df_ans_by_row[col].astype(str).fillna('').str.strip().str.upper()
        total = len(s_stu)
        blank = int((s_stu == '').sum())
        A = int((s_stu == 'A').sum()); B = int((s_stu == 'B').sum())
        C = int((s_stu == 'C').sum()); D = int((s_stu == 'D').sum())
        correct = int((s_stu == s_key).sum())
        p = round((correct/total) if total else 0.0, 4)
        stats.append({
            'Câu': i, 'Tổng bài': total, 'Đúng': correct, 'Độ khó (p)': p,
            'Chọn A': A, 'Chọn B': B, 'Chọn C': C, 'Chọn D': D,
            'Bỏ trống': blank, 'Tỉ lệ bỏ trống': round(blank/total, 4) if total else 0.0
        })
    df_stats = pd.DataFrame(stats)

    # ----- Tính toán chỉ số phân hóa (D) -----
    # Nhóm Cao (Top 27%) và Nhóm Thấp (Bottom 27%)
    try:
        n_students = len(df_detail)
        if n_students >= 10: # Cần đủ mẫu để tính D
            n_group = max(1, int(round(n_students * 0.27)))
            df_sorted = df_detail.sort_values("Điểm", ascending=False)
            upper_group = df_sorted.head(n_group)
            lower_group = df_sorted.tail(n_group)
            
            d_indices = []
            for i in range(1, N+1):
                col = f'Câu {i}'
                code_col = 'Mã đề'
                
                def get_p_correct(group_df):
                    correct_count = 0
                    for _, r in group_df.iterrows():
                        key = answer_key[r[code_col]]
                        if i <= len(key) and r[col] == key[i-1]:
                            correct_count += 1
                    return correct_count / len(group_df)

                p_u = get_p_correct(upper_group)
                p_l = get_p_correct(lower_group)
                d_indices.append(round(p_u - p_l, 3))
            
            df_stats['Độ phân hóa (D)'] = d_indices
        else:
            df_stats['Độ phân hóa (D)'] = "Mẫu ít"
    except Exception:
        df_stats['Độ phân hóa (D)'] = "Lỗi tính"

    return df_detail, df_stats, N

# ===================== GUI TAB =====================
import json

class GradeTab:
    def __init__(self, parent):
        self.parent = parent
        
        # Sử dụng Notebook để chia sub-tabs: Chấm điểm và Kết quả
        self.nb = ttk.Notebook(parent)
        self.nb.pack(fill=tk.BOTH, expand=True)

        self.tab_work = ttk.Frame(self.nb)
        self.tab_results = ttk.Frame(self.nb)

        self.nb.add(self.tab_work, text=" 🎯 Chấm điểm ")
        self.nb.add(self.tab_results, text=" 📊 Kết quả tổng hợp ")

        # --- TAB 1: CHẤM ĐIỂM (Giao diện cũ) ---
        self.setup_grading_tab(self.tab_work)
        
        # --- TAB 2: KẾT QUẢ (Giao diện mới) ---
        self.setup_results_tab(self.tab_results)

        # Gắn sự kiện chuyển tab để tự động làm mới bộ lọc
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    def on_tab_changed(self, event):
        selected_tab = self.nb.index(self.nb.select())
        if selected_tab == 1: # Tab Kết quả tổng hợp
            self.refresh_filter_options()
            self.search_results()


    def setup_grading_tab(self, container):
        # Sử dụng PanedWindow để phân chia khu vực Quản lý và Chấm điểm
        self.paned = ttk.PanedWindow(container, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- KHUNG TRÁI: TreeView quản lý ---
        self.frm_tree = ttk.LabelFrame(self.paned, text="Quản lý đợt thi", padding=utils.PAD_S)
        self.paned.add(self.frm_tree, weight=1)
        
        self.tree = ttk.Treeview(self.frm_tree, show="tree", bootstyle="primary")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        tree_scroll = ttk.Scrollbar(self.frm_tree, orient="vertical", command=self.tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # --- KHUNG PHẢI: Giao diện chấm điểm hiện tại ---
        self.frm_main = ttk.Frame(self.paned)
        self.paned.add(self.frm_main, weight=3)
        
        # Cấu hình grid cho frm_main (thay thế cho parent trước đây)
        for c in range(2):
            self.frm_main.columnconfigure(c, weight=1, uniform="col")
        self.frm_main.rowconfigure(3, weight=1)

        # --- Nội dung phía trên: Nguồn và Tùy chọn ---
        # --- Trái: chọn file ---
        left = ttk.LabelFrame(self.frm_main, text="Nguồn dữ liệu", padding=utils.PAD_M)
        left.grid(row=0, column=0, sticky="nsew", padx=utils.PAD_S, pady=utils.PAD_S)
        left.columnconfigure(0, weight=1)

        r = 0
        ttk.Label(left, text="File đáp án:").pack(anchor="w", padx=utils.PAD_S)
        
        frm_ans_btn = ttk.Frame(left)
        frm_ans_btn.pack(fill=tk.X, padx=utils.PAD_S, pady=utils.PAD_XS)
        ttk.Button(frm_ans_btn, text="File ngoài", command=self.pick_ans, width=12).pack(side=tk.LEFT)
        ttk.Button(frm_ans_btn, text="Làm mới", command=self.load_metadata_tree, width=10).pack(side=tk.LEFT, padx=5)
        
        self.lbl_ans = ttk.Label(left, text="Chưa chọn", foreground="orange", font=utils.FONT_MAIN)
        self.lbl_ans.pack(anchor="w", padx=utils.PAD_S, pady=(0, utils.PAD_S))

        ttk.Separator(left, orient="horizontal").pack(fill=tk.X, pady=utils.PAD_S)

        ttk.Label(left, text="File trả lời:").pack(anchor="w", padx=utils.PAD_S)
        ttk.Button(left, text="Chọn file trả lời", command=self.pick_resp).pack(anchor="w", padx=utils.PAD_S, pady=utils.PAD_XS)
        self.lbl_resp = ttk.Label(left, text="Chưa chọn", foreground="orange", font=utils.FONT_MAIN)
        self.lbl_resp.pack(anchor="w", padx=utils.PAD_S)

        # --- Phải: tùy chọn ---
        self.var_grade_school = tk.StringVar(value="ĐẠI HỌC ĐIỆN LỰC")
        self.var_grade_faculty = tk.StringVar(value="KHOA NĂNG LƯỢNG MỚI")
        self.var_grade_examtitle = tk.StringVar(value="")
        self.var_grade_year = tk.StringVar(value="2025 - 2026")
        self.var_grade_semester = tk.StringVar(value="1")
        self.var_grade_subject = tk.StringVar(value="")
        self.var_grade_class = tk.StringVar(value="")
        self.var_grade_num_qs = tk.StringVar(value="50")
        self.var_blank_wrong = tk.BooleanVar(value=True)
        self.var_pt_correct = tk.StringVar(value="")
        self.var_pt_wrong = tk.StringVar(value="0")
        self.var_ans = tk.StringVar()
        self.var_resp = tk.StringVar()
        self.var_outdir = tk.StringVar(value="KET_QUA_CHAM")

        right = ttk.LabelFrame(self.frm_main, text="Tùy chọn chấm", padding=utils.PAD_M)
        right.grid(row=0, column=1, sticky="nsew", padx=utils.PAD_S, pady=utils.PAD_S)
        right.columnconfigure(1, weight=1)

        rr = 0
        def _add_grade_entry(label, var, row, width=None):
            ttk.Label(right, text=label).grid(row=row, column=0, sticky="e", padx=utils.PAD_S, pady=utils.PAD_XS)
            ttk.Entry(right, textvariable=var, width=width).grid(row=row, column=1, sticky="ew", padx=utils.PAD_S, pady=utils.PAD_XS)

        _add_grade_entry("Trường:", self.var_grade_school, rr); rr+=1
        _add_grade_entry("Khoa:", self.var_grade_faculty, rr); rr+=1
        _add_grade_entry("Tên bài:", self.var_grade_examtitle, rr); rr+=1
        _add_grade_entry("Năm học:", self.var_grade_year, rr); rr+=1
        _add_grade_entry("Học kỳ:", self.var_grade_semester, rr, width=10); rr+=1
        _add_grade_entry("Môn học:", self.var_grade_subject, rr); rr+=1
        _add_grade_entry("Lớp:", self.var_grade_class, rr); rr+=1

        ttk.Separator(right, orient="horizontal").grid(row=rr, column=0, columnspan=2, sticky="ew", pady=utils.PAD_S); rr+=1

        ttk.Checkbutton(right, text="Coi bỏ trống là sai", variable=self.var_blank_wrong)\
            .grid(row=rr, column=1, sticky="w", padx=utils.PAD_S); rr+=1

        _add_grade_entry("Điểm/Câu đúng:", self.var_pt_correct, rr); rr+=1
        _add_grade_entry("Điểm/Câu sai:", self.var_pt_wrong, rr); rr+=1

        ttk.Label(right, text="Thư mục xuất:").grid(row=rr, column=0, sticky="e", padx=utils.PAD_S, pady=utils.PAD_XS)
        frm_out = ttk.Frame(right)
        frm_out.grid(row=rr, column=1, sticky="ew", padx=utils.PAD_S, pady=utils.PAD_XS)
        ttk.Entry(frm_out, textvariable=self.var_outdir).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(frm_out, text="📁", command=self.pick_outdir, width=3).pack(side=tk.RIGHT, padx=(2, 0))
        rr += 1

        self.btn_run = ttk.Button(right, text="🚀 Chấm điểm & Biểu đồ", command=self.run, state="disabled", bootstyle="success")
        self.btn_run.grid(row=rr, column=0, columnspan=2, sticky="ew", padx=utils.PAD_S, pady=utils.PAD_XS)
        rr += 1


        # --- Log ---
        logf = ttk.LabelFrame(self.frm_main, text="Nhật ký", padding=utils.PAD_S)
        logf.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=utils.PAD_S, pady=utils.PAD_S)
        logf.rowconfigure(0, weight=1)
        logf.columnconfigure(0, weight=1)

        self.log = Text(logf, height=6, wrap="word", font=("Consolas", 11))
        self.log.grid(row=0, column=0, sticky="nsew", padx=utils.PAD_XS, pady=utils.PAD_XS)
        vsb = ttk.Scrollbar(logf, orient="vertical", command=self.log.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=vsb.set)
        
        # Metadata caching
        self.metadata_map = {} # path -> metadata
        self.load_metadata_tree()

    def setup_results_tab(self, container):
        """Xây dựng giao diện tab Kết quả tổng hợp."""
        # --- Bộ lọc ---
        frm_filter = ttk.LabelFrame(container, text="Bộ lọc tìm kiếm", padding=utils.PAD_S)
        frm_filter.pack(fill=tk.X, padx=utils.PAD_M, pady=utils.PAD_S)

        # Sử dụng grid để bố trí bộ lọc gọn gàng
        ttk.Label(frm_filter, text="Năm học:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.var_f_year = tk.StringVar(value="Tất cả")
        self.cb_f_year = ttk.Combobox(frm_filter, textvariable=self.var_f_year, width=15, state="readonly")
        self.cb_f_year.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frm_filter, text="Học kỳ:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.var_f_sem = tk.StringVar(value="Tất cả")
        self.cb_f_sem = ttk.Combobox(frm_filter, textvariable=self.var_f_sem, width=10, state="readonly")
        self.cb_f_sem.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(frm_filter, text="Môn học:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.var_f_subj = tk.StringVar(value="Tất cả")
        self.cb_f_subj = ttk.Combobox(frm_filter, textvariable=self.var_f_subj, width=25, state="readonly")
        self.cb_f_subj.grid(row=0, column=5, padx=5, pady=5, sticky="w")

        ttk.Label(frm_filter, text="Lớp:").grid(row=0, column=6, padx=5, pady=5, sticky="e")
        self.var_f_class = tk.StringVar(value="Tất cả")
        self.cb_f_class = ttk.Combobox(frm_filter, textvariable=self.var_f_class, width=15, state="readonly")
        self.cb_f_class.grid(row=0, column=7, padx=5, pady=5, sticky="w")


        btn_search = ttk.Button(frm_filter, text="🔍 Tìm", command=self.search_results, bootstyle="primary")
        btn_search.grid(row=0, column=8, padx=15, pady=5)
        
        btn_export = ttk.Button(frm_filter, text="📥 Xuất Excel", command=self.export_results, bootstyle="success")
        btn_export.grid(row=0, column=9, padx=5, pady=5)

        # --- Bảng kết quả ---
        frm_table = ttk.Frame(container)
        frm_table.pack(fill=tk.BOTH, expand=True, padx=utils.PAD_M, pady=utils.PAD_S)
        
        cols = ("id", "name", "mssv", "subject", "class", "year", "sem", "code", "score", "correct", "wrong", "date")
        self.res_tree = ttk.Treeview(frm_table, columns=cols, show="headings", bootstyle="info")
        
        # Headings
        headers = {
            "id": "ID", "name": "Họ và tên", "mssv": "MSSV", "subject": "Môn học",
            "class": "Lớp", "year": "Năm học", "sem": "HK", "code": "Mã đề",
            "score": "Điểm", "correct": "Đúng", "wrong": "Sai", "date": "Ngày chấm"
        }
        for c, h in headers.items():
            self.res_tree.heading(c, text=h)
            width = 100
            if c == "id": width = 50
            if c == "name": width = 200
            if c == "date": width = 150
            if c in ["sem", "code", "score", "correct", "wrong"]: width = 65
            self.res_tree.column(c, width=width, anchor=tk.CENTER if c not in ["name", "subject"] else tk.W)

        self.res_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(frm_table, orient="vertical", command=self.res_tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.res_tree.configure(yscrollcommand=vsb.set)

        # --- Thống kê ---
        self.lbl_res_stats = ttk.Label(container, text="Đang tải dữ liệu...", font=("Segoe UI", 10, "italic"))
        self.lbl_res_stats.pack(anchor="w", padx=utils.PAD_M, pady=utils.PAD_S)
        
        # Nút làm mới nhanh
        ttk.Button(container, text="🔄 Làm mới dữ liệu", command=self.refresh_and_search, bootstyle="outline-secondary").pack(side=tk.RIGHT, padx=utils.PAD_M, pady=utils.PAD_S)

        # Load dữ liệu mặc định sau khi tạo giao diện
        container.after(200, self.refresh_and_search)

    def refresh_filter_options(self):
        """Cập nhật danh sách các giá trị có trong CSDL vào Combobox."""
        try:
            data = db.get_unique_grading_values()
            
            def set_cb_values(cb, var, values):
                current = var.get()
                new_vals = ["Tất cả"] + sorted(list(set(values)))
                cb['values'] = new_vals
                if current not in new_vals:
                    var.set("Tất cả")

            set_cb_values(self.cb_f_year, self.var_f_year, data["years"])
            set_cb_values(self.cb_f_sem, self.var_f_sem, data["semesters"])
            set_cb_values(self.cb_f_subj, self.var_f_subj, data["subjects"])
            set_cb_values(self.cb_f_class, self.var_f_class, data["classes"])
        except Exception as e:
            print(f"Lỗi refresh_filter_options: {e}")

    def refresh_and_search(self):
        self.refresh_filter_options()
        self.search_results()


    def search_results(self):
        """Truy vấn kết quả từ DB theo bộ lọc."""
        self.res_tree.delete(*self.res_tree.get_children())
        
        def get_v(var):
            val = var.get().strip()
            return "" if val == "Tất cả" else val

        try:
            results = db.get_grading_history(
                school_year=get_v(self.var_f_year),
                semester=get_v(self.var_f_sem),
                subject=get_v(self.var_f_subj),
                class_name=get_v(self.var_f_class),
                limit=500 # Tăng giới hạn xem
            )

            
            for r in results:
                self.res_tree.insert("", "end", values=(
                    r.get("id", ""), r.get("student_name", ""), r.get("student_id", ""),
                    r.get("subject", ""), r.get("class_name", ""),
                    r.get("school_year", ""), r.get("semester", ""),
                    r.get("exam_code", ""), r.get("score", ""),
                    r.get("num_correct", ""), r.get("num_wrong", ""),
                    r.get("graded_at", "")
                ))
            
            # Cập nhật thống kê
            summary = db.get_grading_summary(
                school_year=get_v(self.var_f_year),
                semester=get_v(self.var_f_sem),
                subject=get_v(self.var_f_subj),
                class_name=get_v(self.var_f_class)
            )

            
            if summary and summary.get("total_students"):
                avg = summary.get("avg_score", 0) or 0
                mx = summary.get("max_score", 0) or 0
                mn = summary.get("min_score", 0) or 0
                total = summary.get("total_students", 0) or 0
                stats_text = (
                    f"Tổng: {total} bài  |  Môn: {summary.get('total_subjects', 0)}  |  Lớp: {summary.get('total_classes', 0)}  |  "
                    f"TB: {avg:.2f}  |  Cao nhất: {mx:.2f}  |  Thấp nhất: {mn:.2f}"
                )
                self.lbl_res_stats.config(text=stats_text, bootstyle="success")
            else:
                self.lbl_res_stats.config(text="Không tìm thấy dữ liệu phù hợp.", bootstyle="warning")
                
        except Exception as e:
            self.lbl_res_stats.config(text=f"Lỗi: {e}", bootstyle="danger")

    def export_results(self):
        """Xuất dữ liệu đang hiển thị trong TreeView ra file Excel."""
        items = self.res_tree.get_children()
        if not items:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất.")
            return
        
        rows_data = []
        for it in items:
            rows_data.append(self.res_tree.item(it)['values'])
        
        df_export = pd.DataFrame(rows_data, columns=[
            "ID", "Họ và tên", "MSSV", "Môn học", "Lớp", 
            "Năm học", "HK", "Mã đề", "Điểm", "Đúng", "Sai", "Ngày chấm"
        ])
        
        fn = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
            initialfile=f"Tong_hop_kq_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not fn: return
        
        try:
            df_export.to_excel(fn, index=False)
            messagebox.showinfo("Thành công", f"Đã xuất file thành công:\n{fn}")
            os.startfile(os.path.dirname(fn))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file: {e}")

    def load_metadata_tree(self):
        """Quét DE_TRON để lấy metadata.json và dựng cây."""
        self.tree.delete(*self.tree.get_children())
        self.metadata_map = {}
        
        base_dir = "DE_TRON"
        if not os.path.exists(base_dir):
            return
            
        # Cấu trúc: Năm -> Học kỳ -> Môn -> Lớp
        tree_data = {} 
        
        runs = [d for d in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, d))]
        for r in runs:
            meta_path = os.path.join(base_dir, r, "metadata.json")
            if os.path.exists(meta_path):
                try:
                    with open(meta_path, "r", encoding="utf-8") as f:
                        meta = json.load(f)
                    
                    year = meta.get("school_year", "Năm học khác")
                    semester = f"Học kỳ {meta.get('semester', '?')}"
                    subject = meta.get("subject", "Môn học khác")
                    class_name = meta.get("class_name", "Lớp khác")
                    
                    if year not in tree_data: tree_data[year] = {}
                    if semester not in tree_data[year]: tree_data[year][semester] = {}
                    if subject not in tree_data[year][semester]: tree_data[year][semester][subject] = {}
                    
                    # Một môn có thể có nhiều lớp (hoặc nhiều đợt thi cho cùng lớp)
                    if class_name not in tree_data[year][semester][subject]:
                        tree_data[year][semester][subject][class_name] = []
                    
                    tree_data[year][semester][subject][class_name].append({
                        "id": r,
                        "path": os.path.join(base_dir, r),
                        "meta": meta
                    })
                except Exception:
                    pass
            else:
                # Không có metadata -> đưa vào nhóm khác
                year = "Khác"
                if year not in tree_data: tree_data[year] = {"Không metadata": {"Khác": {"Khác": []}}}
                tree_data[year]["Không metadata"]["Khác"]["Khác"].append({
                    "id": r,
                    "path": os.path.join(base_dir, r),
                    "meta": {"class_name": r, "school_year": "Khác", "subject": "Khác", "semester": "Khác"}
                })

        # Đưa lên TreeView
        for year in sorted(tree_data.keys(), reverse=True):
            y_id = self.tree.insert("", "end", text=f"📅 {year}", open=True)
            for sem in sorted(tree_data[year].keys()):
                s_id = self.tree.insert(y_id, "end", text=f"🕒 {sem}")
                for subj in sorted(tree_data[year][sem].keys()):
                    sub_id = self.tree.insert(s_id, "end", text=f"📘 {subj}")
                    for cls in sorted(tree_data[year][sem][subj].keys()):
                        items = tree_data[year][sem][subj][cls]
                        for it in items:
                            # Nếu cùng lớp có nhiều đợt, hiển thị ID đợt
                            label = f"👥 {cls}" if len(items) == 1 else f"👥 {cls} ({it['id']})"
                            node_id = self.tree.insert(sub_id, "end", text=label, tags=("class_node",))
                            self.metadata_map[node_id] = it

    def on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        node_id = sel[0]
        if node_id in self.metadata_map:
            data = self.metadata_map[node_id]
            folder_path = data["path"]
            ans_file = os.path.join(folder_path, "dap_an.xlsx")
            
            # Luôn điền thông tin đợt thi từ metadata ngay khi chọn
            meta = data["meta"]
            self.var_grade_school.set(meta.get("school", "ĐẠI HỌC ĐIỆN LỰC"))
            self.var_grade_faculty.set(meta.get("faculty", "KHOA NĂNG LƯỢNG MỚI"))
            self.var_grade_examtitle.set(meta.get("exam_title", ""))
            self.var_grade_year.set(meta.get("school_year", ""))
            self.var_grade_semester.set(meta.get("semester", ""))
            self.var_grade_subject.set(meta.get("subject", ""))
            self.var_grade_class.set(meta.get("class_name", ""))
            self.var_grade_num_qs.set(str(meta.get("num_questions", 50)))
            
            # Tự động gợi ý thư mục xuất theo cấu trúc Năm/Kỳ/Môn/Lớp
            sub_path = os.path.join(
                "KET_QUA_CHAM",
                meta.get("school_year", "Khác").replace("/", "-").replace(" ", "_"),
                f"Ky_{meta.get('semester', 'Khac')}",
                meta.get("subject", "Khac")[:50].replace(" ", "_"),
                meta.get("class_name", "Khac")
            )
            self.var_outdir.set(os.path.join(os.getcwd(), sub_path))
            self.current_folder_path = folder_path  # Lưu đường dẫn đợt thi hiện tại
            self.logmsg(f"Đã nạp thông tin đợt thi: {data['id']}")

            if os.path.exists(ans_file):
                self.var_ans.set(ans_file)
                self.lbl_ans.config(text=f"{data['meta'].get('class_name')} - {data['id']}")
                self._toggle_run()
            else:
                self.var_ans.set("")
                self.lbl_ans.config(text="Chưa nạp đáp án", foreground="orange")
                self._toggle_run()
                # Không hiển thị thông báo lỗi/cảnh báo ở đây để tránh làm phiền người dùng khi chỉ muốn tạo script

    # ---------- tiện ích ----------
    def logmsg(self, s):
        """Ghi log vào khung văn bản ở tab Chấm điểm."""
        self.log.config(state="normal")
        self.log.insert(END, f"[{datetime.now().strftime('%H:%M:%S')}] {s}\n")
        self.log.see(END)
        self.log.config(state="disabled")


    def pick_ans(self):
        fn = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx;*.xls;*.csv")])
        if not fn: return
        self.var_ans.set(fn)
        self.lbl_ans.config(text=os.path.basename(fn))
        self._toggle_run()

    def pick_ans_from_manager(self):
        manager_dir = "DE_TRON"
        if not os.path.exists(manager_dir):
            messagebox.showinfo("Thông báo", "Chưa có đợt trộn đề nào.")
            return
            
        runs = [f for f in os.listdir(manager_dir) if os.path.isdir(os.path.join(manager_dir, f))]
        valid_runs = []
        for r in runs:
            if os.path.exists(os.path.join(manager_dir, r, "dap_an.xlsx")):
                valid_runs.append(r)
        
        valid_runs.sort(key=lambda x: os.path.getctime(os.path.join(manager_dir, x)), reverse=True)
        if not valid_runs:
            messagebox.showinfo("Thông báo", "Không tìm thấy file đáp án trong các đợt trộn đề.")
            return
            
        def on_select():
            sel = listbox.curselection()
            if not sel: return
            fname = os.path.join(manager_dir, valid_runs[sel[0]], "dap_an.xlsx")
            dlg.destroy()
            self.var_ans.set(fname)
            self.lbl_ans.config(text=f".../{valid_runs[sel[0]]}/dap_an.xlsx")
            self._toggle_run()

        dlg = tk.Toplevel(self.parent)
        dlg.title("Chọn đáp án từ đề đã trộn")
        utils.set_window_icon(dlg)
        dlg.geometry("400x300")
        dlg.transient(self.parent.winfo_toplevel())
        dlg.grab_set()

        ttk.Label(dlg, text="Chọn đợt trộn đề chứa đáp án:").pack(padx=10, pady=(10, 5), anchor="w")
        listbox = tk.Listbox(dlg)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        for r in valid_runs:
            listbox.insert(tk.END, r)
            
        ttk.Button(dlg, text="Chọn", command=on_select).pack(pady=10)

    def pick_resp(self):
        fn = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx;*.xls;*.csv")])
        if not fn: return
        self.var_resp.set(fn)
        self.lbl_resp.config(text=os.path.basename(fn))
        self._toggle_run()

    def pick_outdir(self):
        d = filedialog.askdirectory()
        if not d: return
        self.var_outdir.set(d)

    def _toggle_run(self):
        if self.var_ans.get().strip() and self.var_resp.get().strip():
            self.btn_run["state"] = "normal"
        else:
            self.btn_run["state"] = "disabled"

    # ---------- xử lý chính ----------
    def run(self):
        ans_path = self.var_ans.get().strip()
        resp_path = self.var_resp.get().strip()
        out_dir = self.var_outdir.get().strip()
        os.makedirs(out_dir, exist_ok=True)

        try:
            self.logmsg("Đang đọc file đáp án...")
            answer_key, num_qs_key = load_answer_key(ans_path)
            self.logmsg(f"- Đã nạp {len(answer_key)} mã đề, mỗi mã có tối đa {num_qs_key} câu.")

            self.logmsg("Đang đọc file trả lời...")
            resp_df, mapping = load_responses(resp_path)
            self.logmsg(f"- Phát hiện cột Mã đề: {mapping['exam_code_col']}")
            if mapping.get('student_id_col'):
                self.logmsg(f"- Phát hiện cột MSSV: {mapping['student_id_col']}")
            else:
                self.logmsg("- Không tìm thấy cột MSSV (không bắt buộc).")
            if mapping.get('student_name_col'):
                self.logmsg(f"- Phát hiện cột Họ & Tên: {mapping['student_name_col']}")
            else:
                self.logmsg("- Không tìm thấy cột Họ & Tên (không bắt buộc).")
            self.logmsg(f"- Nhận diện {len(mapping['question_cols'])} cột câu hỏi.")

            # --- KIỂM TRA TOÀN VẸN DỮ LIỆU (DATA AUDIT) ---
            audit_errors = []
            id_col = mapping.get('student_id_col')
            md_col = mapping.get('exam_code_col')
            
            # 1. Kiểm tra trùng MSSV
            if id_col:
                # Lấy danh sách MSSV (chuẩn hóa để check trùng)
                s_ids = resp_df[id_col].astype(str).str.strip().replace('nan', '')
                # Chỉ check trùng cho các MSSV không rỗng
                valid_ids = s_ids[s_ids != '']
                dupes = valid_ids[valid_ids.duplicated(keep=False)].unique()
                if len(dupes) > 0:
                    audit_errors.append(f"• Phát hiện {len(dupes)} MSSV trùng lặp (nộp bài nhiều lần).")
                    self.logmsg(f"⚠️ MSSV trùng lặp: {', '.join(map(str, dupes[:20]))}{'...' if len(dupes)>20 else ''}")
                    
                    # Kiểm tra mâu thuẫn mã đề cho cùng 1 MSSV
                    conflict_ids = []
                    for d_id in dupes:
                        subset = resp_df[s_ids == d_id]
                        assigned_codes = subset[md_col].unique()
                        if len(assigned_codes) > 1:
                            conflict_ids.append(f"{d_id} ({', '.join(map(str, assigned_codes))})")
                    if conflict_ids:
                        audit_errors.append(f"• Có {len(conflict_ids)} sinh viên nộp bài với CÁC MÃ ĐỀ KHÁC NHAU.")
                        self.logmsg(f"❌ Mâu thuẫn mã đề: {', '.join(conflict_ids[:10])}{'...' if len(conflict_ids)>10 else ''}")

            # 2. Kiểm tra mã đề không tồn tại trong đáp án
            if md_col:
                all_input_codes = resp_df[md_col].apply(_coerce_exam_code).unique()
                invalid_codes = [c for c in all_input_codes if c is not None and c not in answer_key]
                if invalid_codes:
                    audit_errors.append(f"• Phát hiện {len(invalid_codes)} mã đề KHÔNG CÓ trong file đáp án.")
                    self.logmsg(f"⚠️ Mã đề lạ (không có đáp án): {', '.join(map(str, invalid_codes))}")

            # Hiển thị cảnh báo tổng hợp
            if audit_errors:
                msg = "PHÁT HIỆN CÁC BẤT THƯỜNG TRONG FILE TRẢ LỜI:\n\n" + "\n".join(audit_errors) + \
                      "\n\nBạn có muốn TIẾP TỤC chấm điểm không?\n(Chọn 'Yes' để chấm tất cả, 'No' để dừng lại xử lý file)"
                if not messagebox.askyesno("Cảnh báo dữ liệu", msg):
                    self.logmsg("🛑 Đã dừng chấm điểm để người dùng kiểm tra lại file nguồn.")
                    return

            # Tính N thực tế từ key & responses để set mặc định điểm nếu cần
            dummy_df, dummy_stats, N = grade_responses(
                answer_key, num_qs_key, resp_df, mapping,
                treat_blank_as_wrong=self.var_blank_wrong.get(),
                point_per_correct=None,  # None -> 10/N
                point_per_wrong=0.0
            )
            # Lấy mặc định nếu người dùng chưa nhập
            try:
                if not self.var_pt_correct.get().strip():
                    self.var_pt_correct.set(str(round(10.0 / max(N,1), 6)))
            except Exception:
                pass

            # Parse điểm người dùng
            try:
                pt_corr = float(self.var_pt_correct.get().strip())
            except Exception:
                pt_corr = 10.0 / max(N,1)
            try:
                pt_wrong = float(self.var_pt_wrong.get().strip())
            except Exception:
                pt_wrong = 0.0

            self.logmsg(f"- Điểm mỗi câu đúng: {pt_corr}")
            self.logmsg(f"- Điểm mỗi câu sai : {pt_wrong}")
            self.logmsg("Đang chấm điểm...")

            df_detail, df_stats, N = grade_responses(
                answer_key, num_qs_key, resp_df, mapping,
                treat_blank_as_wrong=self.var_blank_wrong.get(),
                point_per_correct=pt_corr,
                point_per_wrong=pt_wrong
            )
                        # --- Thống kê nhanh cho Nhật ký ---
            try:
                avg_score = float(df_detail['Điểm'].mean())
                max_score = float(df_detail['Điểm'].max())
                min_score = float(df_detail['Điểm'].min())

                top_rows = df_detail[df_detail['Điểm'] == max_score]
                low_rows = df_detail[df_detail['Điểm'] == min_score]

                def _fmt_rows(rows, limit=10):
                    lines = []
                    for _, r in rows.head(limit).iterrows():
                        name = (str(r.get('Họ và tên', '')).strip() or '(không tên)')
                        mssv = str(r.get('Mã sinh viên', '')).strip() or str(r.get('MSSV', '')).strip()
                        ma_de = r.get('Mã đề', '')
                        diem = r.get('Điểm', '')
                        dung = r.get('Số lượng câu đúng', '')
                        sai  = r.get('Số lượng câu sai', '')
                        lines.append(f"- {name} ({mssv}) – Mã đề {ma_de} – Điểm {diem} – Đúng {dung} / Sai {sai}")
                    more = max(0, len(rows) - limit)
                    if more > 0:
                        lines.append(f"  ... và {more} bạn nữa cùng mức điểm này.")
                    return "\n".join(lines)

                self.logmsg("== Thống kê nhanh ==")
                self.logmsg(f"• Điểm trung bình: {avg_score:.3f}")
                self.logmsg(f"• Cao điểm nhất: {max_score:.3f}")
                self.logmsg(_fmt_rows(top_rows))
                self.logmsg(f"• Thấp điểm nhất: {min_score:.3f}")
                self.logmsg(_fmt_rows(low_rows))
                self.logmsg("== PHÁT TRIỂN BỞI VUPQ ==")
                self.logmsg("==========================")
            except Exception as e:
                self.logmsg(f"(Không tính được thống kê nhanh: {e})")


            # Ghi Excel
            ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            out_file = os.path.join(out_dir, f"Ket_qua_cham_diem_{ts}.xlsx")
            with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
                df_detail.to_excel(writer, index=False, sheet_name="ChiTiet")
                df_stats.sort_values("Câu", inplace=True)
                df_stats.to_excel(writer, index=False, sheet_name="ThongKe")
                cfg = {
                    "File dap an": [ans_path],
                    "File tra loi": [resp_path],
                    "Coi bo trong la sai": [self.var_blank_wrong.get()],
                    "So cot cau hoi phat hien": [len(mapping['question_cols'])],
                    "So cau N": [N],
                    "Diem moi cau dung": [pt_corr],
                    "Diem moi cau sai": [pt_wrong]
                }
                pd.DataFrame(cfg).to_excel(writer, index=False, sheet_name="CauHinh")

            self.logmsg(f"Đã ghi: {out_file}")
            
            # --- Lưu kết quả vào CSDL ---
            try:
                grade_year = self.var_grade_year.get().strip()
                grade_semester = self.var_grade_semester.get().strip()
                grade_subject = self.var_grade_subject.get().strip()
                grade_class = self.var_grade_class.get().strip()
                
                # Tìm session_id từ metadata_map (nếu chọn từ tree)
                session_id = None
                sel = self.tree.selection()
                if sel and sel[0] in self.metadata_map:
                    data = self.metadata_map[sel[0]]
                    folder = data.get("path", "")
                    session = db.find_session_by_folder(os.path.abspath(folder))
                    if not session:
                        session = db.find_session_by_folder(folder)
                    if session:
                        session_id = session["id"]
                        # Tự điền thông tin từ session nếu user chưa nhập
                        if not grade_year and session.get("school_year"):
                            grade_year = session["school_year"]
                        if not grade_subject and session.get("subject_name"):
                            grade_subject = session["subject_name"]
                        if not grade_class and session.get("class_name"):
                            grade_class = session["class_name"]
                        if not grade_semester and session.get("semester"):
                            grade_semester = session["semester"]
                
                # Chuẩn bị dữ liệu cho DB
                db_results = []
                for _, row in df_detail.iterrows():
                    answers_list = []
                    for q_num in range(1, N+1):
                        col = f'Câu {q_num}'
                        stu_ans = str(row.get(col, '')).strip()
                        code = row.get('Mã đề', 0)
                        key_arr = answer_key.get(code, [])
                        corr_ans = key_arr[q_num-1] if q_num-1 < len(key_arr) else ''
                        answers_list.append({
                            "question_num": q_num,
                            "student_answer": stu_ans,
                            "correct_answer": corr_ans,
                            "is_correct": stu_ans == corr_ans and stu_ans != ''
                        })
                    
                    db_results.append({
                        "student_name": str(row.get('Họ và tên', '')),
                        "student_id": str(row.get('Mã sinh viên', '') or row.get('MSSV', '')),
                        "exam_code": row.get('Mã đề', 0),
                        "score": float(row.get('Điểm', 0)),
                        "num_correct": int(row.get('Số lượng câu đúng', 0)),
                        "num_wrong": int(row.get('Số lượng câu sai', 0)),
                        "total_questions": N,
                        "answers": answers_list
                    })
                
                saved = db.save_grading_results(
                    session_id=session_id,
                    school_year=grade_year,
                    semester=grade_semester,
                    subject=grade_subject,
                    class_name=grade_class,
                    results=db_results,
                    result_file=out_file
                )
                self.logmsg(f"✅ Đã lưu {saved} kết quả vào cơ sở dữ liệu.")
            except Exception as e:
                self.logmsg(f"⚠️ Cảnh báo: Không thể lưu vào DB: {e}")
            
            # --- Hiển thị biểu đồ ---
            if HAS_MATPLOTLIB:
                self.show_score_chart(df_detail)
                
            messagebox.showinfo("Hoàn tất", f"Đã chấm xong. Mở file:\n{out_file}")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
            self.logmsg("Lỗi: " + str(e))

    def show_score_chart(self, df):
        """Mở cửa sổ hiển thị biểu đồ phổ điểm."""
        dlg = tk.Toplevel(self.parent)
        dlg.title("Phổ điểm bài thi")
        utils.set_window_icon(dlg)
        dlg.geometry("700x500")
        
        fig, ax = plt.subplots(figsize=(6, 4))
        scores = df['Điểm']
        
        # Vẽ Histogram
        ax.hist(scores, bins=10, range=(0, 10), color='#3498db', edgecolor='white', alpha=0.7)
        ax.set_title("Biểu đồ Phổ điểm", fontsize=14, fontweight='bold')
        ax.set_xlabel("Điểm số", fontsize=12)
        ax.set_ylabel("Số lượng sinh viên", fontsize=12)
        ax.set_xticks(range(0, 11))
        ax.grid(axis='y', linestyle='--', alpha=0.6)
        
        # Thêm đường trung bình
        avg = scores.mean()
        ax.axvline(avg, color='red', linestyle='dashed', linewidth=2, label=f'Trung bình: {avg:.2f}')
        ax.legend()

        canvas = FigureCanvasTkAgg(fig, master=dlg)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def show_grading_history(self):
        """Hiển thị lịch sử chấm điểm từ cơ sở dữ liệu."""
        dlg = tk.Toplevel(self.parent)
        dlg.title("Lịch sử chấm điểm (CSDL)")
        utils.set_window_icon(dlg)
        dlg.geometry("1100x650")
        dlg.transient(self.parent.winfo_toplevel())
        dlg.grab_set()
        
        # --- Bộ lọc ---
        frm_filter = ttk.LabelFrame(dlg, text="Bộ lọc")
        frm_filter.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(frm_filter, text="Năm học:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        var_f_year = tk.StringVar()
        ttk.Entry(frm_filter, textvariable=var_f_year, width=15).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frm_filter, text="Học kỳ:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        var_f_sem = tk.StringVar()
        ttk.Entry(frm_filter, textvariable=var_f_sem, width=5).grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(frm_filter, text="Môn học:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        var_f_subj = tk.StringVar()
        ttk.Entry(frm_filter, textvariable=var_f_subj, width=20).grid(row=0, column=5, padx=5, pady=5)
        
        ttk.Label(frm_filter, text="Lớp:").grid(row=0, column=6, padx=5, pady=5, sticky="e")
        var_f_class = tk.StringVar()
        ttk.Entry(frm_filter, textvariable=var_f_class, width=15).grid(row=0, column=7, padx=5, pady=5)
        
        # --- Thống kê tổng quan ---
        frm_stats = ttk.LabelFrame(dlg, text="Thống kê tổng quan")
        frm_stats.pack(fill=tk.X, padx=10, pady=(0, 5))
        lbl_stats = ttk.Label(frm_stats, text="", font=("Arial", 10))
        lbl_stats.pack(padx=10, pady=5, anchor="w")
        
        # --- Bảng kết quả ---
        frm_result = ttk.Frame(dlg)
        frm_result.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        result_cols = ("id", "student_name", "student_id", "subject", "class_name", 
                       "school_year", "semester", "exam_code", "score", "correct", "wrong", "graded_at")
        result_tree = ttk.Treeview(frm_result, columns=result_cols, show="headings", height=15)
        result_tree.heading("id", text="ID")
        result_tree.heading("student_name", text="Họ và tên")
        result_tree.heading("student_id", text="MSSV")
        result_tree.heading("subject", text="Môn học")
        result_tree.heading("class_name", text="Lớp")
        result_tree.heading("school_year", text="Năm học")
        result_tree.heading("semester", text="HK")
        result_tree.heading("exam_code", text="Mã đề")
        result_tree.heading("score", text="Điểm")
        result_tree.heading("correct", text="Đúng")
        result_tree.heading("wrong", text="Sai")
        result_tree.heading("graded_at", text="Ngày chấm")
        
        result_tree.column("id", width=35, anchor=tk.CENTER)
        result_tree.column("student_name", width=150, anchor=tk.W)
        result_tree.column("student_id", width=90, anchor=tk.W)
        result_tree.column("subject", width=150, anchor=tk.W)
        result_tree.column("class_name", width=100, anchor=tk.W)
        result_tree.column("school_year", width=90, anchor=tk.CENTER)
        result_tree.column("semester", width=30, anchor=tk.CENTER)
        result_tree.column("exam_code", width=55, anchor=tk.CENTER)
        result_tree.column("score", width=50, anchor=tk.CENTER)
        result_tree.column("correct", width=45, anchor=tk.CENTER)
        result_tree.column("wrong", width=40, anchor=tk.CENTER)
        result_tree.column("graded_at", width=120, anchor=tk.CENTER)
        
        vsb = ttk.Scrollbar(frm_result, orient="vertical", command=result_tree.yview)
        result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        result_tree.config(yscrollcommand=vsb.set)
        
        lbl_count = ttk.Label(dlg, text="", font=("Arial", 10))
        lbl_count.pack(padx=10, pady=(0, 5), anchor="w")
        
        def do_search():
            result_tree.delete(*result_tree.get_children())
            try:
                results = db.get_grading_history(
                    school_year=var_f_year.get().strip(),
                    semester=var_f_sem.get().strip(),
                    subject=var_f_subj.get().strip(),
                    class_name=var_f_class.get().strip()
                )
                for r in results:
                    result_tree.insert("", "end", values=(
                        r.get("id", ""), r.get("student_name", ""), r.get("student_id", ""),
                        r.get("subject", ""), r.get("class_name", ""),
                        r.get("school_year", ""), r.get("semester", ""),
                        r.get("exam_code", ""), r.get("score", ""),
                        r.get("num_correct", ""), r.get("num_wrong", ""),
                        r.get("graded_at", "")
                    ))
                lbl_count.config(text=f"Tìm thấy {len(results)} bản ghi")
                
                # Cập nhật thống kê
                summary = db.get_grading_summary(
                    school_year=var_f_year.get().strip(),
                    semester=var_f_sem.get().strip(),
                    subject=var_f_subj.get().strip(),
                    class_name=var_f_class.get().strip()
                )
                if summary and summary.get("total_students"):
                    avg = summary.get("avg_score", 0) or 0
                    mx = summary.get("max_score", 0) or 0
                    mn = summary.get("min_score", 0) or 0
                    total = summary.get("total_students", 0) or 0
                    stats_text = (
                        f"Tổng: {total} SV  |  "
                        f"Điểm TB: {avg:.2f}  |  "
                        f"Cao nhất: {mx:.2f}  |  "
                        f"Thấp nhất: {mn:.2f}  |  "
                        f"{summary.get('total_subjects', 0)} môn  |  "
                        f"{summary.get('total_classes', 0)} lớp"
                    )
                    lbl_stats.config(text=stats_text)
                else:
                    lbl_stats.config(text="Chưa có dữ liệu")
                    
            except Exception as e:
                messagebox.showerror("Lỗi", str(e), parent=dlg)
        
        ttk.Button(frm_filter, text="🔍 Lọc", command=do_search).grid(row=0, column=8, padx=10, pady=5)
        
        def do_export():
            """Xuất dữ liệu đang hiển thị trong TreeView ra file Excel."""
            items = result_tree.get_children()
            if not items:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất.")
                return
            
            rows_data = []
            for it in items:
                rows_data.append(result_tree.item(it)['values'])
            
            df_export = pd.DataFrame(rows_data, columns=[
                "ID", "Họ và tên", "MSSV", "Môn học", "Lớp", 
                "Năm học", "HK", "Mã đề", "Điểm", "Đúng", "Sai", "Ngày chấm"
            ])
            
            fn = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel file", "*.xlsx")],
                initialfile=f"Lich_su_cham_diem_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            if not fn: return
            
            try:
                df_export.to_excel(fn, index=False)
                messagebox.showinfo("Thành công", f"Đã xuất file thành công:\n{fn}")
                os.startfile(os.path.dirname(fn))
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file: {e}")

        ttk.Button(frm_filter, text="📥 Xuất Excel", command=do_export, bootstyle="outline-success").grid(row=0, column=9, padx=10, pady=5)
        
        # Tải dữ liệu lần đầu
        do_search()

def show_grading_history_for_session(parent, session):
    """Hiển thị lịch sử chấm điểm được lọc theo session cụ thể."""
    dlg = tk.Toplevel(parent)
    dlg.title(f"Kết quả chấm điểm: {session.get('class_name')} - {session.get('subject_name')}")
    utils.set_window_icon(dlg)
    dlg.geometry("1100x650")
    dlg.transient(parent.winfo_toplevel())
    dlg.grab_set()
    
    # --- Bảng kết quả ---
    frm_result = ttk.Frame(dlg)
    frm_result.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    result_cols = ("student_name", "student_id", "exam_code", "score", "correct", "wrong", "graded_at")
    result_tree = ttk.Treeview(frm_result, columns=result_cols, show="headings", height=15)
    result_tree.heading("student_name", text="Họ và tên")
    result_tree.heading("student_id", text="MSSV")
    result_tree.heading("exam_code", text="Mã đề")
    result_tree.heading("score", text="Điểm")
    result_tree.heading("correct", text="Đúng")
    result_tree.heading("wrong", text="Sai")
    result_tree.heading("graded_at", text="Ngày chấm")
    
    result_tree.column("student_name", width=200, anchor=tk.W)
    result_tree.column("student_id", width=120, anchor=tk.W)
    result_tree.column("exam_code", width=80, anchor=tk.CENTER)
    result_tree.column("score", width=80, anchor=tk.CENTER)
    result_tree.column("correct", width=80, anchor=tk.CENTER)
    result_tree.column("wrong", width=80, anchor=tk.CENTER)
    result_tree.column("graded_at", width=150, anchor=tk.CENTER)
    
    vsb = ttk.Scrollbar(frm_result, orient="vertical", command=result_tree.yview)
    result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    result_tree.config(yscrollcommand=vsb.set)
    
    # Load data
    try:
        results = db.get_grading_history_by_session(session["id"])
        for r in results:
            result_tree.insert("", "end", values=(
                r.get("student_name", ""), r.get("student_id", ""),
                r.get("exam_code", ""), r.get("score", ""),
                r.get("num_correct", ""), r.get("num_wrong", ""),
                r.get("graded_at", "")
            ))
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))
        
    def do_export():
        df_export = pd.DataFrame([result_tree.item(it)['values'] for it in result_tree.get_children()], 
                                columns=["Họ và tên", "MSSV", "Mã đề", "Điểm", "Đúng", "Sai", "Ngày chấm"])
        fn = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if fn:
            df_export.to_excel(fn, index=False)
            messagebox.showinfo("Thành công", "Đã xuất Excel.")

    ttk.Button(dlg, text="📥 Xuất Excel", command=do_export, bootstyle="success").pack(pady=10)
