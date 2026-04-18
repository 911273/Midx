import os
import ctypes
import tkinter as tk


try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    HAS_TTKBOOTSTRAP = True
except ImportError:
    from tkinter import ttk as tb
    HAS_TTKBOOTSTRAP = False

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

from tron_de import MixTab
from cham_diem import GradeTab
from ngan_hang import BankManagerTab
from quan_ly_de import ExamManagerTab
import utils
import db

def main():
    config = utils.load_config()

    # Khởi tạo cơ sở dữ liệu
    try:
        db.init_db()
    except Exception as e:
        print(f"Cảnh báo: Không thể khởi tạo CSDL: {e}")


    if HAS_TTKBOOTSTRAP:
        root = tb.Window(title="QBMSys", themename=config.get("theme", "litera"))
    else:
        root = tk.Tk()
        root.title("QBMSys")
    
    if HAS_DND:
        try:
            from tkinterdnd2 import TkinterDnD
            # Gắn dnd vào root hiện có của ttkbootstrap
            root.TkdndVersion = TkinterDnD._require(root)
        except Exception as e:
            print(f"Lỗi khởi tạo DnD: {e}")

    # Thiết lập logo phần mềm
    utils.set_window_icon(root)
                
    # Điều chỉnh scaling to rõ hơn (mặc định cho các màn hình có DPI cao)
    try:
        root.call('tk', 'scaling', 1.5)
    except Exception:
        pass

    # Kích thước mặc định và mở rộng tối đa
    utils.center_window(root, width=1280, height=750)
    root.minsize(1024, 650)
    root.resizable(True, True)
    try:
        root.state('zoomed')
    except Exception:
        pass

    # --- Header: Thanh Menu / Chuyển Theme ---
    header_frame = tb.Frame(root)
    header_frame.pack(side=tk.TOP, fill=tk.X, padx=utils.PAD_M, pady=(utils.PAD_M, 0))
    
    tb.Label(header_frame, text="QBM System", font=utils.FONT_HEADER).pack(side=tk.LEFT)
    
    def change_theme(theme_name):
        root.style.theme_use(theme_name)
        config["theme"] = theme_name
        utils.save_config(config)
        
    if HAS_TTKBOOTSTRAP:
        theme_menu = tb.Menubutton(header_frame, text="🎨 Giao diện", bootstyle="outline-secondary")
        theme_menu.pack(side=tk.RIGHT)
        
        menu = tk.Menu(theme_menu, tearoff=False)
        for t in ["litera", "cosmo", "flatly", "darkly", "superhero", "cyborg"]:
            menu.add_command(label=t.capitalize(), command=lambda x=t: change_theme(x))
        theme_menu["menu"] = menu

    # Thông tin tác giả: Pack trước để luôn bám đáy màn hình
    info_frame = tb.Frame(root)
    info_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=utils.PAD_M, pady=(0, utils.PAD_S))
 
    lbl_author = tb.Label(info_frame, text="VuPQ | vupq@epu.edu.vn", bootstyle="danger" if HAS_TTKBOOTSTRAP else None)
    if not HAS_TTKBOOTSTRAP: lbl_author.config(foreground="red")
    lbl_author.pack(side=tk.RIGHT)

    # Cấu hình Notebook
    nb = tb.Notebook(root, bootstyle="info" if HAS_TTKBOOTSTRAP else None)
    nb.pack(fill=tk.BOTH, expand=True, padx=utils.PAD_M, pady=utils.PAD_M)

    # Tabs
    tab_bank_frame = tb.Frame(nb)
    tab_mix_frame = tb.Frame(nb)
    tab_manager_frame = tb.Frame(nb)
    tab_grade_frame = tb.Frame(nb)

    # Thêm icon (emoji) vào tên Tab
    nb.add(tab_bank_frame, text=" 🏦 Ngân hàng ")
    nb.add(tab_mix_frame, text=" 🔀 Trộn đề ")
    nb.add(tab_manager_frame, text=" 📁 Quản lý đề ")
    nb.add(tab_grade_frame, text=" 📈 Chấm điểm ")

    # Khởi tạo nội dung từng tab
    BankManagerTab(tab_bank_frame)
    MixTab(tab_mix_frame)
    ExamManagerTab(tab_manager_frame)
    GradeTab(tab_grade_frame)

    root.mainloop()

if __name__ == "__main__":
    main()
