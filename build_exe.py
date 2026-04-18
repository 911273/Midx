import os
import subprocess
import sys

def build():
    # Tên ứng dụng
    app_name = "QBM_System"
    main_script = "main.py"
    icon_path = os.path.join("assets", "logo.ico")
    
    # Các folder/file cần đính kèm (dạng "source;dest" trên Windows)
    # assets: thư mục chứa logo và icon
    data_files = [
        ("assets", "assets"),
    ]
    
    # Khởi tạo lệnh PyInstaller
    cmd = [
        "pyinstaller",
        "--noconsole",          # Không hiện cửa sổ terminal khi chạy
        "--onefile",            # Đóng gói thành 1 file duy nhất
        f"--name={app_name}",   # Tên file exe
    ]
    
    # Thêm icon nếu tồn tại
    if os.path.exists(icon_path):
        cmd.append(f"--icon={icon_path}")
    
    # Thêm dữ liệu assets
    for src, dst in data_files:
        cmd.append(f"--add-data={src};{dst}")
    
    # Thu thập toàn bộ thư viện cần thiết (đặc biệt là tkinterdnd2 và ttkbootstrap)
    cmd.append("--collect-all=tkinterdnd2")
    cmd.append("--collect-all=ttkbootstrap")
    
    # Một số hidden imports nếu cần
    cmd.append("--hidden-import=PIL._tkinter_finder")
    
    # Script chính
    cmd.append(main_script)
    
    print(f"--- Starting packaging for {app_name} ---")
    print(f"Running command: {' '.join(cmd)}")
    
    try:
        subprocess.check_call(cmd)
        print("\n" + "="*30)
        print(f"PACKAGING SUCCESSFUL!")
        print(f"EXE file located in: {os.path.join(os.getcwd(), 'dist')}")
        print("="*30)
    except subprocess.CalledProcessError as e:
        print(f"\nERROR during packaging: {e}")
    except Exception as e:
        print(f"\nUnknown error: {e}")

if __name__ == "__main__":
    build()
