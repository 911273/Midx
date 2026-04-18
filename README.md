# EPU Exam Manager - Midx

Phần mềm hỗ trợ quản lý ngân hàng câu hỏi và trộn đề thi tự động dành cho giảng viên EPU.

## Bản cập nhật: Initial commit - Fixed math rendering & Duplicate detection

Bản cập nhật này tập trung vào việc xử lý các công thức toán học phức tạp và cải thiện độ chính xác của hệ thống quản lý cơ sở dữ liệu.

### Các thay đổi chính:
- **Xử lý công thức toán học (OMML)**:
  - Hiển thị tóm tắt công thức (ví dụ: `sqrt(x)`) thay vì chỉ hiện `[Công thức]`.
  - Hỗ trợ hiển thị công thức trong danh sách câu hỏi và danh sách đáp án.
  - Sửa lỗi mất công thức khi mở trình soạn thảo câu hỏi.
- **Hệ thống nhận diện trùng lặp**:
  - Nâng cấp thuật toán băm (hashing) nội dung tính đến cả cấu trúc toán học XML.
  - Tự động chuẩn hóa và cập nhật mã băm khi mở Trung tâm rà soát CSDL.
  - Phân biệt chính xác các đáp án có cùng phần chữ nhưng khác phần công thức.
- **Tính năng mới trong CSDL**:
  - Chức năng tự động sửa lỗi mã băm cho các dữ liệu cũ.
  - Cải thiện tốc độ rà soát trùng lặp cho các ngân hàng câu hỏi lớn.

### Hướng dẫn cài đặt nhanh:
1. Đảm bảo đã cài đặt Python 3.10+
2. Cài đặt thư phư viện cần thiết: `pip install python-docx lxml pillow openpyxl`
3. Chạy ứng dụng: `python main.py`

---
*Phát triển bởi: 911273*
