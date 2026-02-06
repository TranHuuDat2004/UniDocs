# unidocs - Kho tài liệu môn học 📚

**unidocs** là hệ thống lưu trữ và chia sẻ tài liệu học tập miễn phí, giúp sinh viên truy cập giáo trình, đề thi và bài giảng chất lượng cao một cách nhanh chóng và tiện lợi.

![unidocs Preview](img/screenshot.png)

## ✨ Tính năng nổi bật

-   **🗂️ Phân loại môn học thông minh**: Dễ dàng tìm kiếm môn học theo nhóm ngành (Công nghệ thông tin, Toán, Kỹ năng...).
-   **📱 Giao diện Mobile-First**: Trải nghiệm như ứng dụng di động:
    -   Menu trượt mượt mà.
    -   Chế độ xem tài liệu toàn màn hình.
    -   Tương tác chạm vuốt thân thiện.
-   **👁️ Xem trước tài liệu trực tiếp**: Hỗ trợ xem nhanh các định dạng phổ biến (PDF, Word, Excel, PowerPoint) ngay trên trình duyệt mà không cần tải về.
-   **⬇️ Tải xuống trọn bộ**: Tính năng nén và tải toàn bộ tài liệu của một môn học chỉ với 1 cú click (file .zip).
-   **🎨 Giao diện hiện đại**: Thiết kế với phong cách Glassmorphism, sạch sẽ và tập trung vào nội dung.

## 🛠️ Công nghệ sử dụng

Dự án được xây dựng hoàn toàn bằng **Vanilla Web Technologies**, đảm bảo tốc độ tải trang cực nhanh và dễ dàng triển khai:

-   **HTML5**: Cấu trúc ngữ nghĩa.
-   **Tailwind CSS**: Styling nhanh chóng và đẹp mắt.
-   **Vanilla JavaScript**: Xử lý logic, render dữ liệu động và tương tác UI.
-   **Phosphor Icons**: Bộ icon hiện đại, sắc nét.
-   **Thư viện hỗ trợ**:
    -   `JSZip`: Nén file phía client.
    -   `pptxjs`, `docx-preview`, `sheetjs`: Hỗ trợ xem trước file Office.

## 🚀 Hướng dẫn cài đặt & Sử dụng

Dự án này là **Static Web**, bạn không cần cài đặt backend hay database phức tạp.

1.  **Clone dự án**:
    ```bash
    git clone https://github.com/TranHuuDat2004/unidocs.git
    ```
2.  **Khởi chạy**:
    -   Dùng Live Server (VS Code) để mở file `index.html` .

## 🤝 Cấu trúc dữ liệu

Dữ liệu môn học được lưu trữ tại `data/subjects.json`. Để thêm môn học mới:

1.  Mở `data/subjects.json`.
2.  Thêm object mới vào mảng:
    ```json
    {
      "id": "ten_mon_hoc",
      "name": "Tên Môn Học",
      "description": "Mô tả ngắn...",
      "icon": "ph-books",
      "category": "programming_algo"
    }
    ```
3.  Tạo file chi tiết `data/details/ten_mon_hoc.json` chứa danh sách file.

## 👤 Tác giả

Developed by **TranHuuDat2004**.

---
*unidocs - Chia sẻ tri thức, kết nối thành công.*
