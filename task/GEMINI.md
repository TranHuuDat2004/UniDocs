# Gemini Session Log
Date: 05/02/2026

## Summary of Work
We have completed a comprehensive update to the UniDocs application, adding a wide range of new lab subjects and enhancing the UI/UX for better accessibility and mobile responsiveness.

### 1. New Content Added
-   **Lab Subjects**:
    -   **Thực hành Giải tích với Python** (Calculus Python).
    -   **Phương pháp lập trình với Ngôn ngữ C** (Methodology C) - included Assignment file.
    -   **Thực hành Tổ chức máy tính** (Computer Org Lab) - converted DOCX to PDF, renamed files.
    -   **Thực hành Đại số tuyến tính với Python** (Linear Algebra Python).
    -   **Thực hành Hướng đối tượng** (OOP Lab).
    -   **Thực hành Hệ điều hành** (OS Lab).
    -   **Thực hành CTDLGT** (DSA Lab).
    -   **Thực hành Cơ sở dữ liệu** (SQL Lab) - converted DOCX to PDF.
-   **Internship Subjects**:
    -   Added "Tập sự nghề nghiệp" (TSNN) and "Kiến tập công nghiệp" (KTCN).
    -   Restored "Cẩm nang thực tập".
-   **English Certificates**:
    -   Added **Aptis ESOL**.
-   **Skill Review**:
    -   Added **Ôn tập Kĩ năng thực hành chuyên môn** (KNTHCM).
    -   Added **Thái độ sống 2** (Thai Do Song 2) to new **Skills** category.
    -   Added **Thái độ sống 1** (Thai Do Song 1).
    -   Added **Thái độ sống 3** (Thai Do Song 3) and updated with 5 new documents.
    -   Added **Kĩ năng đưa ra quyết định** (Decision Making).
    -   Added **Kĩ năng Kaizen & 5S** (Kaizen & 5S) and updated with 4 new documents.
    -   Added **Tư duy phản biện** (Critical Thinking).
    -   Added **Kĩ năng tự học** (Self-learning).
    -   Organized internship subjects into new **Thực tập** (Internship) category.

### 2. UI/UX Improvements
-   **Responsive Design**:
    -   Adjusted sidebar breakpoints to `min-[1400px]`.
    -   Implemented a **Responsive Preview Modal** for tablet/laptop screens (<1300px).
-   **Mobile Enhancements**:
    -   Fixed Master-Detail layout and added hamburger menu toggle.
    -   **Clickable File List**: Made the entire file row clickable for easier preview access (Desktop & Mobile).
-   **Notification Feature**:
    -   Added blue notification box for external links (Google Drive) in subject headers.

### 4. Git Operations
-   Successfully added and pushed all changes to `origin main`.

### 3. Technical Enhancements
-   **File Processing**:
    -   Created scripts to **convert DOCX to PDF** (`convert_docx_to_pdf.py`).
    -   Created scripts to **rename files** to remove Vietnamese accents (`rename_files_unsigned.py`) to fix web display issues.
-   **PDF Rendering**:
    -   Integrated **PDF.js** to fix scrolling issues on iOS devices.
    -   **Fixed Race Condition**: Implemented `currentRenderId` check to prevent overlapping pages when switching files quickly.
-   **DOCX Preview**:
    -   Integrated `docx-preview` library.

### 4. Git Operations
-   Successfully added and pushed all changes to `origin main`.
