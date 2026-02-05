# Gemini Session Log
Date: 05/02/2026

## Summary of Work
We have completed a comprehensive update to the UniDocs application, focusing on adding new content, fixing mobile/tablet UI issues, and improving file compatibility across devices (especially iOS).

### 1. New Content Added
-   **Internship Subjects**:
    -   Added "Tập sự nghề nghiệp" (TSNN) and "Kiến tập công nghiệp" (KTCN).
    -   Restored "Cẩm nang thực tập".
    -   Added semester disclaimer ("HK2/2025-2026") for these subjects.
-   **English Certificates**:
    -   Added **Aptis ESOL** with full file listings.
-   **Skill Review**:
    -   Added **Ôn tập Kĩ năng thực hành chuyên môn** (KNTHCM) with 2 review documents.
-   **GitHub Integration**:
    -   Added a GitHub Star badge to the homepage.

### 2. UI/UX Improvements
-   **Responsive Design**:
    -   Adjusted sidebar breakpoints to `min-[1400px]` as requested.
    -   Added a "Close" button inside the sidebar for tablet users.
    -   Implemented a **Responsive Preview Modal** for screens smaller than 1300px, triggered by an "Eye" icon.
-   **Mobile Enhancements**:
    -   Fixed Master-Detail layout.
    -   Added hamburger menu toggle.

### 3. Technical Enhancements & Bug Fixes
-   **PDF Rendering (iOS Fix)**:
    -   Replaced `iframe` with **PDF.js**.
    -   Renders PDFs as HTML5 Canvases to fix the scroll issue on iOS devices.
    -   Desktop also uses the new div-based container logic.
-   **DOCX & Excel Preview**:
    -   Integrated `docx-preview` for native Word file viewing.
    -   Added `-webkit-overflow-scrolling: touch` to fix scrolling on iOS for DOCX and Excel.

### 4. Git Operations
-   Successfully added and pushed all changes to `origin main`.
