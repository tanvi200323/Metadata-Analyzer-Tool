## 📁 Metadata Analyzer Pro

Metadata Analyzer Pro is a powerful and extensible GUI tool built with PyQt6 for analyzing file metadata. It supports image, PDF, DOCX, audio, and video files, and includes features like anomaly detection, steganography analysis, and entropy calculation.

---

### 🚀 Features

- 📂 Load individual files or entire folders  
- 📸 Extract EXIF metadata from images  
- 📕 Read PDF metadata and timestamps  
- 📝 Analyze DOCX files  
- 🎵🎬 Parse metadata from MP3/MP4 using mutagen  
- ⚠️ Detect anomalies like:
  - Signature mismatches
  - Suspicious author or tool info
  - Empty or unusually small files
  - Steganography indicators (via stegano)
- 🔍 Searchable metadata tree with highlights  
- 🌗 Toggle between light and dark UI themes  
- 📊 Entropy analysis for hidden/encrypted content  
- 🗺️ View GPS data on Google Maps  
- 📤 Export metadata to JSON  
- 🧪 View logical issues and anomalies in dialogs  

---

### 🛠️ Requirements

Install dependencies via pip:

bash
pip install PyQt6 pillow python-docx PyPDF2 mutagen python-magic-bin stegano


On Linux, you may also need:

bash
sudo apt install libmagic1


---

### 🧩 Optional Libraries

| Feature                    | Dependency        |
|----------------------------|-------------------|
| Image metadata             | Pillow            |
| PDF parsing                | PyPDF2            |
| DOCX parsing               | python-docx       |
| Audio/Video metadata       | mutagen           |
| File signature detection   | python-magic-bin  |
| Steganography (LSB) check  | stegano           |

---

### 🖥️ Running the Application

bash
python main.py


---

### 📦 Supported File Types

- Images: .jpg, .jpeg, .png, .bmp, .gif, .tiff  
- Documents: .pdf, .docx  
- Media: .mp3, .mp4

---

### 👨‍💻 Developer Notes

- Uses PyQt’s QTreeWidget, QStackedWidget, and custom delegates for UI.
- Supports dynamic theming and interactive context menus.
- Includes safe fallbacks for missing libraries.
- Highlights GPS coordinates with live Google Maps linking.