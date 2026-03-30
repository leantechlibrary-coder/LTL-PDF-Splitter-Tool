import sys
import os
from pathlib import Path
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QScrollArea, QLabel, QFileDialog, QMessageBox,
    QStatusBar, QFrame, QGridLayout, QSlider, QDialog, QTextEdit
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QPixmap, QImage, QDragEnterEvent, QDropEvent, QFont


class TextViewerDialog(QDialog):
    """Text viewer dialog"""

    def __init__(self, parent, title: str, content: str):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(620, 500)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(content)
        text_edit.setFont(QFont("Segoe UI", 9))
        text_edit.moveCursor(text_edit.textCursor().MoveOperation.Start)
        layout.addWidget(text_edit)

        close_btn = QPushButton("Close")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)


class AboutDialog(QDialog):
    """カスタムAboutダイアログ（操作説明書・README・ライセンス情報へのリンク付き）"""

    # --- 埋め込みテキスト定数 ---
    # MSIX / Microsoft Store 配布前提で改訂済み

    README_TEXT = (
        "================================================================\n"
        "LTL PDF Splitter\n"
        "README\n"
        "================================================================\n\n"
        "Thank you for using LTL PDF Splitter.\n\n"
        "This tool allows you to split multi-page PDF files\n"
        "into separate files at any page you choose.\n\n\n"
        "## System Requirements\n"
        "================================================================\n\n"
        "OS: Windows 10 / 11 (64-bit)\n"
        "Memory: 8GB or more recommended\n"
        "Storage: 500MB free space\n\n\n"
        "## Getting Started\n"
        "================================================================\n\n"
        "After installing from the Microsoft Store,\n"
        "launch the app from the Start menu.\n\n\n"
        "## Quick Start\n"
        "================================================================\n\n"
        "1. Click 'Open PDF' to select a file\n"
        "2. Click on pages where you want to split (multiple selections OK)\n"
        "3. Click 'Split' to execute\n\n\n"
        "## FAQ\n"
        "================================================================\n\n"
        "Q. Will the original PDF file be modified?\n"
        "A. No. Split files are always saved as new files.\n\n"
        "Q. Can I save to Google Drive?\n"
        "A. Yes. Select a Google Drive sync folder\n"
        "   as the output destination.\n\n\n"
        "## Important Notes\n"
        "================================================================\n\n"
        "- This software is provided as-is\n"
        "- Password-protected PDFs are not supported\n"
        "- Always back up important files before use\n\n\n"
        "## Disclaimer\n"
        "================================================================\n\n"
        "The developer assumes no liability for any damages\n"
        "arising from the use of this software.\n\n\n"
        "## Copyright and License\n"
        "================================================================\n\n"
        "Developer: Lean Tech Library\n\n"
        "This software is distributed under the AGPL-3.0 license.\n"
        "Please comply with the license terms when redistributing.\n\n"
        "Source code:\n"
        "https://github.com/leantechlibrary-coder/LTL-PDF-Splitter-Tool/tree/english\n"
    )

    MANUAL_TEXT = (
        "================================================================\n"
        "LTL PDF Splitter - User Guide\n"
        "================================================================\n\n"
        "## How to Use\n\n"
        "(1) Open a PDF file\n"
        "  - Click the 'Open PDF' button\n"
        "  - Or drag and drop a PDF file onto the window\n\n"
        "(2) Select split points\n"
        "  - Page thumbnails will be displayed\n"
        "  - Click on the first page of each section (a blue border appears)\n"
        "  - Multiple selections are possible (click again to deselect)\n\n"
        "(3) Execute split\n"
        "  - Click the 'Split' button\n"
        "  - Confirm in the dialog by clicking 'Yes'\n"
        "  - The output folder will open automatically after completion\n\n\n"
        "## Thumbnail Display Settings\n\n"
        "  - Use the sliders at the top right to adjust thumbnail size and columns\n"
        "  - Size: 100px - 500px\n"
        "  - Columns: 2 - 6\n\n\n"
        "## Output\n\n"
        "  - Output location: A '_split' folder is created in the same\n"
        "    directory as the original PDF file\n"
        "  - If the folder already exists, '_split_2', '_split_3', etc. will be used\n"
        "  - File names: originalname_001.pdf, originalname_002.pdf, ...\n\n\n"
        "## FAQ\n\n"
        "Q. Can I save to Google Drive?\n"
        "A. Yes. Select a Google Drive sync folder\n"
        "   as the output destination.\n\n"
        "Q. Will the original file be overwritten?\n"
        "A. No. Files are always saved as new files in a separate folder.\n\n"
        "Q. What happens if I split the same PDF twice?\n"
        "A. A new folder with '_split_2', '_split_3', etc. will be created.\n"
        "   Previous results will not be overwritten.\n\n"
        "Q. What about password-protected PDFs?\n"
        "A. Password-protected PDFs are not supported.\n"
        "   Please remove the password before processing.\n"
    )

    LICENSE_TEXT = (
        "================================================================================\n"
        "THIRD-PARTY SOFTWARE LICENSES\n"
        "LTL PDF Splitter\n"
        "================================================================================\n\n"
        "This software uses the following open-source software.\n"
        "License information is provided in accordance with each license.\n\n\n"
        "================================================================================\n"
        "1. PyMuPDF (fitz)\n"
        "================================================================================\n\n"
        "License: GNU Affero General Public License v3.0 (AGPL-3.0)\n"
        "Copyright: Artifex Software, Inc.\n"
        "Website: https://github.com/pymupdf/PyMuPDF\n\n"
        "Full license: https://www.gnu.org/licenses/agpl-3.0.txt\n\n\n"
        "================================================================================\n"
        "2. PyQt6\n"
        "================================================================================\n\n"
        "License: GNU General Public License v3.0 (GPL-3.0)\n"
        "Copyright: Riverbank Computing Limited\n"
        "Website: https://www.riverbankcomputing.com/software/pyqt/\n\n"
        "Full license: https://www.gnu.org/licenses/gpl-3.0.txt\n\n\n"
        "================================================================================\n"
        "3. Python\n"
        "================================================================================\n\n"
        "License: Python Software Foundation License (PSF)\n"
        "Copyright: Python Software Foundation\n"
        "Website: https://www.python.org/\n\n"
        "Full license: https://docs.python.org/3/license.html\n\n\n"
        "================================================================================\n"
        "This Software's License\n"
        "================================================================================\n\n"
        "This software (LTL PDF Splitter) is distributed under the\n"
        "GNU Affero General Public License v3.0 (AGPL-3.0).\n"
        "Please comply with the license terms when redistributing.\n\n"
        "Source code:\n"
        "https://github.com/leantechlibrary-coder/LTL-PDF-Splitter-Tool/tree/english\n\n\n"
        "================================================================================\n"
        "Disclaimer\n"
        "================================================================================\n\n"
        "This software is provided 'AS IS' without warranty of any kind.\n"
        "The developer assumes no liability for any damages arising\n"
        "from the use of this software.\n"
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About")
        self.resize(520, 480)
        self.setMinimumSize(400, 350)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(12)

        title_label = QLabel("LTL PDF Splitter v1.0")
        title_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        about_text = QTextEdit()
        about_text.setReadOnly(True)
        about_text.setFont(QFont("Segoe UI", 9))
        about_text.setPlainText(
            "System Requirements:\n"
            "Windows 10 / 11 (64-bit)\n\n"
            "Disclaimer:\n"
            "This software is provided 'AS IS'.\n"
            "The developer assumes no liability for any damages\n"
            "arising from the use of this software.\n"
            "Always back up important files before use.\n\n"
            "Developer:\n"
            "Lean Tech Library\n\n"
            "Please review the User Guide and README before use."
        )
        layout.addWidget(about_text)

        link_layout = QHBoxLayout()
        link_layout.setSpacing(8)

        manual_btn = QPushButton("User Guide")
        manual_btn.setToolTip("View the user guide")
        manual_btn.clicked.connect(self._show_manual)

        readme_btn = QPushButton("README")
        readme_btn.setToolTip("View the README")
        readme_btn.clicked.connect(self._show_readme)

        license_btn = QPushButton("Licenses")
        license_btn.setToolTip("View third-party license information")
        license_btn.clicked.connect(self._show_licenses)

        link_layout.addWidget(manual_btn)
        link_layout.addWidget(readme_btn)
        link_layout.addWidget(license_btn)
        layout.addLayout(link_layout)

        close_layout = QHBoxLayout()
        close_layout.addStretch()
        close_btn = QPushButton("Close")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        close_layout.addWidget(close_btn)
        close_layout.addStretch()
        layout.addLayout(close_layout)

    def _show_manual(self):
        dlg = TextViewerDialog(self, "User Guide", self.MANUAL_TEXT)
        dlg.exec()

    def _show_readme(self):
        dlg = TextViewerDialog(self, "README", self.README_TEXT)
        dlg.exec()

    def _show_licenses(self):
        dlg = TextViewerDialog(self, "Licenses", self.LICENSE_TEXT)
        dlg.exec()


def show_about_dialog():
    """Aboutダイアログを表示"""
    dlg = AboutDialog()
    dlg.exec()


class ThumbnailWidget(QFrame):
    """サムネイル表示用のウィジェット"""
    def __init__(self, page_num, pixmap, main_window, parent=None):
        super().__init__(parent)
        self.page_num = page_num
        self.selected = False
        self.main_window = main_window
        
        # フレームスタイルを設定
        self.setFrameShape(QFrame.Shape.Box)
        self.setLineWidth(2)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        
        # サムネイル画像
        self.image_label = QLabel()
        self.image_label.setPixmap(pixmap)
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.mousePressEvent = self.mousePressEvent  # クリックイベントを転送
        
        # ページ番号
        self.page_label = QLabel(f"{page_num + 1}")
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_label.setStyleSheet("font-size: 14pt; font-weight: bold;")
        self.page_label.mousePressEvent = self.mousePressEvent  # クリックイベントを転送
        
        layout.addWidget(self.image_label)
        layout.addWidget(self.page_label)
        self.setLayout(layout)
        
        self.update_style()
        
    def mousePressEvent(self, event):
        """クリックで選択状態を切り替え"""
        print(f"Page {self.page_num + 1} clicked")  # Debug
        self.selected = not self.selected
        print(f"Selected: {self.selected}")  # Debug
        self.update_style()
        self.main_window.update_status()
        
    def update_style(self):
        """選択状態に応じてスタイルを更新"""
        from PyQt6.QtGui import QPalette, QColor
        
        if self.selected:
            # 選択時：青い太枠と水色背景
            self.setLineWidth(5)
            palette = self.palette()
            palette.setColor(QPalette.ColorRole.Window, QColor("#E3F2FD"))
            self.setAutoFillBackground(True)
            self.setPalette(palette)
            self.setStyleSheet("QFrame { border: 5px solid #2196F3; }")
            self.page_label.setStyleSheet("font-size: 14pt; font-weight: bold; color: #2196F3;")
        else:
            # 非選択時：細いグレー枠と白背景
            self.setLineWidth(2)
            palette = self.palette()
            palette.setColor(QPalette.ColorRole.Window, QColor("white"))
            self.setAutoFillBackground(True)
            self.setPalette(palette)
            self.setStyleSheet("QFrame { border: 2px solid #CCCCCC; }")
            self.page_label.setStyleSheet("font-size: 14pt; font-weight: bold; color: black;")

class PDFSplitterWindow(QMainWindow):
    """PDF分割ツールのメインウィンドウ"""
    def __init__(self):
        super().__init__()
        self.pdf_path = None
        self.pdf_document = None
        self.thumbnails = []
        
        self.init_ui()
        
    def init_ui(self):
        """UIの初期化"""
        self.setWindowTitle("LTL PDF Splitter")
        self.setGeometry(100, 100, 1000, 800)
        
        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウト
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        # ツールバー
        toolbar = QHBoxLayout()
        
        self.open_button = QPushButton("Open PDF")
        self.open_button.clicked.connect(self.open_pdf)
        
        self.split_button = QPushButton("Split")
        self.split_button.clicked.connect(self.split_pdf)
        self.split_button.setEnabled(False)
        
        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_selection)
        self.clear_button.setEnabled(False)
        
        toolbar.addWidget(self.open_button)
        toolbar.addWidget(self.split_button)
        toolbar.addWidget(self.clear_button)
        toolbar.addStretch()
        
        # Aboutリンク
        about_label = QLabel('<a href="#" style="color: #888;">About</a>')
        about_label.setOpenExternalLinks(False)
        about_label.linkActivated.connect(lambda: show_about_dialog())
        toolbar.addWidget(about_label)
        
        main_layout.addLayout(toolbar)
        
        # 使い方説明とサムネイル調整
        help_and_control_layout = QHBoxLayout()
        
        help_label = QLabel(
            "How to use:\n"
            "Click on the first page of each section to set split points (multiple selections OK)\n"
            "Click 'Split' to split the PDF. Click 'Clear' to deselect all.\n"
            "To split another file, just click 'Open PDF' again."
        )
        help_label.setStyleSheet("""
            QLabel {
                background-color: #FFF9C4;
                padding: 10px;
                border: 1px solid #FBC02D;
                font-size: 11pt;
            }
        """)
        help_and_control_layout.addWidget(help_label, stretch=3)
        
        # サムネイル調整コントロール
        control_widget = QWidget()
        control_widget.setStyleSheet("""
            QWidget {
                background-color: #E8F5E9;
                padding: 10px;
                border: 1px solid #4CAF50;
            }
        """)
        control_layout = QVBoxLayout()
        control_layout.setContentsMargins(5, 5, 5, 5)
        
        # サイズ調整
        size_layout = QHBoxLayout()
        size_label = QLabel("Size:")
        self.size_slider = QSlider(Qt.Orientation.Horizontal)
        self.size_slider.setMinimum(100)
        self.size_slider.setMaximum(500)
        self.size_slider.setValue(250)
        self.size_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.size_slider.setTickInterval(100)
        self.size_value_label = QLabel("250px")
        self.size_value_label.setMinimumWidth(50)
        
        size_layout.addWidget(size_label)
        size_layout.addWidget(self.size_slider)
        size_layout.addWidget(self.size_value_label)
        
        # 列数調整
        column_layout = QHBoxLayout()
        column_label = QLabel("Columns:")
        self.column_slider = QSlider(Qt.Orientation.Horizontal)
        self.column_slider.setMinimum(2)
        self.column_slider.setMaximum(6)
        self.column_slider.setValue(3)
        self.column_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.column_slider.setTickInterval(1)
        self.column_value_label = QLabel("3 cols")
        self.column_value_label.setMinimumWidth(50)
        
        column_layout.addWidget(column_label)
        column_layout.addWidget(self.column_slider)
        column_layout.addWidget(self.column_value_label)
        
        control_layout.addLayout(size_layout)
        control_layout.addLayout(column_layout)
        control_widget.setLayout(control_layout)
        
        help_and_control_layout.addWidget(control_widget, stretch=2)
        
        main_layout.addLayout(help_and_control_layout)
        
        # スライダーの変更イベントを接続
        self.size_slider.valueChanged.connect(self.on_slider_changed)
        self.column_slider.valueChanged.connect(self.on_slider_changed)
        
        # スクロールエリア
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        # サムネイルコンテナ
        self.thumbnail_container = QWidget()
        self.thumbnail_layout = QGridLayout()
        self.thumbnail_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.thumbnail_layout.setSpacing(10)
        self.thumbnail_container.setLayout(self.thumbnail_layout)
        
        scroll_area.setWidget(self.thumbnail_container)
        main_layout.addWidget(scroll_area)
        
        # ステータスバー
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Open a PDF file to get started")
        
        # ドラッグ&ドロップを有効化
        self.setAcceptDrops(True)
        
        # 最大化
        self.showMaximized()
        
    def on_slider_changed(self):
        """スライダーが変更された時にラベルを更新し、PDFを再読み込み"""
        # ラベルを更新
        size = self.size_slider.value()
        columns = self.column_slider.value()
        self.size_value_label.setText(f"{size}px")
        self.column_value_label.setText(f"{columns} cols")
        
        # PDFが読み込まれている場合は再読み込み
        if self.pdf_path and self.pdf_document:
            self.reload_pdf()
    
    def reload_pdf(self):
        """現在のPDFを新しい設定で再読み込み"""
        current_path = self.pdf_path
        selected_pages = [t.page_num for t in self.thumbnails if t.selected]
        
        # サムネイルをクリア
        for thumbnail in self.thumbnails:
            self.thumbnail_layout.removeWidget(thumbnail)
            thumbnail.deleteLater()
        self.thumbnails.clear()
        
        # 新しい設定で再生成
        size = self.size_slider.value()
        columns = self.column_slider.value()
        
        for page_num in range(self.pdf_document.page_count):
            page = self.pdf_document[page_num]
            
            # 指定されたサイズでレンダリング
            zoom = size / page.rect.width
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # QPixmapに変換
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(img)
            
            # サムネイルウィジェットを作成
            thumbnail = ThumbnailWidget(page_num, pixmap, self)
            
            # 以前の選択状態を復元
            if page_num in selected_pages:
                thumbnail.selected = True
                thumbnail.update_style()
            
            self.thumbnails.append(thumbnail)
            
            # グリッド配置
            row = page_num // columns
            col = page_num % columns
            self.thumbnail_layout.addWidget(thumbnail, row, col)
        
        self.update_status()
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """ドラッグされたファイルを受け入れる"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            
    def dropEvent(self, event: QDropEvent):
        """ドロップされたファイルを処理"""
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith('.pdf'):
                self.load_pdf(file_path)
            else:
                QMessageBox.warning(self, "Error", "Please drop a PDF file")
    
    def open_pdf(self):
        """PDFファイルを開く"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select a PDF file", "", "PDF Files (*.pdf)"
        )
        if file_path:
            self.load_pdf(file_path)
    
    def load_pdf(self, file_path):
        """PDFを読み込んでサムネイルを表示"""
        try:
            # 既存のサムネイルをクリア
            self.clear_thumbnails()
            
            # PDFを開く
            self.pdf_path = file_path
            self.pdf_document = fitz.open(file_path)
            
            self.status_bar.showMessage(f"Loading... {self.pdf_document.page_count} pages")
            QApplication.processEvents()
            
            # スライダーの値を取得
            size = self.size_slider.value()
            columns = self.column_slider.value()
            
            # サムネイルを生成
            for page_num in range(self.pdf_document.page_count):
                page = self.pdf_document[page_num]
                
                # 指定されたサイズでレンダリング
                zoom = size / page.rect.width
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)
                
                # QPixmapに変換
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                
                # サムネイルウィジェットを作成
                thumbnail = ThumbnailWidget(page_num, pixmap, self)
                self.thumbnails.append(thumbnail)
                
                # グリッド配置（行, 列を計算）
                row = page_num // columns
                col = page_num % columns
                self.thumbnail_layout.addWidget(thumbnail, row, col)
                
                # 進行状況を更新
                if (page_num + 1) % 10 == 0:
                    self.status_bar.showMessage(f"Loading... {page_num + 1}/{self.pdf_document.page_count} pages")
                    QApplication.processEvents()
            
            self.split_button.setEnabled(True)
            self.clear_button.setEnabled(True)
            self.status_bar.showMessage(f"Loaded {Path(file_path).name} ({self.pdf_document.page_count} pages)")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load PDF:\n{str(e)}")
    
    def clear_thumbnails(self):
        """サムネイルをすべてクリア"""
        for thumbnail in self.thumbnails:
            self.thumbnail_layout.removeWidget(thumbnail)
            thumbnail.deleteLater()
        self.thumbnails.clear()
        
        if self.pdf_document:
            self.pdf_document.close()
            self.pdf_document = None
    
    def clear_selection(self):
        """選択をすべてクリア"""
        for thumbnail in self.thumbnails:
            if thumbnail.selected:
                thumbnail.selected = False
                thumbnail.update_style()
        self.update_status()
    
    def update_status(self):
        """ステータスバーを更新"""
        selected_pages = [t.page_num + 1 for t in self.thumbnails if t.selected]
        if selected_pages:
            selected_pages.sort()
            self.status_bar.showMessage(f"Selected: page {', '.join(map(str, selected_pages))}")
        else:
            self.status_bar.showMessage(f"Loaded {Path(self.pdf_path).name} ({self.pdf_document.page_count} pages)")
    
    def split_pdf(self):
        """PDFを分割"""
        # 選択されたページを取得
        selected_pages = sorted([t.page_num for t in self.thumbnails if t.selected])
        
        if not selected_pages:
            QMessageBox.warning(self, "Warning", "Please select at least one split point")
            return
        
        # 分割範囲を計算
        ranges = []
        start = 0
        for page_num in selected_pages:
            if page_num > start:
                ranges.append((start, page_num - 1))
            start = page_num
        ranges.append((start, self.pdf_document.page_count - 1))
        
        # 確認ダイアログ
        message = f"Split into {len(ranges)} files. Proceed?\n\n"
        for i, (start, end) in enumerate(ranges, 1):
            message += f"File {i}: pages {start + 1} - {end + 1}\n"
        
        reply = QMessageBox.question(
            self, "Confirm", message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        try:
            # 出力先フォルダを作成（同名フォルダがある場合は連番を付与）
            parent_dir = Path(self.pdf_path).parent
            base_name = Path(self.pdf_path).stem
            output_dir = parent_dir / f"{base_name}_split"
            if output_dir.exists():
                counter = 2
                while True:
                    output_dir = parent_dir / f"{base_name}_split_{counter}"
                    if not output_dir.exists():
                        break
                    counter += 1
            output_dir.mkdir()
            
            # 分割実行
            output_files = []
            for i, (start, end) in enumerate(ranges, 1):
                output_path = output_dir / f"{base_name}_{i:03d}.pdf"
                
                # 新しいPDFを作成
                new_pdf = fitz.open()
                new_pdf.insert_pdf(self.pdf_document, from_page=start, to_page=end)
                new_pdf.save(output_path)
                new_pdf.close()
                
                output_files.append(output_path)
            
            # 完了メッセージ
            QMessageBox.information(
                self, "Complete",
                f"Successfully split into {len(ranges)} files"
            )
            
            # 出力フォルダを開く
            if sys.platform == 'win32':
                os.startfile(output_dir)
            elif sys.platform == 'darwin':
                os.system(f'open "{output_dir}"')
            else:
                os.system(f'xdg-open "{output_dir}"')
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to split PDF:\n{str(e)}")


def main():
    app = QApplication(sys.argv)
    window = PDFSplitterWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
