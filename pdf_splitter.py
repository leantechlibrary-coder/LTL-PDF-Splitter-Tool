import sys
import os
from pathlib import Path
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QScrollArea, QLabel, QFileDialog, QMessageBox, QStatusBar, QFrame, QGridLayout, QSlider
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QPixmap, QImage, QDragEnterEvent, QDropEvent


def show_about_dialog():
    """Aboutダイアログを表示"""
    text = (
        "PDF分割ツール v1.0\n\n"
        "【動作環境】\n"
        "Windows 11 (64bit)\n\n"
        "【重要】\n"
        "本ソフトウェアは法律専門職の業務効率化を目的としており、\n"
        "専門知識を前提とした設計です。\n\n"
        "【免責事項】\n"
        "本ソフトウェアは「現状有姿」(AS IS) で提供されます。\n"
        "- 不具合の修正、機能追加は行いません\n"
        "- 使用方法に関するサポートは行いません\n"
        "- 本ソフトウェアの使用により生じたいかなる損害についても、\n"
        "  開発者は一切の責任を負いません\n"
        "- 重要なファイルは必ずバックアップを取ってからご使用ください\n\n"
        "ご購入前に必ず取扱説明書をご確認ください。"
    )
    QMessageBox.about(None, "このソフトについて", text)


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
        print(f"ページ {self.page_num + 1} がクリックされました")  # デバッグ用
        self.selected = not self.selected
        print(f"選択状態: {self.selected}")  # デバッグ用
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
        self.setWindowTitle("PDF分割ツール")
        self.setGeometry(100, 100, 1000, 800)
        
        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウト
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        # ツールバー
        toolbar = QHBoxLayout()
        
        self.open_button = QPushButton("PDFを開く")
        self.open_button.clicked.connect(self.open_pdf)
        
        self.split_button = QPushButton("分割実行")
        self.split_button.clicked.connect(self.split_pdf)
        self.split_button.setEnabled(False)
        
        self.clear_button = QPushButton("クリア")
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
            "使い方：\n"
            "分割したい先頭ページをクリック（複数選択可）→「分割実行」で分割されます\n"
            "「クリア」で全選択解除\n"
            "※枝番にする証拠は、1枚ずつ分割してください（枝番ごとに1ファイルになります）"
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
        size_label = QLabel("サイズ:")
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
        column_label = QLabel("列数:")
        self.column_slider = QSlider(Qt.Orientation.Horizontal)
        self.column_slider.setMinimum(2)
        self.column_slider.setMaximum(6)
        self.column_slider.setValue(3)
        self.column_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.column_slider.setTickInterval(1)
        self.column_value_label = QLabel("3列")
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
        self.status_bar.showMessage("PDFファイルを開いてください")
        
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
        self.column_value_label.setText(f"{columns}列")
        
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
                QMessageBox.warning(self, "エラー", "PDFファイルをドロップしてください")
    
    def open_pdf(self):
        """PDFファイルを開く"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "PDFファイルを選択", "", "PDF Files (*.pdf)"
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
            
            self.status_bar.showMessage(f"読み込み中... {self.pdf_document.page_count}ページ")
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
                    self.status_bar.showMessage(f"読み込み中... {page_num + 1}/{self.pdf_document.page_count}ページ")
                    QApplication.processEvents()
            
            self.split_button.setEnabled(True)
            self.clear_button.setEnabled(True)
            self.status_bar.showMessage(f"{Path(file_path).name} を読み込みました ({self.pdf_document.page_count}ページ)")
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"PDFの読み込みに失敗しました:\n{str(e)}")
    
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
            self.status_bar.showMessage(f"選択中: {', '.join(map(str, selected_pages))}ページ目")
        else:
            self.status_bar.showMessage(f"{Path(self.pdf_path).name} を読み込みました ({self.pdf_document.page_count}ページ)")
    
    def split_pdf(self):
        """PDFを分割"""
        # 選択されたページを取得
        selected_pages = sorted([t.page_num for t in self.thumbnails if t.selected])
        
        if not selected_pages:
            QMessageBox.warning(self, "警告", "分割ポイントを選択してください")
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
        message = f"{len(ranges)}つのファイルに分割します。よろしいですか?\n\n"
        for i, (start, end) in enumerate(ranges, 1):
            message += f"ファイル{i}: {start + 1}〜{end + 1}ページ\n"
        
        reply = QMessageBox.question(
            self, "確認", message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        try:
            # 出力先フォルダを作成
            parent_dir = Path(self.pdf_path).parent
            base_name = Path(self.pdf_path).stem
            output_dir = parent_dir / f"{base_name}_分割"
            output_dir.mkdir(exist_ok=True)
            
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
                self, "完了",
                f"{len(ranges)}個のファイルに分割しました"
            )
            
            # 出力フォルダを開く
            if sys.platform == 'win32':
                os.startfile(output_dir)
            elif sys.platform == 'darwin':
                os.system(f'open "{output_dir}"')
            else:
                os.system(f'xdg-open "{output_dir}"')
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"分割に失敗しました:\n{str(e)}")


def main():
    app = QApplication(sys.argv)
    window = PDFSplitterWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
