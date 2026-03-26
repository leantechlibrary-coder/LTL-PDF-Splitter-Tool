import sys
import os
import json
import webbrowser
from pathlib import Path
from datetime import datetime
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QScrollArea, QLabel, QFileDialog, QMessageBox,
    QStatusBar, QFrame, QGridLayout, QSlider, QDialog, QTextEdit
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QPixmap, QImage, QDragEnterEvent, QDropEvent, QFont


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# トライアル管理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TRIAL_DAYS = 7
# 有料版のMicrosoft StoreページURL
STORE_URL = "https://apps.microsoft.com/detail/9pm38hpwfngj?hl=ja-JP&gl=JP"


class TrialManager:
    """無料トライアルの期限を管理する。

    初回起動日を %APPDATA%/LeanTechLibrary/ に記録し、
    経過日数を返す。アンインストールしても記録は残る。
    """

    def __init__(self):
        appdata = os.environ.get("APPDATA", str(Path.home()))
        self._dir = Path(appdata) / "LeanTechLibrary"
        self._file = self._dir / "trial_pdf_splitter.json"

    def _read(self) -> dict:
        try:
            return json.loads(self._file.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _write(self, data: dict):
        self._dir.mkdir(parents=True, exist_ok=True)
        self._file.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def first_launch_date(self) -> datetime:
        """初回起動日を返す。未記録なら今日を記録して返す。"""
        data = self._read()
        if "first_launch" in data:
            return datetime.fromisoformat(data["first_launch"])
        now = datetime.now()
        data["first_launch"] = now.isoformat()
        self._write(data)
        return now

    def days_remaining(self) -> int:
        """トライアル残り日数を返す（0以下 = 期限切れ）。"""
        first = self.first_launch_date()
        elapsed = (datetime.now() - first).days
        return TRIAL_DAYS - elapsed

    def is_expired(self) -> bool:
        return self.days_remaining() <= 0


class TrialDialog(QDialog):
    """起動時に表示するトライアル情報ダイアログ。"""

    def __init__(self, days_remaining: int, expired: bool, parent=None):
        super().__init__(parent)
        self.expired = expired
        self.setWindowTitle("無料トライアル版 — PDF分割ツール")
        self.setMinimumWidth(460)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 16)
        layout.setSpacing(14)

        # --- タイトル ---
        title = QLabel("PDF分割ツール — 無料トライアル版")
        title.setFont(QFont("Yu Gothic UI", 13, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # --- 残り日数 / 期限切れメッセージ ---
        if expired:
            msg = QLabel(
                "無料トライアル期間が終了しました。\n"
                "ご利用いただきありがとうございました。\n\n"
                "引き続きご利用いただくには、\n"
                "Microsoft Store で有料版をご購入ください。"
            )
            msg.setStyleSheet(
                "color: #C62828; font-size: 11pt; padding: 8px;"
            )
        else:
            msg = QLabel(
                f"無料トライアル終了まで あと {days_remaining} 日 です。\n\n"
                "トライアル期間終了後もご利用いただくには、\n"
                "Microsoft Store で有料版をご購入ください。"
            )
            msg.setStyleSheet("font-size: 11pt; padding: 8px;")

        msg.setAlignment(Qt.AlignmentFlag.AlignCenter)
        msg.setWordWrap(True)
        layout.addWidget(msg)

        # --- Storeリンクボタン ---
        store_btn = QPushButton("Microsoft Store で有料版を見る")
        store_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078D4;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #106EBE; }
        """)
        store_btn.clicked.connect(self._open_store)
        layout.addWidget(store_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        # --- 閉じる / 続けるボタン ---
        if expired:
            close_btn = QPushButton("閉じる")
            close_btn.clicked.connect(self.reject)
        else:
            close_btn = QPushButton("閉じて続ける")
            close_btn.clicked.connect(self.accept)

        close_btn.setFixedWidth(160)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        # --- フッター ---
        footer = QLabel("開発・販売：Lean Tech Library")
        footer.setStyleSheet("color: #888; font-size: 9pt;")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(footer)

    def _open_store(self):
        """Microsoft Store の有料版ページを開く。"""
        webbrowser.open(STORE_URL)


class TextViewerDialog(QDialog):
    """テキスト全文表示用の子ダイアログ"""

    def __init__(self, parent, title: str, content: str):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(620, 500)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(content)
        text_edit.setFont(QFont("Yu Gothic UI", 9))
        text_edit.moveCursor(text_edit.textCursor().MoveOperation.Start)
        layout.addWidget(text_edit)

        close_btn = QPushButton("閉じる")
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
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "PDF証拠整理ツール\n"
        "README\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "この度はPDF証拠整理ツールをご利用いただき、\n"
        "誠にありがとうございます。\n\n"
        "本ツールは、訴訟・紛争案件における証拠整理業務を\n"
        "効率化するために開発された専用ツールです。\n\n\n"
        "■ 収録ツール\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "・PDF分割ツール\n"
        "  PDFファイルを複数の分割ポイントで一度に分割\n\n"
        "・証拠番号付与ツール\n"
        "  PDFファイルに証拠番号（甲第○号証等）を自動付与\n\n\n"
        "■ 動作環境\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "OS：Windows 10 / 11（64bit）\n"
        "メモリ：8GB以上推奨\n"
        "ストレージ：500MB以上の空き容量\n\n\n"
        "■ 起動方法\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Microsoft Storeからインストール後、\n"
        "スタートメニューから起動してください。\n\n\n"
        "■ クイックスタート\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "＜PDF分割ツール＞\n"
        "1. 「PDFを開く」で対象ファイルを選択\n"
        "2. 分割したい先頭ページをクリック（複数選択可）\n"
        "3. 「分割実行」をクリック\n\n"
        "＜証拠番号付与ツール＞\n"
        "1. 「フォルダを開く」でPDFファイルを読み込み\n"
        "2. ファイルの順番を調整（ドラッグ＆ドロップ）\n"
        "3. 証拠種別（甲/乙等）とフォント設定\n"
        "4. 「証拠番号を付与して保存」をクリック\n\n\n"
        "■ よくある質問\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Q. 元のPDFファイルが変更されることはありますか？\n"
        "A. ありません。常に新しいファイルとして保存されます。\n\n"
        "Q. Googleドライブに保存できますか？\n"
        "A. Googleドライブデスクトップの同期フォルダを\n"
        "   出力先に指定することで可能です。\n\n"
        "Q. 何号証まで対応していますか？\n"
        "A. システム上は9999号証まで対応しています。\n\n\n"
        "■ ご注意事項\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "・本ツールは現状有姿での提供となります\n"
        "・パスワード保護されたPDFには対応していません\n"
        "・重要なファイルは必ずバックアップを取ってからご使用ください\n\n\n"
        "■ 免責事項\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "本ソフトウェアの使用により生じたいかなる損害についても、\n"
        "開発者は一切の責任を負いかねます。\n\n\n"
        "■ 著作権とライセンス\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "開発・販売：Lean Tech Library\n\n"
        "本ソフトウェアはAGPL-3.0ライセンスの下で配布されています。\n"
        "再配布の際はライセンス条件に従ってください。\n\n"
        "ソースコード：\n"
        "https://github.com/leantechlibrary-coder/LTL-PDF-Splitter-Tool\n"
    )

    MANUAL_TEXT = (
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "PDF証拠整理ツール 操作説明書\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "■ 目次\n"
        "  1. PDF分割ツールの使い方\n"
        "  2. 証拠番号付与ツールの使い方\n"
        "  3. よくある質問（FAQ）\n\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "1. PDF分割ツールの使い方\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "【基本操作】\n\n"
        "(1) PDFファイルを開く\n"
        "  ・「PDFを開く」ボタンをクリック\n"
        "  ・または、PDFファイルをウィンドウにドラッグ＆ドロップ\n\n"
        "(2) 分割ポイントを選択\n"
        "  ・サムネイル一覧が表示されます\n"
        "  ・分割したい先頭ページをクリック（青い枠が表示されます）\n"
        "  ・複数選択可能です（再クリックで解除）\n\n"
        "(3) 分割実行\n"
        "  ・「分割実行」ボタンをクリック\n"
        "  ・確認ダイアログが表示されるので「Yes」を選択\n"
        "  ・分割完了後、出力フォルダが自動的に開きます\n\n"
        "【サムネイル表示の調整】\n"
        "  ・画面右上のスライダーでサムネイルサイズと列数を調整できます\n"
        "  ・サイズ：100px～500px\n"
        "  ・列数：2列～6列\n\n"
        "【出力について】\n"
        "  ・出力先：元のPDFファイルと同じフォルダ内に\n"
        "    「元ファイル名_分割」フォルダを自動作成\n"
        "  ・同名フォルダがある場合は「_分割_2」「_分割_3」と連番が付きます\n"
        "  ・ファイル名：元ファイル名_001.pdf、元ファイル名_002.pdf...\n\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "2. 証拠番号付与ツールの使い方\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "【基本操作】\n\n"
        "(1) PDFファイルを読み込む\n"
        "  ・「フォルダを開く」：フォルダ内の全PDFファイルを読み込み\n"
        "  ・「ファイルを追加」：個別にファイルを選択して追加\n"
        "  ・ドラッグ＆ドロップ：ファイル/フォルダをウィンドウに直接ドロップ\n\n"
        "(2) ファイルの順番を調整\n"
        "  ・ファイルリストをドラッグ＆ドロップで並び替え\n"
        "  ・または「↑上へ」「↓下へ」ボタンで移動\n\n"
        "(3) 枝番の設定（必要な場合のみ）\n"
        "  ・枝番にしたいファイルを選択\n"
        "  ・「枝番にする」ボタンをクリック\n"
        "  ・例：第2号証の後に枝番を設定すると「甲02の1」「甲02の2」\n"
        "  ・解除する場合は「枝番を解除」ボタン\n\n"
        "(4) 証拠番号の設定\n"
        "  ・証拠種別：甲/乙/その他（カスタム文字列）\n"
        "  ・開始番号：通常は1から\n"
        "  ・証拠番号を印字する：チェックONで1ページ目右上に番号を印字\n"
        "  ・フォントサイズ：8pt～72pt（デフォルト16pt）\n"
        "  ・フォント色：赤/黒/青（デフォルト赤）\n\n"
        "(5) 実行\n"
        "  ・「証拠番号を付与して保存」ボタンをクリック\n"
        "  ・確認ダイアログで内容を確認\n"
        "  ・完了後、出力フォルダが自動的に開きます\n\n"
        "【出力について】\n"
        "  ・出力先：読み込んだファイルの親フォルダ内に\n"
        "    「親フォルダ名_番号付」フォルダを自動作成\n"
        "  ・ファイル名：「甲01.pdf」「甲02.pdf」「甲03の1.pdf」など\n\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "3. よくある質問（FAQ）\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Q. Googleドライブに保存できますか？\n"
        "A. Googleドライブデスクトップの同期フォルダを\n"
        "   出力先に指定することで可能です。\n\n"
        "Q. 何号証まで対応していますか？\n"
        "A. システム上は9999号証まで対応しています。\n\n"
        "Q. 元のファイルが上書きされることはありますか？\n"
        "A. ありません。常に別フォルダに新規ファイルとして出力されます。\n\n"
        "Q. 同じPDFを2回分割するとどうなりますか？\n"
        "A. 出力フォルダに「_分割_2」「_分割_3」と連番が付き、\n"
        "   前回の分割結果は上書きされません。\n\n"
        "Q. PDFにパスワードがかかっている場合は？\n"
        "A. パスワード保護されたPDFには対応していません。\n"
        "   事前にパスワードを解除してから処理してください。\n"
    )

    LICENSE_TEXT = (
        "================================================================================\n"
        "THIRD-PARTY SOFTWARE LICENSES\n"
        "PDF証拠整理ツール\n"
        "================================================================================\n\n"
        "本ソフトウェアは、以下のオープンソースソフトウェアを使用しています。\n"
        "各ソフトウェアのライセンス条項に従い、ライセンス情報を記載します。\n\n\n"
        "================================================================================\n"
        "1. PyMuPDF (fitz)\n"
        "================================================================================\n\n"
        "License: GNU Affero General Public License v3.0 (AGPL-3.0)\n"
        "Copyright: Artifex Software, Inc.\n"
        "Website: https://github.com/pymupdf/PyMuPDF\n\n"
        "ライセンス全文：https://www.gnu.org/licenses/agpl-3.0.txt\n\n\n"
        "================================================================================\n"
        "2. PyQt6\n"
        "================================================================================\n\n"
        "License: GNU General Public License v3.0 (GPL-3.0)\n"
        "Copyright: Riverbank Computing Limited\n"
        "Website: https://www.riverbankcomputing.com/software/pyqt/\n\n"
        "ライセンス全文：https://www.gnu.org/licenses/gpl-3.0.txt\n\n\n"
        "================================================================================\n"
        "3. Python\n"
        "================================================================================\n\n"
        "License: Python Software Foundation License (PSF)\n"
        "Copyright: Python Software Foundation\n"
        "Website: https://www.python.org/\n\n"
        "ライセンス全文：https://docs.python.org/3/license.html\n\n\n"
        "================================================================================\n"
        "本ソフトウェアのライセンス\n"
        "================================================================================\n\n"
        "本ソフトウェア（PDF証拠整理ツール）は、\n"
        "GNU Affero General Public License v3.0 (AGPL-3.0) の下で配布されます。\n"
        "再配布の際はライセンス条件に従ってください。\n\n"
        "ソースコード：\n"
        "https://github.com/leantechlibrary-coder/LTL-PDF-Splitter-Tool\n\n\n"
        "================================================================================\n"
        "免責事項\n"
        "================================================================================\n\n"
        "本ソフトウェアは「現状有姿」(AS IS) で提供され、いかなる保証もありません。\n"
        "本ソフトウェアの使用により生じたいかなる損害についても、開発者は\n"
        "一切の責任を負いません。\n"
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("このソフトについて")
        self.resize(520, 480)
        self.setMinimumSize(400, 350)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(12)

        title_label = QLabel("PDF分割ツール v1.0（無料トライアル版）")
        title_label.setFont(QFont("Yu Gothic UI", 12, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        about_text = QTextEdit()
        about_text.setReadOnly(True)
        about_text.setFont(QFont("Yu Gothic UI", 9))
        about_text.setPlainText(
            "【動作環境】\n"
            "Windows 10 / 11 (64bit)\n\n"
            "【重要】\n"
            "本ソフトウェアは法律専門職の業務効率化を目的としており、\n"
            "専門知識を前提とした設計です。\n\n"
            "【免責事項】\n"
            "本ソフトウェアは「現状有姿」(AS IS) で提供されます。\n"
            "本ソフトウェアの使用により生じたいかなる損害についても、\n"
            "開発者は一切の責任を負いません。\n"
            "重要なファイルは必ずバックアップを取ってからご使用ください。\n\n"
            "【開発・販売】\n"
            "Lean Tech Library\n\n"
            "ご使用前に操作説明書・READMEをご確認ください。"
        )
        layout.addWidget(about_text)

        link_layout = QHBoxLayout()
        link_layout.setSpacing(8)

        manual_btn = QPushButton("操作説明書")
        manual_btn.setToolTip("操作説明書を表示します")
        manual_btn.clicked.connect(self._show_manual)

        readme_btn = QPushButton("README")
        readme_btn.setToolTip("READMEを表示します")
        readme_btn.clicked.connect(self._show_readme)

        license_btn = QPushButton("ライセンス情報")
        license_btn.setToolTip("サードパーティライセンス情報を表示します")
        license_btn.clicked.connect(self._show_licenses)

        link_layout.addWidget(manual_btn)
        link_layout.addWidget(readme_btn)
        link_layout.addWidget(license_btn)
        layout.addLayout(link_layout)

        close_layout = QHBoxLayout()
        close_layout.addStretch()
        close_btn = QPushButton("閉じる")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        close_layout.addWidget(close_btn)
        close_layout.addStretch()
        layout.addLayout(close_layout)

    def _show_manual(self):
        dlg = TextViewerDialog(self, "操作説明書", self.MANUAL_TEXT)
        dlg.exec()

    def _show_readme(self):
        dlg = TextViewerDialog(self, "README", self.README_TEXT)
        dlg.exec()

    def _show_licenses(self):
        dlg = TextViewerDialog(self, "ライセンス情報", self.LICENSE_TEXT)
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
        self.setWindowTitle("PDF分割ツール（無料トライアル版）")
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
            "※枝番を付与する証拠は、枝番ごとに分割してください"
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
            # 出力先フォルダを作成（同名フォルダがある場合は連番を付与）
            parent_dir = Path(self.pdf_path).parent
            base_name = Path(self.pdf_path).stem
            output_dir = parent_dir / f"{base_name}_分割"
            if output_dir.exists():
                counter = 2
                while True:
                    output_dir = parent_dir / f"{base_name}_分割_{counter}"
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

    # --- トライアルチェック ---
    trial = TrialManager()
    remaining = trial.days_remaining()
    expired = trial.is_expired()

    dlg = TrialDialog(remaining, expired)

    if expired:
        dlg.exec()
        sys.exit(0)
    else:
        result = dlg.exec()
        if result == QDialog.DialogCode.Rejected:
            sys.exit(0)

    window = PDFSplitterWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
