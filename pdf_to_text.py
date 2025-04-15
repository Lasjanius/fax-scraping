#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import pandas as pd
import fitz  # PyMuPDF
import re
from tqdm import tqdm
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QProgressBar, QTextEdit, QFileDialog,
    QLineEdit, QFrame, QGroupBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt

def extract_fax_numbers(pdf_path, output_path=None):
    """PDFからFAX番号を抽出してCSVファイルに保存"""
    print(f"PDFファイルからFAX番号を抽出中: {pdf_path}")
    
    # 出力パスが指定されていない場合は、PDFと同じ場所に作成
    if output_path is None:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_dir = os.path.dirname(os.path.abspath(pdf_path))
        output_path = os.path.join(output_dir, f"{base_name}_fax_numbers.csv")
    
    try:
        # PDFを開く
        doc = fitz.open(pdf_path)
        num_pages = len(doc)
        print(f"PDFを開きました: {num_pages}ページ")
        
        # FAX番号を抽出
        fax_numbers = []
        
        # FAX番号を抽出するための正規表現パターン
        # 括弧()に挟まれて、ハイフン2つを含む10桁の数字
        fax_pattern = r'\((\d{3}-\d{3}-\d{4})\)'
        
        for page_num in tqdm(range(num_pages), desc="FAX番号抽出中"):
            try:
                page = doc[page_num]
                text = page.get_text()
                
                # 正規表現でFAX番号を検索
                matches = re.findall(fax_pattern, text)
                
                if matches:
                    # 重複を除去
                    unique_fax = list(set(matches))
                    
                    for fax in unique_fax:
                        fax_numbers.append({
                            "page": page_num + 1,
                            "fax_number": fax,
                            "context": get_context(text, fax)  # FAX番号の周辺テキストを取得
                        })
                        
            except Exception as e:
                print(f"ページ {page_num+1} の処理中にエラー: {e}")
        
        # データフレームに変換
        df = pd.DataFrame(fax_numbers)
        
        if len(df) > 0:
            # CSVに保存
            df.to_csv(output_path, index=False, encoding="utf-8-sig")
            print(f"FAX番号を保存しました: {output_path}")
            print(f"合計 {len(df)} 件のFAX番号が見つかりました")
        else:
            print("FAX番号が見つかりませんでした")
        
        # PDFを閉じる
        doc.close()
        
        return output_path, len(df)
    
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return None, 0

def get_context(text, fax_number):
    """FAX番号の前後のテキストを抽出して、どの施設のFAX番号かを特定しやすくする"""
    # FAX番号を含む行を探す
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if fax_number in line:
            # FAX番号を含む行とその前後の行（あれば）を返す
            start = max(0, i-1)
            end = min(len(lines), i+2)
            return '\n'.join(lines[start:end])
    
    # 見つからない場合はFAX番号だけを返す
    return f"({fax_number})"

class ExtractWorker(QThread):
    progress_updated = pyqtSignal(int, int)  # current_page, total_pages
    log_updated = pyqtSignal(str)
    finished = pyqtSignal(str, int)  # output_path, num_found
    error_occurred = pyqtSignal(str)

    def __init__(self, pdf_path, output_path=None):
        super().__init__()
        self.pdf_path = pdf_path
        self.output_path = output_path

    def run(self):
        try:
            self.log_updated.emit(f"PDFファイルからFAX番号を抽出中: {self.pdf_path}")
            
            # 抽出処理（既存のコードを使用）
            result_path, num_found = extract_fax_numbers(self.pdf_path, self.output_path)
            
            if result_path:
                self.finished.emit(result_path, num_found)
            else:
                self.error_occurred.emit("抽出処理に失敗しました")
        except Exception as e:
            self.error_occurred.emit(f"エラーが発生しました: {str(e)}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF FAX番号抽出ツール")
        self.setGeometry(100, 100, 800, 600)
        
        # メインウィジェットとレイアウトの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        main_widget.setLayout(layout)
        
        # 入力ファイル選択部分
        input_group = QGroupBox("入力PDFファイル")
        input_layout = QHBoxLayout()
        input_group.setLayout(input_layout)
        
        self.input_path = QLineEdit()
        self.input_path.setReadOnly(True)
        input_layout.addWidget(self.input_path)
        
        browse_input_button = QPushButton("参照...")
        browse_input_button.clicked.connect(self.browse_input_file)
        input_layout.addWidget(browse_input_button)
        
        layout.addWidget(input_group)
        
        # 出力ファイル選択部分
        output_group = QGroupBox("出力CSVファイル")
        output_layout = QHBoxLayout()
        output_group.setLayout(output_layout)
        
        self.output_path = QLineEdit()
        self.output_path.setReadOnly(True)
        output_layout.addWidget(self.output_path)
        
        browse_output_button = QPushButton("参照...")
        browse_output_button.clicked.connect(self.browse_output_file)
        output_layout.addWidget(browse_output_button)
        
        layout.addWidget(output_group)
        
        # 進捗バー
        progress_group = QGroupBox("進捗状況")
        progress_layout = QVBoxLayout()
        progress_group.setLayout(progress_layout)
        
        self.progress_bar = QProgressBar()
        progress_layout.addWidget(self.progress_bar)
        
        layout.addWidget(progress_group)
        
        # ログ表示部分
        log_group = QGroupBox("ログ")
        log_layout = QVBoxLayout()
        log_group.setLayout(log_layout)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        layout.addWidget(log_group)
        
        # ボタン部分
        button_frame = QFrame()
        button_layout = QHBoxLayout()
        button_frame.setLayout(button_layout)
        
        self.extract_button = QPushButton("抽出開始")
        self.extract_button.clicked.connect(self.start_extraction)
        button_layout.addWidget(self.extract_button)
        
        button_layout.addStretch()
        
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.close)
        button_layout.addWidget(close_button)
        
        layout.addWidget(button_frame)
        
        # 抽出ワーカー
        self.worker = None

    def browse_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "PDFファイルを選択",
            "",
            "PDFファイル (*.pdf);;すべてのファイル (*.*)"
        )
        if file_path:
            self.input_path.setText(file_path)
            self.log(f"入力ファイルを選択しました: {file_path}")
            
            # デフォルトの出力パスを設定
            if not self.output_path.text():
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_dir = os.path.dirname(os.path.abspath(file_path))
                default_output = os.path.join(output_dir, f"{base_name}_fax_numbers.csv")
                self.output_path.setText(default_output)
                self.log(f"デフォルトの出力先を設定しました: {default_output}")

    def browse_output_file(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存先を選択",
            self.output_path.text() if self.output_path.text() else "",
            "CSVファイル (*.csv);;すべてのファイル (*.*)"
        )
        if file_path:
            # 拡張子がない場合は.csvを追加
            if not os.path.splitext(file_path)[1]:
                file_path += '.csv'
            self.output_path.setText(file_path)
            self.log(f"出力先を選択しました: {file_path}")

    def start_extraction(self):
        if not self.input_path.text():
            self.log("エラー: PDFファイルを選択してください")
            return
            
        if not self.output_path.text():
            self.log("エラー: 出力先を選択してください")
            return
            
        if self.worker and self.worker.isRunning():
            self.log("すでに処理中です")
            return
            
        self.worker = ExtractWorker(self.input_path.text(), self.output_path.text())
        self.worker.log_updated.connect(self.log)
        self.worker.finished.connect(self.extraction_finished)
        self.worker.error_occurred.connect(self.handle_error)
        
        self.extract_button.setEnabled(False)
        self.log("抽出処理を開始します...")
        
        self.worker.start()

    def extraction_finished(self, output_path, num_found):
        self.extract_button.setEnabled(True)
        self.log(f"抽出が完了しました")
        self.log(f"合計 {num_found} 件のFAX番号が見つかりました")
        self.log(f"結果は {output_path} に保存されました")

    def handle_error(self, error_message):
        self.extract_button.setEnabled(True)
        self.log(f"エラー: {error_message}")

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

if __name__ == "__main__":
    # macOS向けの設定
    if sys.platform == 'darwin':
        os.environ['QT_MAC_WANTS_LAYER'] = '1'
        os.environ['QT_QPA_PLATFORM'] = 'cocoa'
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
