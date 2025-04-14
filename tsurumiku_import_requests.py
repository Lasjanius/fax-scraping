import requests
from bs4 import BeautifulSoup
import re
import time
import csv
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                           QFileDialog, QProgressBar, QTextEdit, QMessageBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

class FaxScraperThread(QThread):
    progress_updated = pyqtSignal(int)
    log_updated = pyqtSignal(str)
    finished = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, output_path):
        super().__init__()
        self.output_path = output_path
        self.stop_requested = False

    def run(self):
        try:
            base_url = "https://www.tsurumiku-med.org/Renewal/search/list.html"
            results = []
            
            # ベースページを取得（ユーザーエージェントを設定）
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Content-Type': 'application/x-www-form-urlencoded',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
                'Origin': 'https://www.tsurumiku-med.org',
                'Referer': 'https://www.tsurumiku-med.org/Renewal/search/index.html'
            }
            
            self.log_updated.emit("=== 検索ページを取得 ===")
            search_page_url = "https://www.tsurumiku-med.org/Renewal/search/index.html"
            response = requests.get(search_page_url, headers=headers)
            response.encoding = response.apparent_encoding
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 検索フォームのデータ（キーワード検索を使用）
            data = {
                'mode': '1',  # キーワード検索モード
                'keyword': '',  # 空のキーワードで全件検索
                'week[]': ['1', '2', '3', '4', '5', '6', '7', '8'],  # すべての診療日を選択
                'submit': 'この条件でさがす'
            }
            
            # POSTリクエストを送信
            response = requests.post(base_url, headers=headers, data=data)
            response.encoding = response.apparent_encoding
            
            soup = BeautifulSoup(response.text, 'html.parser')
            links = soup.find_all('a')
            
            total_links = len(links)
            self.log_updated.emit(f"\n合計 {total_links} 個のリンクが見つかりました。処理を開始します...")
            
            # 各リンク先にアクセスしてFAX番号を取得
            for i, link in enumerate(links):
                if self.stop_requested:
                    self.log_updated.emit("処理を中止しました")
                    break
                
                href = link.get('href')
                if href and not href.startswith('#') and not href.startswith('javascript'):
                    # 相対URLを絶対URLに変換
                    if not href.startswith('http'):
                        if base_url.endswith('/'):
                            full_url = base_url + href
                        else:
                            full_url = base_url[:base_url.rfind('/')+1] + href
                    else:
                        full_url = href
                    
                    link_text = link.get_text().strip()
                    if not link_text:
                        link_text = f"リンク {i+1}"
                    
                    self.log_updated.emit(f"処理中 ({i+1}/{total_links}): {link_text} - {full_url}")
                    
                    # FAX番号を取得
                    fax_number = self.get_fax_number(full_url)
                    
                    # 結果を保存
                    results.append({
                        'リンク名': link_text,
                        'URL': full_url,
                        'FAX番号': fax_number
                    })
                    
                    # 進捗を更新
                    progress = int((i + 1) / total_links * 100)
                    self.progress_updated.emit(progress)
                    
                    # サーバーに負荷をかけないよう少し待機
                    time.sleep(1)
            
            # 結果をCSVファイルに保存
            with open(self.output_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = ['リンク名', 'URL', 'FAX番号']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                for result in results:
                    writer.writerow(result)
            
            self.log_updated.emit(f"\n処理が完了しました。結果は '{self.output_path}' に保存されています。")
            self.finished.emit()
            
        except Exception as e:
            self.error_occurred.emit(str(e))

    def get_fax_number(self, url):
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            response.encoding = response.apparent_encoding
            
            soup = BeautifulSoup(response.text, 'html.parser')
            page_text = soup.get_text()
            
            fax_pattern = re.compile(r'[Ff][Aa][Xx]:?\s*(\d[\d\-]+)')
            match = fax_pattern.search(page_text)
            
            if match:
                fax_number = match.group(1)
                return fax_number
            else:
                return "FAX番号が見つかりませんでした"
                
        except Exception as e:
            return f"エラー: {str(e)}"

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("鶴見区医師会 FAX番号収集ツール")
        self.setGeometry(100, 100, 800, 600)
        
        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # レイアウト
        layout = QVBoxLayout()
        main_widget.setLayout(layout)
        
        # 出力ファイル選択
        output_frame = QHBoxLayout()
        layout.addLayout(output_frame)
        
        output_label = QLabel("保存先:")
        output_frame.addWidget(output_label)
        
        self.output_path = QLineEdit()
        output_frame.addWidget(self.output_path)
        
        browse_button = QPushButton("参照...")
        browse_button.clicked.connect(self.browse_output_file)
        output_frame.addWidget(browse_button)
        
        # 進捗バー
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # ログ表示
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        
        # ボタン
        button_frame = QHBoxLayout()
        layout.addLayout(button_frame)
        
        self.start_button = QPushButton("開始")
        self.start_button.clicked.connect(self.start_scraping)
        button_frame.addWidget(self.start_button)
        
        self.stop_button = QPushButton("中止")
        self.stop_button.clicked.connect(self.stop_scraping)
        self.stop_button.setEnabled(False)
        button_frame.addWidget(self.stop_button)
        
        # スレッド
        self.scraper_thread = None

    def browse_output_file(self):
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "出力ファイルの保存先を選択",
            "",
            "CSVファイル (*.csv);;すべてのファイル (*.*)"
        )
        if filename:
            self.output_path.setText(filename)

    def start_scraping(self):
        if not self.output_path.text():
            QMessageBox.warning(self, "警告", "保存先を選択してください。")
            return
        
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.progress_bar.setValue(0)
        self.log_text.clear()
        
        self.scraper_thread = FaxScraperThread(self.output_path.text())
        self.scraper_thread.progress_updated.connect(self.update_progress)
        self.scraper_thread.log_updated.connect(self.update_log)
        self.scraper_thread.finished.connect(self.scraping_finished)
        self.scraper_thread.error_occurred.connect(self.handle_error)
        self.scraper_thread.start()

    def stop_scraping(self):
        if self.scraper_thread:
            self.scraper_thread.stop_requested = True
            self.stop_button.setEnabled(False)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_log(self, message):
        self.log_text.append(message)

    def scraping_finished(self):
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        QMessageBox.information(self, "完了", "処理が完了しました。")

    def handle_error(self, error_message):
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        QMessageBox.critical(self, "エラー", f"エラーが発生しました:\n{error_message}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())