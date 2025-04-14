import sys
import os

# macOS向けの設定をインポート前に行う
os.environ['QT_MAC_WANTS_LAYER'] = '1'
os.environ['QT_QPA_PLATFORM'] = 'cocoa'  # macOS特有の設定

import csv
import time
import json
import re
import threading
import pandas as pd
import requests
from bs4 import BeautifulSoup
from googlesearch import search
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QProgressBar, QTextEdit, QFileDialog,
    QLineEdit, QFrame, QGroupBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

class ScrapingWorker(QThread):
    progress_updated = pyqtSignal(str, int, int)  # clinic_name, current, total
    log_updated = pyqtSignal(str)
    finished = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, csv_path):
        super().__init__()
        self.csv_path = csv_path
        self.stop_requested = False
        self.retry_delay = 5  # 基本待機時間（秒）
        self.max_retries = 3  # 最大リトライ回数

    def search_with_retry(self, query, retry_count=0):
        """Google検索を実行し、ネットワークエラーのみリトライする"""
        try:
            # 検索実行前に待機（リトライ回数に応じて待機時間を増加）
            wait_time = self.retry_delay * (2 ** retry_count)  # エクスポネンシャルバックオフ
            self.log_updated.emit(f"- 検索前に {wait_time}秒待機します...")
            time.sleep(wait_time)
            
            # Google検索を実行
            search_results = list(search(query, num=1))
            
            if search_results:
                return search_results
            else:
                raise Exception("検索結果が0件でした")
                
        except requests.exceptions.HTTPError as e:
            # 429エラー（Too Many Requests）の場合は長めに待機
            if "429" in str(e):
                if retry_count < self.max_retries:
                    wait_time = self.retry_delay * (2 ** (retry_count + 2))  # 通常より長い待機時間
                    self.log_updated.emit(f"- リクエスト制限エラー(429): {str(e)}")
                    self.log_updated.emit(f"- より長く待機します... {wait_time}秒 ({retry_count + 1}/{self.max_retries})")
                    time.sleep(wait_time)  # 追加で待機
                    return self.search_with_retry(query, retry_count + 1)
                else:
                    raise Exception(f"リクエスト制限エラーが続いています: {str(e)}")
            else:
                # その他のHTTPエラー
                raise Exception(f"HTTPエラー: {str(e)}")
                
        except requests.exceptions.RequestException as e:
            # ネットワークエラーのみリトライ
            if retry_count < self.max_retries:
                wait_time = self.retry_delay * (2 ** retry_count)
                self.log_updated.emit(f"- ネットワークエラー: {str(e)}")
                self.log_updated.emit(f"- リトライします... {wait_time}秒後 ({retry_count + 1}/{self.max_retries})")
                time.sleep(wait_time)  # 追加で待機
                return self.search_with_retry(query, retry_count + 1)
            else:
                raise Exception(f"ネットワークエラーが続いています: {str(e)}")
                
        except Exception as e:
            # その他のエラーはリトライせずに例外を投げる
            if "Too Many Requests" in str(e) or "429" in str(e):
                if retry_count < self.max_retries:
                    wait_time = self.retry_delay * (2 ** (retry_count + 2))
                    self.log_updated.emit(f"- リクエスト制限エラー: {str(e)}")
                    self.log_updated.emit(f"- より長く待機します... {wait_time}秒 ({retry_count + 1}/{self.max_retries})")
                    time.sleep(wait_time)
                    return self.search_with_retry(query, retry_count + 1)
                else:
                    raise Exception(f"リクエスト制限エラーが続いています: {str(e)}")
            else:
                raise Exception(f"検索に失敗しました: {str(e)}")

    def run(self):
        try:
            self.log_updated.emit("処理を開始します...")
            
            # CSVファイルを読み込む
            try:
                df = pd.read_csv(self.csv_path)
                self.log_updated.emit(f"CSVファイルを読み込みました: {len(df)}件のデータ")
            except Exception as e:
                self.log_updated.emit(f"CSVファイルの読み込みに失敗しました: {str(e)}")
                self.error_occurred.emit(f"CSVファイルの読み込みに失敗しました: {str(e)}")
                return

            total = len(df)
            self.log_updated.emit(f"総処理件数: {total}件")

            # FAX番号カラムがなければ追加
            if 'FAX番号' not in df.columns:
                df['FAX番号'] = df['FAX番号'].astype('object')  # 文字列として扱う
                self.log_updated.emit("FAX番号カラムを追加しました")

            # エラー詳細カラムがなければ追加
            if 'エラー詳細' not in df.columns:
                df['エラー詳細'] = None
                self.log_updated.emit("エラー詳細カラムを追加しました")

            # 各クリニックに対して処理
            for index, row in df.iterrows():
                if self.stop_requested:
                    self.log_updated.emit("処理を中断しました")
                    break

                try:
                    clinic_name = row.iloc[0]  # インデックス列はCSVの最初の列と仮定
                    self.progress_updated.emit(clinic_name, index + 1, total)
                    self.log_updated.emit(f"処理中: {clinic_name} ({index + 1}/{total})")

                    # エラー詳細をリセット
                    df.at[index, 'エラー詳細'] = None

                    # すでにFAX番号がある場合はスキップ
                    if not pd.isna(df.at[index, 'FAX番号']):
                        self.log_updated.emit(f"- すでにFAX番号があります: {df.at[index, 'FAX番号']}")
                        continue

                    try:
                        # Google検索でクリニックのウェブサイトを探す
                        search_query = clinic_name  # クリニック名のみで検索
                        self.log_updated.emit(f"- 検索クエリ: {search_query}")
                        search_results = self.search_with_retry(search_query)
                        
                        if search_results:
                            # 検索結果から最適なURLを選択
                            matching_url = None
                            for url in search_results:
                                try:
                                    # ページを取得
                                    response = requests.get(url, timeout=10)
                                    soup = BeautifulSoup(response.text, 'html.parser')
                                    
                                    # タイトルを取得
                                    title = soup.title.string if soup.title else ""
                                    
                                    # タイトルにクリニック名が含まれているかチェック
                                    if clinic_name in title:
                                        matching_url = url
                                        self.log_updated.emit(f"- タイトルに一致するURLを発見: {url}")
                                        break
                                        
                                except Exception as e:
                                    self.log_updated.emit(f"- URL取得エラー: {str(e)}")
                                    continue
                            
                            if matching_url:
                                url = matching_url
                            else:
                                # マッチするURLが見つからない場合は最初の結果を使用
                                url = search_results[0]
                                self.log_updated.emit(f"- タイトルに一致するURLが見つかりませんでした。最初の結果を使用: {url}")
                            
                            # ウェブページを取得（タイムアウトを設定）
                            try:
                                response = requests.get(url, timeout=10)
                                response.raise_for_status()  # ステータスコードチェック
                                soup = BeautifulSoup(response.text, 'html.parser')
                                self.log_updated.emit("- ページの取得に成功しました")

                                # 詳細ページへのリンクを探す
                                detail_links = []
                                
                                # パターン1: クリニック名のリンク
                                clinic_links = soup.find_all('a', href=re.compile(r'detail\.html\?id=\d+'))
                                for link in clinic_links:
                                    link_text = link.text.strip()
                                    # クリニック名の部分一致をチェック
                                    if any(name in link_text for name in [clinic_name, clinic_name.replace('クリニック', ''), clinic_name.replace('医院', '')]):
                                        detail_links.append(link['href'])
                                        self.log_updated.emit(f"- リンクを検出: {link_text}")
                                
                                # パターン2: テーブル内のリンク
                                if not detail_links:
                                    tables = soup.find_all('table')
                                    for table in tables:
                                        rows = table.find_all('tr')
                                        for row in rows:
                                            cells = row.find_all(['td', 'th'])
                                            for cell in cells:
                                                links = cell.find_all('a', href=re.compile(r'detail\.html\?id=\d+'))
                                                for link in links:
                                                    link_text = link.text.strip()
                                                    # クリニック名の部分一致をチェック
                                                    if any(name in link_text for name in [clinic_name, clinic_name.replace('クリニック', ''), clinic_name.replace('医院', '')]):
                                                        detail_links.append(link['href'])
                                                        self.log_updated.emit(f"- テーブル内でリンクを検出: {link_text}")

                                if detail_links:
                                    # 詳細ページのURLを構築
                                    base_url = '/'.join(url.split('/')[:-1]) + '/'
                                    detail_url = base_url + detail_links[0]
                                    self.log_updated.emit(f"- 詳細ページを検出: {detail_url}")

                                    # 詳細ページを取得
                                    try:
                                        detail_response = requests.get(detail_url, timeout=10)
                                        detail_response.raise_for_status()
                                        detail_soup = BeautifulSoup(detail_response.text, 'html.parser')
                                        self.log_updated.emit("- 詳細ページの取得に成功しました")

                                        # FAX番号を探す
                                        fax_number = None
                                        
                                        # パターン1: FAXという文字の後ろの数字
                                        fax_patterns = detail_soup.find_all(string=re.compile(r'FAX.*?(\d[\d\-]+\d)'))
                                        if fax_patterns:
                                            match = re.search(r'FAX.*?(\d[\d\-]+\d)', fax_patterns[0])
                                            if match:
                                                fax_number = match.group(1)
                                                self.log_updated.emit("- パターン1でFAX番号を検出")
                                        
                                        # パターン2: class名やid名にfaxを含む要素
                                        if not fax_number:
                                            fax_elements = detail_soup.find_all(class_=re.compile('fax', re.I))
                                            fax_elements.extend(detail_soup.find_all(id=re.compile('fax', re.I)))
                                            for elem in fax_elements:
                                                match = re.search(r'(\d[\d\-]+\d)', elem.text)
                                                if match:
                                                    fax_number = match.group(1)
                                                    self.log_updated.emit("- パターン2でFAX番号を検出")
                                                    break
                                        
                                        # パターン3: dt/dd タグの組み合わせ
                                        if not fax_number:
                                            dt_elements = detail_soup.find_all('dt')
                                            for dt in dt_elements:
                                                dt_text = dt.text.strip().upper()
                                                if 'FAX' in dt_text and not any(x in dt_text for x in ['TEL', 'PHONE', '電話']):
                                                    next_dd = dt.find_next('dd')
                                                    if next_dd:
                                                        match = re.search(r'(\d[\d\-]+\d)', next_dd.text)
                                                        if match:
                                                            fax_number = match.group(1)
                                                            self.log_updated.emit("- パターン3でFAX番号を検出")
                                                            break
                                        
                                        # パターン4: テーブル内のFAX番号
                                        if not fax_number:
                                            tables = detail_soup.find_all('table')
                                            for table in tables:
                                                rows = table.find_all('tr')
                                                for row in rows:
                                                    cells = row.find_all(['td', 'th'])
                                                    for i, cell in enumerate(cells):
                                                        cell_text = cell.text.strip().upper()
                                                        if ('FAX' in cell_text and not any(x in cell_text for x in ['TEL', 'PHONE', '電話']) 
                                                            and i + 1 < len(cells)):
                                                            match = re.search(r'(\d[\d\-]+\d)', cells[i + 1].text)
                                                            if match:
                                                                fax_number = match.group(1)
                                                                self.log_updated.emit("- パターン4でFAX番号を検出")
                                                                break
                                                    if fax_number:
                                                        break
                                                if fax_number:
                                                    break
                                        
                                        # パターン5: 一般的な電話番号パターン（FAXの前後）
                                        if not fax_number:
                                            text_blocks = detail_soup.find_all(string=re.compile(r'FAX|fax'))
                                            for text in text_blocks:
                                                # FAXの前後100文字を検索
                                                context = text.parent.text
                                                fax_index = context.upper().find('FAX')
                                                if fax_index != -1:
                                                    # FAXの前後を検索
                                                    before = context[max(0, fax_index-100):fax_index]
                                                    after = context[fax_index:min(len(context), fax_index+100)]
                                                    # FAXの直後を優先的に検索
                                                    match = re.search(r'FAX[^\d]*(\d[\d\-]+\d)', after)
                                                    if not match:
                                                        match = re.search(r'(\d[\d\-]+\d)[^\d]*FAX', before)
                                                    if match:
                                                        fax_number = match.group(1)
                                                        self.log_updated.emit("- パターン5でFAX番号を検出")
                                                        break
                                        
                                        # パターン6: 「お問い合わせ」や「連絡先」セクション内のFAX番号
                                        if not fax_number:
                                            contact_sections = detail_soup.find_all(['div', 'section'], class_=re.compile(r'contact|inquiry|access', re.I))
                                            contact_sections.extend(detail_soup.find_all(['div', 'section'], id=re.compile(r'contact|inquiry|access', re.I)))
                                            for section in contact_sections:
                                                text = section.text
                                                match = re.search(r'FAX[^\d]*(\d[\d\-]+\d)', text)
                                                if match:
                                                    fax_number = match.group(1)
                                                    self.log_updated.emit("- パターン6でFAX番号を検出")
                                                    break
                                        
                                        if fax_number:
                                            # 番号の正規化（ハイフン統一のみ）
                                            fax_number = re.sub(r'[^\d\-]', '', fax_number)
                                            df.at[index, 'FAX番号'] = str(fax_number)  # 文字列として保存
                                            df.at[index, 'エラー詳細'] = None  # エラーをクリア
                                            self.log_updated.emit(f"- メインページでFAX番号が見つかりました: {fax_number}")
                                        else:
                                            df.at[index, 'エラー詳細'] = "メインページでもFAX番号が見つかりませんでした"
                                            self.log_updated.emit("- メインページでもFAX番号が見つかりませんでした")
                                    
                                    except requests.exceptions.RequestException as e:
                                        self.log_updated.emit(f"- 詳細ページの取得に失敗しました: {str(e)}")
                                        df.at[index, 'エラー詳細'] = f"詳細ページの取得に失敗: {str(e)}"
                                else:
                                    df.at[index, 'エラー詳細'] = "詳細ページへのリンクが見つかりませんでした"
                                    self.log_updated.emit("- 詳細ページへのリンクが見つかりませんでした")
                                    # メインページから直接FAX番号を探す
                                    self.log_updated.emit("- メインページから直接FAX番号を探します")
                                    fax_number = None
                                    
                                    # パターン1: FAXという文字の後ろの数字
                                    fax_patterns = soup.find_all(string=re.compile(r'FAX.*?(\d[\d\-]+\d)'))
                                    if fax_patterns:
                                        match = re.search(r'FAX.*?(\d[\d\-]+\d)', fax_patterns[0])
                                        if match:
                                            fax_number = match.group(1)
                                            self.log_updated.emit("- パターン1でFAX番号を検出")
                                    
                                    # パターン2: class名やid名にfaxを含む要素
                                    if not fax_number:
                                        fax_elements = soup.find_all(class_=re.compile('fax', re.I))
                                        fax_elements.extend(soup.find_all(id=re.compile('fax', re.I)))
                                        for elem in fax_elements:
                                            match = re.search(r'(\d[\d\-]+\d)', elem.text)
                                            if match:
                                                fax_number = match.group(1)
                                                self.log_updated.emit("- パターン2でFAX番号を検出")
                                                break
                                    
                                    # パターン3: dt/dd タグの組み合わせ
                                    if not fax_number:
                                        dt_elements = soup.find_all('dt')
                                        for dt in dt_elements:
                                            dt_text = dt.text.strip().upper()
                                            if 'FAX' in dt_text and not any(x in dt_text for x in ['TEL', 'PHONE', '電話']):
                                                next_dd = dt.find_next('dd')
                                                if next_dd:
                                                    match = re.search(r'(\d[\d\-]+\d)', next_dd.text)
                                                    if match:
                                                        fax_number = match.group(1)
                                                        self.log_updated.emit("- パターン3でFAX番号を検出")
                                                        break
                                    
                                    # パターン4: テーブル内のFAX番号
                                    if not fax_number:
                                        tables = soup.find_all('table')
                                        for table in tables:
                                            rows = table.find_all('tr')
                                            for row in rows:
                                                cells = row.find_all(['td', 'th'])
                                                for i, cell in enumerate(cells):
                                                    cell_text = cell.text.strip().upper()
                                                    if ('FAX' in cell_text and not any(x in cell_text for x in ['TEL', 'PHONE', '電話']) 
                                                        and i + 1 < len(cells)):
                                                        match = re.search(r'(\d[\d\-]+\d)', cells[i + 1].text)
                                                        if match:
                                                            fax_number = match.group(1)
                                                            self.log_updated.emit("- パターン4でFAX番号を検出")
                                                            break
                                                    if fax_number:
                                                        break
                                            if fax_number:
                                                break
                                    
                                    # パターン5: 一般的な電話番号パターン（FAXの前後）
                                    if not fax_number:
                                        text_blocks = soup.find_all(string=re.compile(r'FAX|fax'))
                                        for text in text_blocks:
                                            # FAXの前後100文字を検索
                                            context = text.parent.text
                                            fax_index = context.upper().find('FAX')
                                            if fax_index != -1:
                                                # FAXの前後を検索
                                                before = context[max(0, fax_index-100):fax_index]
                                                after = context[fax_index:min(len(context), fax_index+100)]
                                                # FAXの直後を優先的に検索
                                                match = re.search(r'FAX[^\d]*(\d[\d\-]+\d)', after)
                                                if not match:
                                                    match = re.search(r'(\d[\d\-]+\d)[^\d]*FAX', before)
                                                if match:
                                                    fax_number = match.group(1)
                                                    self.log_updated.emit("- パターン5でFAX番号を検出")
                                                    break
                                    
                                    # パターン6: 「お問い合わせ」や「連絡先」セクション内のFAX番号
                                    if not fax_number:
                                        contact_sections = soup.find_all(['div', 'section'], class_=re.compile(r'contact|inquiry|access', re.I))
                                        contact_sections.extend(soup.find_all(['div', 'section'], id=re.compile(r'contact|inquiry|access', re.I)))
                                        for section in contact_sections:
                                            text = section.text
                                            match = re.search(r'FAX[^\d]*(\d[\d\-]+\d)', text)
                                            if match:
                                                fax_number = match.group(1)
                                                self.log_updated.emit("- パターン6でFAX番号を検出")
                                                break
                                    
                                    if fax_number:
                                        # 番号の正規化（ハイフン統一のみ）
                                        fax_number = re.sub(r'[^\d\-]', '', fax_number)
                                        df.at[index, 'FAX番号'] = str(fax_number)  # 文字列として保存
                                        df.at[index, 'エラー詳細'] = None  # エラーをクリア
                                        self.log_updated.emit(f"- メインページでFAX番号が見つかりました: {fax_number}")
                                    else:
                                        df.at[index, 'エラー詳細'] = "メインページでもFAX番号が見つかりませんでした"
                                        self.log_updated.emit("- メインページでもFAX番号が見つかりませんでした")
                            
                            except requests.exceptions.RequestException as e:
                                self.log_updated.emit(f"- ページの取得に失敗しました: {str(e)}")
                                df.at[index, 'エラー詳細'] = f"ページの取得に失敗: {str(e)}"
                                continue
                        else:
                            df.at[index, 'エラー詳細'] = "ウェブサイトが見つかりませんでした"
                            self.log_updated.emit("- ウェブサイトが見つかりませんでした")
                    
                    except Exception as e:
                        error_msg = f"処理中にエラーが発生しました: {str(e)}"
                        self.log_updated.emit(f"- {error_msg}")
                        df.at[index, 'エラー詳細'] = error_msg

                    # 定期的に保存
                    if (index + 1) % 10 == 0:
                        try:
                            df.to_csv(self.csv_path, index=False)
                            self.log_updated.emit(f"- {index + 1}件目を保存しました")
                        except Exception as e:
                            self.log_updated.emit(f"- 保存に失敗しました: {str(e)}")

                    # サーバーに負荷をかけないよう少し待機
                    time.sleep(0.5)  # 0.5秒待機

                except Exception as e:
                    self.log_updated.emit(f"行の処理中にエラーが発生しました: {str(e)}")
                    continue

            # 最終結果を保存
            try:
                df.to_csv(self.csv_path, index=False)
                self.log_updated.emit("処理が完了しました")
            except Exception as e:
                self.log_updated.emit(f"最終保存に失敗しました: {str(e)}")
                self.error_occurred.emit(f"最終保存に失敗しました: {str(e)}")
            
        except Exception as e:
            error_msg = f"予期せぬエラーが発生しました: {str(e)}"
            self.log_updated.emit(error_msg)
            self.error_occurred.emit(error_msg)
        
        finally:
            self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("クリニックFAX番号収集ツール")
        self.setGeometry(100, 100, 800, 600)
        
        # メインウィジェットとレイアウトの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        main_widget.setLayout(layout)
        
        # 入力ファイル選択部分
        file_group = QGroupBox("入力ファイル")
        file_layout = QHBoxLayout()
        file_group.setLayout(file_layout)
        
        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        file_layout.addWidget(self.file_path)
        
        browse_button = QPushButton("参照...")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_button)
        
        layout.addWidget(file_group)
        
        # 処理状況表示部分
        status_group = QGroupBox("処理状況")
        status_layout = QVBoxLayout()
        status_group.setLayout(status_layout)
        
        # ステータス
        status_frame = QFrame()
        status_frame_layout = QHBoxLayout()
        status_frame.setLayout(status_frame_layout)
        
        status_frame_layout.addWidget(QLabel("ステータス:"))
        self.status_label = QLabel("ファイルを選択してください")
        status_frame_layout.addWidget(self.status_label)
        status_frame_layout.addStretch()
        
        status_layout.addWidget(status_frame)
        
        # 現在の処理
        current_frame = QFrame()
        current_frame_layout = QHBoxLayout()
        current_frame.setLayout(current_frame_layout)
        
        current_frame_layout.addWidget(QLabel("処理中のクリニック:"))
        self.current_clinic_label = QLabel("")
        current_frame_layout.addWidget(self.current_clinic_label)
        current_frame_layout.addStretch()
        
        status_layout.addWidget(current_frame)
        
        # プログレスバー
        self.progress_bar = QProgressBar()
        status_layout.addWidget(self.progress_bar)
        
        # ログ表示
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        status_layout.addWidget(self.log_text)
        
        layout.addWidget(status_group)
        
        # ボタン
        button_frame = QFrame()
        button_layout = QHBoxLayout()
        button_frame.setLayout(button_layout)
        
        self.start_button = QPushButton("実行")
        self.start_button.clicked.connect(self.start_scraping)
        button_layout.addWidget(self.start_button)
        
        self.stop_button = QPushButton("中止")
        self.stop_button.clicked.connect(self.stop_scraping)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.stop_button)
        
        button_layout.addStretch()
        
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.close)
        button_layout.addWidget(close_button)
        
        layout.addWidget(button_frame)
        
        # スクレイピングワーカー
        self.worker = None

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "CSVファイルを選択",
            "",
            "CSVファイル (*.csv);;すべてのファイル (*.*)"
        )
        if file_path:
            self.file_path.setText(file_path)
            self.log("ファイルを選択しました: " + file_path)
            self.analyze_csv(file_path)

    def analyze_csv(self, file_path):
        try:
            df = pd.read_csv(file_path)
            total = len(df)
            self.status_label.setText(f"読み込み完了: {total}件のクリニックが見つかりました")
            self.log(f"クリニック総数: {total}件")
            
            if 'FAX番号' in df.columns:
                self.log("既存のFAX番号カラムが見つかりました")
            else:
                self.log("FAX番号カラムが見つかりません。新規に作成します")
            
            if 'エラー詳細' in df.columns:
                self.log("既存のエラー詳細カラムが見つかりました")
            else:
                self.log("エラー詳細カラムが見つかりません。新規に作成します")
                
        except Exception as e:
            self.status_label.setText("エラー: CSVファイルの読み込みに失敗しました")
            self.log(f"エラー: {str(e)}")

    def start_scraping(self):
        if not self.file_path.text():
            self.log("エラー: CSVファイルを選択してください")
            return
            
        if self.worker and self.worker.isRunning():
            self.log("すでに処理中です")
            return
            
        self.worker = ScrapingWorker(self.file_path.text())
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.log_updated.connect(self.log)
        self.worker.finished.connect(self.scraping_finished)
        self.worker.error_occurred.connect(self.handle_error)
        
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.status_label.setText("処理中...")
        self.progress_bar.setValue(0)
        
        self.worker.start()

    def stop_scraping(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop_requested = True
            self.status_label.setText("処理を中止しています...")
            self.log("処理中止リクエストを受け付けました")
            self.stop_button.setEnabled(False)

    def update_progress(self, clinic_name, current, total):
        self.current_clinic_label.setText(clinic_name)
        progress = int((current / total) * 100)
        self.progress_bar.setValue(progress)
        self.status_label.setText(f"処理中... {current}/{total} ({progress}%)")

    def scraping_finished(self):
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        if self.worker and self.worker.stop_requested:
            self.status_label.setText("処理が中止されました")
        else:
            self.status_label.setText("処理が完了しました")
        
        # ワーカースレッドの後処理
        if self.worker:
            self.worker.quit()
            self.worker.wait()  # スレッドが完全に終了するまで待機
            self.worker.deleteLater()  # メモリリソースを解放

    def handle_error(self, error_message):
        self.log(f"エラーが発生しました: {error_message}")
        self.status_label.setText("エラーが発生しました")

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def closeEvent(self, event):
        # アプリケーション終了時の処理
        if self.worker and self.worker.isRunning():
            self.worker.stop_requested = True
            self.worker.quit()
            self.worker.wait(1000)  # 最大1秒待機
            
        # 親クラスのcloseEventを呼び出す
        super().closeEvent(event)

def main():
    try:
        # QApplication作成前にデフォルトスタイルを設定
        os.environ['QT_MAC_WANTS_LAYER'] = '1'
        
        # アプリケーション起動
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"アプリケーション起動エラー: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # 既存のQApplicationインスタンスがあれば削除
    app = QApplication.instance()
    if app is not None:
        del app
    
    # Pythonのガベージコレクションを強制実行
    import gc
    gc.collect()
    
    # シンプルに起動
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 