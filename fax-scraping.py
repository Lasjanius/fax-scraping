import sys
import os
import csv
import time
import json
import re
import threading
import pandas as pd
import requests
from bs4 import BeautifulSoup
from googlesearch import search
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


class ClinicFaxScraperGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("クリニックFAX番号収集ツール")
        self.master.geometry("600x450")
        self.master.resizable(True, True)
        
        # 変数初期化
        self.csv_path = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="ファイルを選択してください")
        self.current_clinic_var = tk.StringVar(value="")
        self.total_clinics = 0
        self.processed_clinics = 0
        
        self.create_widgets()
        
    def create_widgets(self):
        # フレーム作成
        input_frame = ttk.LabelFrame(self.master, text="入力ファイル")
        input_frame.pack(padx=10, pady=10, fill=tk.X)
        
        status_frame = ttk.LabelFrame(self.master, text="処理状況")
        status_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        button_frame = ttk.Frame(self.master)
        button_frame.pack(padx=10, pady=10, fill=tk.X)
        
        # 入力ファイル選択
        ttk.Label(input_frame, text="CSVファイル:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.csv_path, width=50).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        ttk.Button(input_frame, text="参照...", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        
        # 処理状況表示
        ttk.Label(status_frame, text="ステータス:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(status_frame, textvariable=self.status_var).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(status_frame, text="処理中のクリニック:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(status_frame, textvariable=self.current_clinic_var).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(status_frame, text="進捗状況:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # ログ表示エリア
        ttk.Label(status_frame, text="ログ:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.NW)
        self.log_text = tk.Text(status_frame, height=10, width=50)
        self.log_text.grid(row=3, column=1, padx=5, pady=5, sticky=tk.NSEW)
        
        scrollbar = ttk.Scrollbar(status_frame, command=self.log_text.yview)
        scrollbar.grid(row=3, column=2, sticky=tk.NS)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # ボタン
        ttk.Button(button_frame, text="実行", command=self.start_scraping).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="中止", command=self.stop_scraping).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="閉じる", command=self.master.quit).pack(side=tk.RIGHT, padx=5)
        
        # 行と列の設定
        status_frame.grid_columnconfigure(1, weight=1)
        status_frame.grid_rowconfigure(3, weight=1)
        
        # スクレイピングスレッド
        self.scraping_thread = None
        self.stop_requested = False
        
    def browse_file(self):
        filetypes = [("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.csv_path.set(filename)
            self.log(f"ファイル選択: {filename}")
            self.analyze_csv()
    
    def analyze_csv(self):
        try:
            # CSVファイルを読み込んでクリニック数を確認
            df = pd.read_csv(self.csv_path.get())
            self.total_clinics = len(df)
            self.status_var.set(f"読み込み完了: {self.total_clinics}件のクリニックが見つかりました")
            self.log(f"クリニック総数: {self.total_clinics}件")
            
            # FAX番号カラムの有無を確認
            if 'FAX番号' in df.columns:
                self.log("既存のFAX番号カラムが見つかりました")
            else:
                self.log("FAX番号カラムが見つかりません。新規に作成します")
                
            # エラー詳細カラムの有無を確認
            if 'エラー詳細' in df.columns:
                self.log("既存のエラー詳細カラムが見つかりました")
            else:
                self.log("エラー詳細カラムが見つかりません。新規に作成します")
                
        except Exception as e:
            self.status_var.set("エラー: CSVファイルの読み込みに失敗しました")
            self.log(f"エラー: {str(e)}")
    
    def start_scraping(self):
        if not self.csv_path.get():
            messagebox.showerror("エラー", "CSVファイルを選択してください")
            return
        
        if self.scraping_thread and self.scraping_thread.is_alive():
            messagebox.showinfo("情報", "すでに処理中です")
            return
        
        self.stop_requested = False
        self.scraping_thread = threading.Thread(target=self.scrape_fax_numbers)
        self.scraping_thread.daemon = True
        self.scraping_thread.start()
    
    def stop_scraping(self):
        if self.scraping_thread and self.scraping_thread.is_alive():
            self.stop_requested = True
            self.status_var.set("処理を中止しています...")
            self.log("処理中止リクエストを受け付けました。現在の処理が完了次第停止します")
        else:
            self.log("停止する処理がありません")
    
    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
    
    def update_progress(self, clinic_name):
        self.processed_clinics += 1
        progress = (self.processed_clinics / self.total_clinics) * 100
        self.progress_var.set(progress)
        self.current_clinic_var.set(clinic_name)
        self.status_var.set(f"処理中... {self.processed_clinics}/{self.total_clinics} ({progress:.1f}%)")
    
    def scrape_fax_numbers(self):
        try:
            self.status_var.set("処理を開始しています...")
            self.processed_clinics = 0
            self.progress_var.set(0)
            
            # CSVファイルを読み込む
            df = pd.read_csv(self.csv_path.get())
            
            # FAX番号カラムがなければ追加
            if 'FAX番号' not in df.columns:
                df['FAX番号'] = None
                
            # エラー詳細カラムがなければ追加
            if 'エラー詳細' not in df.columns:
                df['エラー詳細'] = None
            
            # 各クリニックに対して処理
            for index, row in df.iterrows():
                if self.stop_requested:
                    self.log("処理を中止しました")
                    self.status_var.set("処理中止")
                    break
                
                clinic_name = row.iloc[0]  # インデックス列はCSVの最初の列と仮定
                self.current_clinic_var.set(clinic_name)
                self.log(f"処理中: {clinic_name}")
                
                # エラー詳細をリセット
                df.at[index, 'エラー詳細'] = None
                
                # すでにFAX番号がある場合はスキップ
                if not pd.isna(df.at[index, 'FAX番号']):
                    self.log(f"- すでにFAX番号があります: {df.at[index, 'FAX番号']}")
                else:
                    try:
                        # URLがあれば使用、なければ検索
                        url = None
                        try:
                            if 'URL' in df.columns and not pd.isna(row['URL']):
                                url = row['URL']
                                self.log(f"- 既存のURLを使用: {url}")
                            else:
                                url = self.search_clinic_url(clinic_name)
                                if url:
                                    self.log(f"- URL検索結果: {url}")
                                else:
                                    raise Exception("URLを検索できませんでした")
                        except Exception as url_error:
                            df.at[index, 'エラー詳細'] = f"URL取得エラー: {str(url_error)}"
                            self.log(f"- {df.at[index, 'エラー詳細']}")
                            continue
                        
                        # FAX番号の取得
                        try:
                            fax_number = self.get_fax_number(url)
                            if fax_number:
                                df.at[index, 'FAX番号'] = fax_number
                                self.log(f"- FAX番号を取得: {fax_number}")
                            else:
                                df.at[index, 'エラー詳細'] = "FAX番号が見つかりませんでした"
                                self.log(f"- {df.at[index, 'エラー詳細']}")
                        except Exception as fax_error:
                            df.at[index, 'エラー詳細'] = f"FAX番号取得エラー: {str(fax_error)}"
                            self.log(f"- {df.at[index, 'エラー詳細']}")
                    except Exception as e:
                        df.at[index, 'エラー詳細'] = f"予期せぬエラー: {str(e)}"
                        self.log(f"- {df.at[index, 'エラー詳細']}")
                
                # 進捗更新
                self.update_progress(clinic_name)
                
                # 一時保存（10件ごと）
                if (index + 1) % 10 == 0 or index == len(df) - 1:
                    temp_path = self.csv_path.get().replace('.csv', '_temp.csv')
                    df.to_csv(temp_path, index=False)
                    self.log(f"中間保存完了: {temp_path}")
                
                # サーバー負荷軽減のための待機
                time.sleep(2)
            
            # 処理完了、結果を保存
            output_path = self.csv_path.get().replace('.csv', '_result.csv')
            df.to_csv(output_path, index=False)
            
            # エラー集計
            error_count = df['エラー詳細'].notna().sum()
            success_count = self.processed_clinics - error_count
            
            self.status_var.set(f"処理完了: {self.processed_clinics}/{self.total_clinics}")
            self.log(f"処理が完了しました。結果は {output_path} に保存されました")
            self.log(f"成功: {success_count}件、エラー: {error_count}件")
            messagebox.showinfo("完了", f"処理が完了しました。\n処理件数: {self.processed_clinics}/{self.total_clinics}\n成功: {success_count}件、エラー: {error_count}件")
            
        except Exception as e:
            self.status_var.set("エラーが発生しました")
            self.log(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"処理中にエラーが発生しました: {str(e)}")
    
    def search_clinic_url(self, clinic_name):
        """クリニック名からURLを検索する"""
        query = f"{clinic_name} 公式サイト"
        try:
            for url in search(query, num_results=1):
                return url
            raise Exception("検索結果が0件でした")
        except Exception as e:
            raise Exception(f"URL検索中にエラー発生: {str(e)}")
    
    def get_fax_number(self, url):
        """複数の手法を順番に試して最適なFAX番号抽出を行う"""
        # 1. 通常のリクエストでHTMLを取得
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        try:
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            raise Exception(f"サイトアクセスエラー: {str(e)}")
            
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 2. 連絡先セクションを特定
        contact_sections = self.find_contact_section(soup)
        for section in contact_sections:
            if isinstance(section, list):
                for element in section:
                    section_soup = BeautifulSoup(str(element), 'html.parser')
                    fax = self.extract_fax_number(section_soup)
                    if fax:
                        return fax
            else:
                fax = self.extract_fax_number(section)
                if fax:
                    return fax
        
        # 3. 電話番号の近くにあるFAX番号を探す
        fax = self.find_fax_near_phone(soup)
        if fax:
            return fax
        
        # 4. 一般的なFAX番号抽出
        fax = self.extract_fax_number(soup)
        if fax:
            return fax
        
        # 5. FAX番号が見つからなかった
        return None
    
    def find_contact_section(self, soup):
        """連絡先情報が含まれる可能性が高いセクションを特定"""
        contact_keywords = ['contact', 'アクセス', '連絡', 'お問い合わせ', '診療時間', '案内', 'info']
        
        # IDやクラスでセクションを探す
        elements = []
        for keyword in contact_keywords:
            elements += soup.find_all(attrs={"id": re.compile(keyword, re.IGNORECASE)})
            elements += soup.find_all(attrs={"class": re.compile(keyword, re.IGNORECASE)})
        
        if elements:
            return elements
        
        # 見出しでセクションを探す
        headings = soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
        contact_sections = []
        for heading in headings:
            if any(keyword in heading.get_text().lower() for keyword in contact_keywords):
                # 見出し以降の兄弟要素を取得
                section = [heading]
                current = heading
                while current.next_sibling and not (hasattr(current.next_sibling, 'name') and 
                                                   current.next_sibling.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                    current = current.next_sibling
                    if current:
                        section.append(current)
                contact_sections.append(section)
        
        return contact_sections
    
    def find_fax_near_phone(self, soup):
        """電話番号の近くにあるFAX番号を探索"""
        # 電話番号パターン
        phone_pattern = r'(?:電話|TEL|Tel|tel)[番号]*[\s:：][\s\d\-()（）]+\d{3,4}[\-\d()（）]+'
        
        # 電話番号を含む要素を見つける
        text = soup.get_text()
        phone_matches = list(re.finditer(phone_pattern, text, re.IGNORECASE))
        
        for phone_match in phone_matches:
            # 電話番号の後ろ50文字以内にFAX番号があるか確認
            phone_end = phone_match.end()
            search_text = text[phone_end:phone_end+150]
            
            fax_pattern = r'(?:FAX|ファックス|ＦＡＸ)[番号]*[\s:：][\s\d\-()（）]+\d{3,4}[\-\d()（）]+'
            fax_match = re.search(fax_pattern, search_text, re.IGNORECASE)
            
            if fax_match:
                number_only = re.search(r'\d[\d\-()（）\s]{7,}', fax_match.group(0))
                if number_only:
                    return self.clean_fax_number(number_only.group(0))
        
        return None
    
    def extract_fax_number(self, soup):
        """
        複数のパターンを試してFAX番号を抽出する高度な関数
        """
        # 1. まず明示的な「FAX:」パターンを探す
        patterns = [
            r'(?:FAX|ファックス|ＦＡＸ)[番号]*[\s:：][\s\d\-()（）]+\d{3,4}[\-\d()（）]+',
            r'F[\s\.]*A[\s\.]*X[\s:：]*[\s\d\-()（）]+\d{3,4}[\-\d()（）]+',
            r'(?<!携帯電話)(?<!携帯)(?<!電話)(?<!TEL)(?<!Tel)(?<!tel)[\(（]?[\s]*(?:FAX|ファックス|ＦＡＸ)[\)）]?[\s]*[\(（]?\d{2,5}[\)）]?[\-−‐]?\d{1,4}[\-−‐]?\d{3,4}'
        ]
        
        text = soup.get_text()
        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                fax_text = match.group(0)
                # 番号部分だけを抽出する追加処理
                number_only = re.search(r'\d[\d\-()（）\s]{7,}', fax_text)
                if number_only:
                    return self.clean_fax_number(number_only.group(0))
        
        # 2. ページ内の「FAX」という単語の近くにある電話番号パターンを探す
        fax_keywords = ['fax', 'ファックス', 'ＦＡＸ', 'ファクス']
        for keyword in fax_keywords:
            elements = soup.find_all(string=re.compile(keyword, re.IGNORECASE))
            for element in elements:
                # 親要素とその周辺のテキストを調査
                parent = element.parent
                surrounding_text = parent.get_text() if parent else ""
                
                # 親の親も確認（テーブルセルなどの場合）
                if parent and parent.parent:
                    surrounding_text += parent.parent.get_text()
                
                # 電話番号パターンを探す
                phone_pattern = r'\d{2,5}[\-\(\)（）]?\d{1,4}[\-\(\)（）]?\d{3,4}'
                phone_matches = re.finditer(phone_pattern, surrounding_text)
                
                # 最も「FAX」に近い番号を選ぶ
                fax_pos = surrounding_text.lower().find(keyword.lower())
                closest_number = None
                min_distance = float('inf')
                
                for phone_match in phone_matches:
                    phone_pos = phone_match.start()
                    distance = abs(phone_pos - fax_pos)
                    if distance < min_distance and distance < 100:  # 距離の閾値
                        min_distance = distance
                        closest_number = phone_match.group(0)
                
                if closest_number:
                    return self.clean_fax_number(closest_number)
        
        # 3. テーブル構造内でFAXを探す
        tables = soup.find_all('table')
        for table in tables:
            # テーブルヘッダーでFAXを探す
            headers = table.find_all(['th', 'td'])
            for header in headers:
                if any(keyword.lower() in header.get_text().lower() for keyword in fax_keywords):
                    # FAXヘッダーの次か同じ行のセルを確認
                    next_cells = header.find_next_siblings(['th', 'td'])
                    if next_cells:
                        for cell in next_cells:
                            phone_pattern = r'\d{2,5}[\-\(\)（）]?\d{1,4}[\-\(\)（）]?\d{3,4}'
                            phone_match = re.search(phone_pattern, cell.get_text())
                            if phone_match:
                                return self.clean_fax_number(phone_match.group(0))
                    
                    # 同じ列の下のセルを確認
                    row = header.parent
                    if row:
                        try:
                            header_index = list(row.children).index(header)
                            for next_row in row.find_next_siblings('tr'):
                                cells = list(next_row.children)
                                if header_index < len(cells):
                                    cell = cells[header_index]
                                    phone_pattern = r'\d{2,5}[\-\(\)（）]?\d{1,4}[\-\(\)（）]?\d{3,4}'
                                    phone_match = re.search(phone_pattern, cell.get_text())
                                    if phone_match:
                                        return self.clean_fax_number(phone_match.group(0))
                        except ValueError:
                            # headerがrow.childrenのリストに存在しない場合
                            continue
        
        # 4. classやid属性にFAX関連のキーワードがある要素を確認
        for keyword in fax_keywords:
            fax_elements = soup.find_all(attrs={"class": re.compile(keyword, re.IGNORECASE)})
            fax_elements += soup.find_all(attrs={"id": re.compile(keyword, re.IGNORECASE)})
            for element in fax_elements:
                phone_pattern = r'\d{2,5}[\-\(\)（）]?\d{1,4}[\-\(\)（）]?\d{3,4}'
                phone_match = re.search(phone_pattern, element.get_text())
                if phone_match:
                    return self.clean_fax_number(phone_match.group(0))
        
        return None
    
    def clean_fax_number(self, number):
        """
        抽出したFAX番号を整形する
        """
        # 数字、ハイフン、括弧以外の文字を削除
        cleaned = re.sub(r'[^\d\-\(\)（）]', '', number)
        # 標準的な形式に整形（オプション）
        # 例: 0312345678 → 03-1234-5678
        if re.match(r'^\d+$', cleaned):  # 数字のみの場合
            if len(cleaned) == 10:  # 市外局番2桁
                cleaned = f"{cleaned[:2]}-{cleaned[2:6]}-{cleaned[6:]}"
            elif len(cleaned) == 11:  # 市外局番3桁
                cleaned = f"{cleaned[:3]}-{cleaned[3:7]}-{cleaned[7:]}"
        return cleaned


if __name__ == "__main__":
    root = tk.Tk()
    app = ClinicFaxScraperGUI(root)
    root.mainloop()

