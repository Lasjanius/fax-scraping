# クリニックFAX番号収集ツール

## 概要
このプロジェクトは、クリニックのFAX番号を自動的にスクレイピングして収集するためのツールです。PyQt5ベースのGUIアプリケーションとして実装されており、CSVファイルから読み込んだクリニック名のリストに対して、Google検索を行い、見つかったウェブサイトからFAX番号を抽出します。

## 主要ファイル
- `fax_scraper_qt.py`: メインのGUIアプリケーション（PyQt5実装）まだ使えないです
- `tsurumiku_import_requests.py`: 鶴見区関連データのインポートスクリプト。鶴見区医師会の検索結果から抽出する際に使いました
- `pdf_to_text.py`: PDFからテキストを抽出するツール。厚生局のデータPDFをここにぶち込むといい。これが一番ちゃんと使える
- `tsurumi-fax_numbers_with_area_code.csv`: サンプルデータまたは結果ファイル
- `test_*.py`: 各種テストスクリプト

## 必要条件
- Python 3.6 以上
- 必要ライブラリ：PyQt5, pandas, requests, beautifulsoup4, googlesearch-python

## インストール方法

1. リポジトリをクローン
```bash
git clone https://github.com/Lasjanius/fax-scraping.git
cd fax-scraping
```

2. 仮想環境を作成して有効化（推奨）
```bash
python -m venv venv
# Windowsの場合
venv\Scripts\activate
# macOS/Linuxの場合
source venv/bin/activate
```

3. 必要なパッケージをインストール
```bash
pip install -r requirements.txt
```

注意: `requirements.txt`がリポジトリに含まれていない場合は、以下のコマンドで必要なパッケージを個別にインストールしてください。
```bash
pip install pyqt5 pandas requests beautifulsoup4 google
```

## 使用方法

### GUI版（推奨）
1. メインアプリケーションを起動します
```bash
python fax_scraper_qt.py
```

2. 「参照...」ボタンをクリックしてクリニック名が含まれるCSVファイルを選択します。
3. 「実行」ボタンをクリックして処理を開始します。
4. 処理中に一時停止する場合は「中止」ボタンを使用します。
5. 処理が中断された場合は「リフレッシュ」ボタンで続きから再開できます。

### 出力
- 入力CSVファイルに「FAX番号」列が追加され、各クリニックのFAX番号が追記されます。
- 情報が見つからなかった場合は「エラー詳細」列にエラー内容が記録されます。

## 機能詳細
- Google検索を使用してクリニックのウェブサイトを特定
- 複数のパターンによるFAX番号の検出
- User-Agentのランダム化によるブロック回避
- 待機時間の調整による検索制限の回避
- 途中からの処理再開機能

## 注意事項
- Googleの使用制限に注意してください。大量の検索を行うと一時的にブロックされる可能性があります。
- スクレイピングはウェブサイトの利用規約に違反する可能性があります。使用前に法的な確認を行ってください。
- 取得したデータの利用に関しては適切な法令や規制を遵守してください。

## ライセンス
このプロジェクトは個人利用目的で作成されました。商用利用には作者の許可が必要です。 