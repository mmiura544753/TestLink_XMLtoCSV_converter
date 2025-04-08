import csv
import codecs
import traceback

def read_csv_file(csv_file):
    """CSVファイルを読み込み、ヘッダーとデータ行を返す"""
    try:
        # CSV読み込み
        rows = []
        try:
            with codecs.open(csv_file, 'r', 'shift_jis', errors='replace') as f:
                reader = csv.reader(f)
                headers = next(reader)
                rows.append(headers)
                line_num = 1 # ヘッダーが1行目
                for row in reader:
                    line_num += 1
                    if len(row) != len(headers):
                         print(f"警告: 行 {line_num} の列数がヘッダー ({len(headers)}列) と異なります ({len(row)}列)。スキップします。")
                         continue
                    rows.append(row)
        except StopIteration:
             raise ValueError("CSVファイルにヘッダー行がありません")
        except FileNotFoundError:
             raise Exception(f"CSVファイルが見つかりません: {csv_file}")
        except Exception as e:
             raise Exception(f"CSVファイルの読み込み中にエラーが発生しました: {str(e)}\n{traceback.format_exc()}")

        if len(rows) < 2:
            raise ValueError("CSVファイルにデータ行がありません")

        return rows
    except Exception as e:
        raise Exception(f"CSVファイルの読み込み中にエラーが発生しました: {str(e)}")

def get_header_indices(headers):
    """ヘッダー行から各カラムのインデックスを取得する"""
    # 必須・オプショナルヘッダーの定義とインデックス取得
    required_headers = [
        "テストケース名", "バージョン", "サマリ（概要）", "重要度",
        "ステップ番号", "アクション（手順）", "期待結果", "実行タイプ"
    ]
    # IDと外部IDは必須ではない（新規作成のため）
    optional_headers = ["ID", "外部ID", "事前条件", "推定実行時間", "ステータス", "有効/無効", "開いているか", "親テストスイート名"]
    
    # カスタムフィールドのリスト
    custom_field_headers = [
        "AutomationAction", "AutomationParameters", "AutomationEnabled", 
        "AutomationTargetNode", "AutomationValidation"
    ]
    
    header_indices = {}
    missing_required = []
    for header in required_headers:
        try:
            header_indices[header] = headers.index(header)
        except ValueError:
            missing_required.append(header)
    if missing_required:
        raise ValueError(f"CSVファイルに必要なヘッダーが見つかりません: {', '.join(missing_required)}")

    for header in optional_headers:
        try:
            header_indices[header] = headers.index(header)
        except ValueError:
            header_indices[header] = -1 # 見つからない場合は -1
    
    # カスタムフィールドのインデックスを取得
    header_indices["custom_fields"] = {}
    for cf_header in custom_field_headers:
        try:
            cf_index = headers.index(cf_header)
            header_indices["custom_fields"][cf_header] = cf_index
        except ValueError:
            # カスタムフィールドは見つからない場合も問題ない
            pass
            
    return header_indices