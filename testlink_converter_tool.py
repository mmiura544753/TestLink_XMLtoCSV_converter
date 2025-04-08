import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import traceback

# 他の処理モジュールをインポート
import xml_processor
import csv_processor

class TestLinkConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("TestLink XML-CSV Converter")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        # GUI要素の作成
        self.create_widgets()

        # ステータス表示の初期化
        self.update_status("待機中")

    def create_widgets(self):
        # 説明文ラベル
        self.lbl_description = tk.Label(
            self.root, 
            text="このツールは、TestLinkからエクスポートしたテストケースのXMLファイルを\nCSV形式に変換したり、CSVファイルからTestLink用のXMLファイルに\n変換したりするためのユーティリティです。",
            justify=tk.LEFT,
            anchor="w",
            pady=10
        )
        self.lbl_description.pack(fill=tk.X, padx=20)

        # XMLからCSV変換ボタン
        self.btn_xml_to_csv = tk.Button(self.root, text="XML→CSV変換", width=20, height=2,
                                        command=self.process_xml_to_csv)
        self.btn_xml_to_csv.pack(pady=10)

        # CSVからXML変換ボタン
        self.btn_csv_to_xml = tk.Button(self.root, text="CSV→XML変換", width=20, height=2,
                                        command=self.process_csv_to_xml)
        self.btn_csv_to_xml.pack(pady=10)

        # 終了ボタン
        self.btn_exit = tk.Button(self.root, text="終了", width=20, height=2,
                                  command=self.root.destroy)
        self.btn_exit.pack(pady=10)

        # ステータス表示ラベル
        self.lbl_status = tk.Label(self.root, text="ステータス: ", anchor="w")
        self.lbl_status.pack(fill=tk.X, padx=10, pady=10)

    def update_status(self, message):
        """ステータスメッセージを更新する"""
        self.lbl_status.config(text=f"ステータス: {message}")
        self.root.update()

    def process_xml_to_csv(self):
        """XMLファイルをCSVに変換するプロセス"""
        xml_file = filedialog.askopenfilename(
            title="XMLファイルを選択してください",
            filetypes=[("XMLファイル", "*.xml"), ("すべてのファイル", "*.*")]
        )
        if not xml_file:
            self.update_status("ファイルが選択されていません")
            return

        try:
            self.update_status("XMLファイルを解析中...")

            # XML内容の読み込みと修正（二重CDATAなど）
            with open(xml_file, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            fixed_xml_content = xml_processor.fix_double_cdata(xml_content)

            # XMLをパースしてルート要素とテストスイート名を取得
            root_element, testsuite_name = xml_processor.parse_xml_root(fixed_xml_content)

            # 出力CSVファイル名の生成
            output_file = os.path.splitext(xml_file)[0] + ".csv"

            # XMLからCSVへの変換処理を呼び出し
            xml_processor.convert_xml_to_csv(root_element, testsuite_name, output_file)

            self.update_status(f"変換完了: {output_file}")
            messagebox.showinfo("変換完了", f"CSVファイルに変換しました:\n{output_file}")

        except Exception as e:
            error_details = traceback.format_exc()
            self.update_status(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"XML→CSV変換中にエラーが発生しました:\n{str(e)}\n\n詳細:\n{error_details}")

    def process_csv_to_xml(self):
        """CSVファイルをXMLに変換するプロセス"""
        csv_file = filedialog.askopenfilename(
            title="CSVファイルを選択してください",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if not csv_file:
            self.update_status("ファイルが選択されていません")
            return

        try:
            self.update_status("CSVファイルを解析中...")

            # 出力XMLファイル名の生成
            output_file = os.path.splitext(csv_file)[0] + "_converted.xml"

            # CSVからXMLへの変換処理を呼び出し
            csv_processor.convert_csv_to_xml(csv_file, output_file)

            self.update_status(f"変換完了: {output_file}")
            messagebox.showinfo("変換完了", f"XMLファイルに変換しました:\n{output_file}")

        except Exception as e:
            error_details = traceback.format_exc()
            self.update_status(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"CSV→XML変換中にエラーが発生しました:\n{str(e)}\n\n詳細:\n{error_details}")

def main():
    """アプリケーションを起動する"""
    root = tk.Tk()
    app = TestLinkConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()