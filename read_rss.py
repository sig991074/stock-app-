import openpyxl
import time
import os

# CドライブのExcelファイルの場所
excel_file_path = r"C:\Users\sigen\Documents\rakuten_rss.xlsx"

print("🔄 楽天RSS（10銘柄モニター）の読み込みを開始します（Ctrl+Cで終了）...")

while True:
    try:
        # Excelを数値読み取り専用モードで超高速オープン
        wb = openpyxl.load_workbook(excel_file_path, data_only=True, keep_vba=False)
        sheet = wb.active

        # 画面を綺麗にクリア
        os.system('cls' if os.name == 'nt' else 'clear')
        print("==================== リアルタイム 10銘柄同時モニター ====================")
        print("銘柄名称             | 現在値   | 始値     | 高値     | 安値     | 銘柄コード")
        print("------------------------------------------------------------------------")

        # 【修正！】1行目は見出しなので飛ばし、「2行目から11行目まで」を処理します
        for row in range(2, 12):
            stock_name    = sheet.cell(row=row, column=1).value # A列 (1)
            current_price = sheet.cell(row=row, column=2).value # B列 (2)
            open_price    = sheet.cell(row=row, column=3).value # C列 (3)
            high_price    = sheet.cell(row=row, column=4).value # D列 (4)
            low_price     = sheet.cell(row=row, column=5).value # E列 (5)
            stock_code    = sheet.cell(row=row, column=9).value # I列 (9)

            # 空白の行があればスキップする（10番目の空欄対策）
            if not stock_name:
                continue

            # 表示の幅をきれいに揃える
            print(f"{str(stock_name):<20} | {current_price:<8} | {open_price:<8} | {high_price:<8} | {low_price:<8} | {stock_code}")

        print("================================================================--------")
        print("※ 1秒ごとに自動で最新情報に更新中...")

        wb.close()
        time.sleep(2) # 2秒ごとに同時更新

    except FileNotFoundError:
        print(f"❌ エラー: {excel_file_path} が見つかりません。")
        time.sleep(3)
    except PermissionError:
        time.sleep(0.5)
    except KeyboardInterrupt:
        print("\n⏹️ モニターを終了しました。")
        break
