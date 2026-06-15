import pdfplumber
import re
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import messagebox

def start_transfer(pdf_paths):
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        ws = excel.ActiveSheet 

        # 💡 関数セルを正確に見分け、新行(41行目〜)にも無限追従するループ
        target_row = 0
        r = 5
        while True:
            val_b = ws.Cells(r, 2).Value
            val_c = ws.Cells(r, 3).Value
            form_b = str(ws.Cells(r, 2).Formula).strip()
            
            str_b = str(val_b).strip() if val_b is not None else ""
            str_c = str(val_c).strip() if val_c is not None else ""
            
            if str_b == "" and str_c == "" and (form_b == "" or form_b == "0" or form_b == str_b):
                target_row = r
                break
            r += 1

        def to_int(v):
            if not v: return 0
            s = str(v).replace(',', '').replace('▲', '-').replace('△', '-')
            res = re.findall(r'-?\d+', s)
            return int(res[0]) if res else 0

        # ==========================================================
        # 🔍 【事前チェック】書き込む総行数を事前にカウントする
        # ==========================================================
        total_lines_to_write = 0
        for pdf_path in pdf_paths:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                text_full = page.extract_text()
                for line in text_full.split('\n'):
                    if re.match(r'^\d+\s+', line):
                        parts = line.split()
                        if len(parts) < 5: continue
                        
                        try:
                            amount = to_int(parts[-2])
                            price  = to_int(parts[-3])
                            total_lines_to_write += 1
                        except: continue
                            
                    elif "値引き" in line:
                        total_lines_to_write += 1

        # 💡 Excelに書き込める限界の行（40行目がSUMなら、39行目まで）
        max_allowable_row = 39
        # 今回のデータが到達する予定の最終行
        expected_end_row = target_row + total_lines_to_write - 1

        # 🪟 ポップアップ用の画面を最前面で初期化
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)

        # ⚠️ もし39行目をはみ出る場合の案内ウィンドウ
        if expected_end_row > max_allowable_row:
            overflow_count = expected_end_row - max_allowable_row
            ans = messagebox.askyesno(
                "⚠️ 行数オーバーの確認", 
                f"【転記エラーを未然に防ぐための案内】\n\n"
                f"読み込んだPDFのデータ量が多いため、\n"
                f"指定の枠（39行目）を 【 {overflow_count} 行 】 はみ出します。\n\n"
                f"このまま合計行(40行目)を無視して上書き転記しますか？\n\n"
                f"※安全のため、一度『いいえ』を選んで処理を中断し、\n"
                f"Excel側で必要行数を右クリック挿入してから、再実行することをおすすめします。"
            )
            # 「いいえ」を押したらエクセルを汚さずに安全に終了
            if not ans:
                return "ユーザーによりキャンセル（行数不足）"

        # ==========================================================
        # ✍️ 【実際の書き込み処理】
        # ==========================================================
        current_row = target_row
        for pdf_path in pdf_paths:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                
                words = page.extract_words()
                vendor_parts = [w['text'] for w in words if w['x0'] > page.width * 0.55 and w['bottom'] < page.height * 0.3]
                joined_vendor = "".join(vendor_parts)
                clean_vendor = re.sub(r'発行日[:：]?\d{4}/\d{2}/\d{2}', '', joined_vendor)
                clean_vendor = re.sub(r'〒?\d{3}-\d{4}.*', '', clean_vendor)
                clean_vendor = re.sub(r'(?:東京都|北海道|(?:京都|大阪)府|.{2,3}県).*', '', clean_vendor)
                clean_vendor = re.sub(r'(請求|No|　|住所|TEL[:：]?.*)', '', clean_vendor)
                fixed_vendor = clean_vendor.strip()

                text_full = page.extract_text()
                date_match = re.search(r'(\d{4}/\d{2}/\d{2})', text_full)
                invoice_date = date_match.group(1) if date_match else "2026/4/2"

                for line in text_full.split('\n'):
                    if re.match(r'^\d+\s+', line):
                        parts = line.split()
                        if len(parts) < 5: continue
                        
                        try:
                            amount = to_int(parts[-2])
                            price  = to_int(parts[-3])
                            unit   = parts[-4]
                            qty    = parts[-5].replace(',', '')
                            item_name = " ".join(parts[1:-5])

                            tax = int(amount * 0.1)
                            total = amount + tax

                            ws.Cells(current_row, 2).Value = invoice_date
                            ws.Cells(current_row, 3).Value = fixed_vendor
                            ws.Cells(current_row, 4).Value = item_name
                            ws.Cells(current_row, 5).Value = qty
                            ws.Cells(current_row, 6).Value = unit
                            ws.Cells(current_row, 7).Value = price
                            ws.Cells(current_row, 8).Value = amount
                            ws.Cells(current_row, 9).Value = tax
                            ws.Cells(current_row, 10).Value = total
                            current_row += 1
                        except: continue
                            
                    elif "値引き" in line:
                        p = line.split()
                        val = to_int(p[-1])
                        if val > 0: val *= -1
                        
                        tax_neg = int(val * 0.1)
                        total_neg = val + tax_neg
                        
                        ws.Cells(current_row, 2).Value = invoice_date
                        ws.Cells(current_row, 3).Value = fixed_vendor
                        ws.Cells(current_row, 4).Value = "特別値引き"
                        ws.Cells(current_row, 8).Value = val          
                        ws.Cells(current_row, 9).Value = tax_neg      
                        ws.Cells(current_row, 10).Value = total_neg   
                        current_row += 1

        # ✨ 無事に最後まで入った場合の完了ウィンドウ
        messagebox.showinfo("転記完了", "すべてのPDFデータの転記が正常に完了しました！")
        return "完了"
    except Exception as e:
        raise Exception(str(e))
    finally:
        pythoncom.CoUninitialize()
