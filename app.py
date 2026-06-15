import streamlit as st
import pandas as pd
import pdfplumber
import re
import os

# 💡 画面を横いっぱいに広く使う設定
st.set_page_config(layout="wide", page_title="小野寺システム お試しページ")

st.title("請求書の自動転記体験ホームページ")

# ==========================================================
# 📊 【画面の上半分】エクセル画面の再現
# ==========================================================
st.subheader("📊 現在開いているエクセルの画面")

# 転記されたデータを記憶しておく箱（セッション状態）
if "excel_rows" not in st.session_state:
    st.session_state.excel_rows = [
        {"行": 5, "計上日": "", "取引先": "", "項目": "", "数量": 0, "単位": "", "単価": 0, "金額（税抜）": 0, "消費税": 0, "合計（税込）": 0, "備考": ""},
        {"行": 6, "計上日": "", "取引先": "", "項目": "", "数量": 0, "単位": "", "単価": 0, "金額（税抜）": 0, "消費税": 0, "合計（税込）": 0, "備考": ""},
        {"行": 7, "計上日": "", "取引先": "", "項目": "", "数量": 0, "単位": "", "単価": 0, "金額（税抜）": 0, "消費税": 0, "合計（税込）": 0, "備考": ""},
        {"行": 8, "計上日": "", "取引先": "", "項目": "", "数量": 0, "単位": "", "単価": 0, "金額（税抜）": 0, "消費税": 0, "合計（税込）": 0, "備考": ""},
        {"行": 9, "計上日": "", "取引先": "", "項目": "", "数量": 0, "単位": "", "単価": 0, "金額（税抜）": 0, "消費税": 0, "合計（税込）": 0, "備考": ""}
    ]

if "current_pdf" not in st.session_state:
    st.session_state.current_pdf = None

# 表示用にきれいな表に変換
df_display = pd.DataFrame(st.session_state.excel_rows)

# ─── 確実に対象の列名を直接指定して合計します ───
total_kingaku = pd.to_numeric(df_display["金額（税抜）"], errors='coerce').fillna(0).sum()
total_zei = pd.to_numeric(df_display["消費税"], errors='coerce').fillna(0).sum()
total_gokei = pd.to_numeric(df_display["合計（税込）"], errors='coerce').fillna(0).sum()

# 合計行の作成
total_row = pd.DataFrame([{
    "計上日": "【合計】", "取引先": "", "項目": "", "数量": "", "単位": "", "単価": "", 
    "金額（税抜）": int(total_kingaku), "消費税": int(total_zei), "合計（税込）": int(total_gokei), "備考": ""
}])

# いつもの表と合計の行をドッキング
df_final = pd.concat([df_display.drop(columns=["行"]), total_row], ignore_index=True)

# 画面に表を表示する（左端の行番号は非表示）
st.dataframe(df_final, use_container_width=True, hide_index=True)

# リセットボタン
if st.button("エクセル画面をクリアして最初から試す"):
    del st.session_state.excel_rows
    del st.session_state.current_pdf
    st.rerun()

st.markdown("---")

# ==========================================================
# 📄 【裏側の処理】PDFを解析してセッションに流し込む関数
# ==========================================================
def to_int(v):
    if not v: return 0
    s = str(v).replace(',', '').replace('▲', '-').replace('△', '-')
    res = re.findall(r'-?\d+', s)
    return int(res) if res else 0

def run_analysis(pdf_file_obj):
    try:
        with pdfplumber.open(pdf_file_obj) as pdf:
            page = pdf.pages[0]
            text_full = page.extract_text()
            
            words = page.extract_words()
            vendor_area = [w['text'] for w in words if w['x0'] > page.width * 0.5 and w['bottom'] < page.height * 0.3]
            vendor_text = "".join(vendor_area)
            
            current_rows = []
            date_match = re.search(r'(\d{4}/\d{2}/\d{2})', text_full)
            invoice_date = date_match.group(1) if date_match else "2026/04/02"

            # 1️⃣ 小野寺ネットワークス の場合
            if "小野寺ネットワークス" in vendor_text or "小野寺ネットワークス" in text_full:
                vendor_name = "(株)小野寺ネットワークス"
                for line in text_full.split('\n'):
                    if re.match(r'^\d+\s+', line):
                        p = line.split()
                        if len(p) < 6: continue
                        amt = to_int(p[-2])
                        current_rows.append({
                            "計上日": invoice_date, "取引先": vendor_name, "項目": " ".join(p[1:-5]),
                            "数量": p[-5], "単位": p[-4], "単価": to_int(p[-3]), "金額（税抜）": amt,
                            "消費税": int(amt * 0.1), "合計（税込）": int(amt * 1.1), "備考": ""
                        })

            # 2️⃣ 小野寺企画（飲食事業部） の場合
            elif "小野寺企画" in vendor_text or "小野寺企画" in text_full:
                vendor_name = "小野寺企画・飲食事業部"
                for line in text_full.split('\n'):
                    if re.match(r'^\d+\.\s+', line) or re.match(r'^\d+\s+', line):
                        p = line.split()
                        if len(p) < 6: continue
                        amt = to_int(p[-2])
                        tax_rate = 0.1 if "10" in p[-1] else 0.08
                        current_rows.append({
                            "計上日": invoice_date, "取引先": vendor_name, "項目": " ".join(p[1:-5]),
                            "数量": p[-5], "単位": p[-4], "単価": to_int(p[-3]), "金額（税抜）": amt,
                            "消費税": int(amt * tax_rate), "合計（税込）": amt + int(amt * tax_rate), "備考": ""
                        })

            # 3️⃣ それ以外の複雑な請求書（値引き対応） の場合
            else:
                clean_vendor = re.sub(r'発行日[:：]?\d{4}/\d{2}/\d{2}|〒?\d{3}-\d_4}.*|(?:東京都|北海道|(?:京都|大阪)府|.{2,3}県).*|(?:請求|No|　|住所|TEL[:：]?.*)', '', "".join([w['text'] for w in words if w['x0'] > page.width * 0.55 and w['bottom'] < page.height * 0.3])).strip()
                vendor_name = clean_vendor if clean_vendor else "株式会社 総合建築"
                
                for line in text_full.split('\n'):
                    if re.match(r'^\d+\s+', line):
                        parts = line.split()
                        if len(parts) < 5: continue
                        try:
                            amt = to_int(parts[-2])
                            current_rows.append({
                                "計上日": invoice_date, "取引先": vendor_name, "項目": " ".join(parts[1:-5]),
                                "数量": parts[-5].replace(',', ''), "単位": parts[-4], "単価": to_int(parts[-3]), "金額（税抜）": amt,
                                "消費税": int(amt * 0.1), "合計（税込）": amt + int(amt * 0.1), "備考": ""
                            })
                        except: continue
                    elif "値引き" in line:
                        p = line.split()
                        val = to_int(p[-1])
                        if val > 0: val *= -1
                        current_rows.append({
                            "計上日": invoice_date, "取引先": vendor_name, "項目": "特別値引き",
                            "数量": "", "単位": "", "単価": "", "金額（税抜）": val,
                            "消費税": int(val * 0.1), "合計（税込）": val + int(val * 0.1), "備考": ""
                        })

        if current_rows:
            active_rows = [r for r in st.session_state.excel_rows if r["取引先"] != ""]
            for r in current_rows:
                active_rows.append(r)
            row_idx = 5
            for r in active_rows:
                r["行"] = row_idx
                row_idx += 1
            while len(active_rows) < 5:
                active_rows.append({"行": row_idx, "計上日": "", "取引先": "", "項目": "", "数量": 0, "単位": "", "単価": 0, "金額（税抜）": 0, "消費税": 0, "合計（税込）": 0, "備考": ""})
                row_idx += 1
            st.session_state.excel_rows = active_rows
        else:
            st.error("⚠️ PDFから有効な明細データが見つかりませんでした。")
            
    except Exception as e:
        st.error(f"❌ 読み込みエラーが発生しました: {e}")

# 📢 最初からテキストとして中身を表示するための関数
def get_pdf_preview_text(file_path):
    if os.path.exists(file_path):
        with pdfplumber.open(file_path) as pdf:
            return pdf.pages[0].extract_text()
    return "ファイルが読み込めません。"

# ==========================================================
# 📄 【画面の下半分】お試しボタンと請求書テキストの常時表示
# ==========================================================
st.subheader("📄 お試し用 請求書データ一覧（中身を見て転記ボタンを押してね）")
st.info("💡 下の「⚡ この内容を転記する」を押すと、枠内のテキストデータが自動解析されて上のエクセルに入ります！")

btn_col1, btn_col2, btn_col3 = st.columns(3)

selected_file = None

with btn_col1:
    st.write("### 🏢 株式会社 総合建築")
    if st.button("⚡ この内容を転記する", key="btn1", use_container_width=True):
        if os.path.exists("Final_Const.pdf"): selected_file = "Final_Const.pdf"
    
    # 📢 最初からテキストとして請求書の中身を画面に常駐させる
    st.text_area("📄 請求書テキスト（Final_Const.pdf）", value=get_pdf_preview_text("Final_Const.pdf"), height=300, disabled=True)

with btn_col2:
    st.write("### 🌐 小野寺ネットワークス")
    if st.button("⚡ この内容を転記する", key="btn2", use_container_width=True):
        if os.path.exists("Final_IT.pdf"): selected_file = "Final_IT.pdf"
        
    st.text_area("📄 請求書テキスト（Final_IT.pdf）", value=get_pdf_preview_text("Final_IT.pdf"), height=300, disabled=True)

with btn_col3:
    st.write("### 🍳 小野寺企画・飲食事業部")
    if st.button("⚡ この内容を転記する", key="btn3", use_container_width=True):
        if os.path.exists("Final_Mixed.pdf"): selected_file = "Final_Mixed.pdf"
        
    st.text_area("📄 請求書テキスト（Final_Mixed.pdf）", value=get_pdf_preview_text("Final_Mixed.pdf"), height=300, disabled=True)

# ボタンが押されたら解析して画面リフレッシュ
if selected_file is not None:
    st.session_state.current_pdf = selected_file
    with open(selected_file, "rb") as f:
        run_analysis(f)
    st.rerun()

if st.session_state.current_pdf:
    st.success(f"🎉 `{st.session_state.current_pdf}` の自動転記が正常に完了しました！")
