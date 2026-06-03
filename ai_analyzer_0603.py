import streamlit as st
from google import genai
from google.genai import types
import pandas as pd
import json
import re
import time
from pptx import Presentation

# --- 1. アプリ設定 & UI ---
st.set_page_config(page_title="AI Analyzer", layout="wide", initial_sidebar_state="expanded")

# サイドバーのみを劇的に圧縮するカスタムCSS（メインエリアは標準の余白）
st.html("""
    <style>
        /* サイドバー上部の見えない空白を完全に消去 */
        [data-testid="stSidebarHeader"] {
            display: none !important;
            padding: 0 !important;
            height: 0 !important;
        }
        /* サイドバー内の余白を最小化 */
        [data-testid="stSidebarUserContent"] {
            padding-top: 0.5rem !important;
        }
        /* 入力部品の間隔を詰めてファーストビューに収める */
        [data-testid="stSidebar"] .stTextInput, 
        [data-testid="stSidebar"] .stFileUploader {
            margin-bottom: -10px !important;
        }
    </style>
""")

# メインタイトル
st.html("""
    <div style="background-color:#f8f9fa; padding:20px; border-radius:10px; border-left:8px solid #007bff; margin-bottom:25px;">
        <h2 style="margin:0; color:#1e293b;">📊 AI Analyzer Pro</h2>
        <p style="margin:5px 0 0 0; color:#64748b; font-size:15px;">公式戦略と生成AIのギャップをAIが自動解析し、次の一手を導き出すマーケティングダッシュボード</p>
    </div>
""")

# セッション状態の管理
if "bas_result" not in st.session_state:
    st.session_state.bas_result = None

# サイドバー：設定エリア（圧縮版）
with st.sidebar:
    st.markdown("### ⚙️ 解析設定")
    api_key = st.text_input("🔑 Gemini API Key", type="password", help="Google AI Studioで発行したAPIキー")
    brand_name = st.text_input("🏷️ ブランド名", value="ディアナチュラ")
    brand_url = st.text_input("🌐 公式サイトURL", value="https://www.dear-natura.com/")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📂 データのアップロード")
    uploaded_pptx = st.file_uploader("❶ 生成AI分析資料（PPTX）", type=["pptx"])
    uploaded_csv = st.file_uploader("❷ 生成AI言及データ(CSV)", type=["csv", "txt"])
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⏹️ 画面をリセット", type="secondary", use_container_width=True):
        st.session_state.bas_result = None
        st.rerun()
        
    debug_mode = st.checkbox("🔧 開発用ログ表示", value=False)


# --- ユーティリティ関数 ---
def extract_text_from_pptx(file) -> str:
    try:
        prs = Presentation(file)
        text_runs = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text_runs.append(shape.text)
        return "\n".join(text_runs)[:3000]
    except Exception as e:
        return f"PPTX抽出エラー: {e}"

def load_csv_data(file):
    encodings = ["utf-8", "shift_jis", "cp932", "utf-8-sig", "euc-jp", "iso-8859-1"]
    for encoding in encodings:
        try:
            file.seek(0)
            df = pd.read_csv(file, encoding=encoding, sep=None, engine="python", on_bad_lines="skip")
            if not df.empty:
                return df
        except Exception:
            continue
    return None

def generate_with_retry(client, model_name, prompt, schema, phase_name):
    """制限に当たった際に裏側で静かに自動待機して突破する通信関数"""
    for attempt in range(3):
        try:
            res = client.models.generate_content(
                model=model_name,
                contents=[prompt],
                config=types.GenerateContentConfig(
                    temperature=0.1, 
                    response_mime_type="application/json", 
                    response_schema=schema
                )
            )
            return json.loads(res.text)
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                match = re.search(r"retry in ([\d\.]+)s", err_msg)
                if match and attempt < 2:
                    wait_sec = int(float(match.group(1))) + 2
                    time.sleep(wait_sec)
                    continue
                else:
                    st.error("❌ APIの1日上限に達しました。明日までお待ちいただくか、有料APIキーへ切り替えてください。")
                    st.stop()
            elif attempt < 2:
                time.sleep(5)
                continue
            else:
                st.error(f"❌ 分析中にエラーが発生しました: {e}")
                st.stop()
    return None


# --- 2. 分析ロジック ---
if st.button("🚀 戦略ギャップ分析を実行", type="primary", use_container_width=True):
    if not api_key or not uploaded_pptx or not uploaded_csv:
        st.error("左側のサイドバーで、APIキーと2つのファイルをセットしてください。")
        st.stop()

    with st.spinner("AIが戦略とデータのギャップを深掘り解析中..."):
        try:
            client = genai.Client(api_key=api_key)
            model_name = "gemini-2.5-flash"
            
            pptx_text = extract_text_from_pptx(uploaded_pptx)
            df_raw = load_csv_data(uploaded_csv)
            
            if df_raw is None:
                st.error("CSVファイルの読み込みに失敗しました。")
                st.stop()
            
            # Phase 1: 理想のキーワード抽出
            prompt_dict = f"ブランド '{brand_name}' の戦略資料から理想のブランド像キーワードを5つずつ抽出せよ:\n{pptx_text}"
            dict_schema = {
                "type": "object",
                "properties": {
                    "core": {"type":
