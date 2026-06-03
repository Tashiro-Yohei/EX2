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

st.html("""
    <style>
        [data-testid="stSidebarHeader"] {
            display: none !important;
            padding: 0 !important;
            height: 0 !important;
        }
        [data-testid="stSidebarUserContent"] {
            padding-top: 0.5rem !important;
        }
        [data-testid="stSidebar"] .stTextInput, 
        [data-testid="stSidebar"] .stFileUploader {
            margin-bottom: -10px !important;
        }
    </style>
""")

st.html("""
    <div style="background-color:#f8f9fa; padding:20px; border-radius:10px; border-left:8px solid #007bff; margin-bottom:25px;">
        <h2 style="margin:0; color:#1e293b;">📊 AI Analyzer Pro</h2>
        <p style="margin:5px 0 0 0; color:#64748b; font-size:15px;">公式戦略と生成AIのギャップをAIが自動解析し、次の一手を導き出すマーケティングダッシュボード</p>
    </div>
""")

if "bas_result" not in st.session_state:
    st.session_state.bas_result = None

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
                    "core": {"type": "array", "items": {"type": "string"}},
                    "functional": {"type": "array", "items": {"type": "string"}},
                    "professional": {"type": "array", "items": {"type": "string"}}
                },
                "required": ["core", "functional", "professional"]
            }
            dictionary_data = generate_with_retry(client, model_name, prompt_dict, dict_schema, "キーワード抽出")
            
            if not dictionary_data:
                st.stop()
            
            # Phase 2: ストーリー構築とエビデンス表の生成
            csv_context = df_raw.head(35).to_csv(index=False)
            response_schema = {
                "type": "object",
                "properties": {
                    "diagnosis_story": {
                        "type": "object",
                        "properties": {
                            "match": {"type": "string"},
                            "positive_gap": {"type": "string"},
                            "negative_gap": {"type": "string"}
                        },
                        "required": ["match", "positive_gap", "negative_gap"]
                    },
                    "topline": {"type": "string"},
                    "improvement_actions": {"type": "array", "items": {"type": "string"}},
                    "ranking_data": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                # ★ 単語単位ではなく、言及内容（文脈）の対比に変更
                                "official_content": {"type": "string"},
                                "ai_content": {"type": "string"},
                                "score": {"type": "integer"},
                                "category": {
                                    "type": "string",
                                    "enum": ["一致", "乖離（ポジティブ）", "乖離（ネガティブ）"]
                                }
                            },
                            "required": ["official_content", "ai_content", "score", "category"]
                        }
                    },
                    "consistency_score": {"type": "integer"}
                },
                "required": ["diagnosis_story", "topline", "improvement_actions", "ranking_data", "consistency_score"]
            }

            prompt_analysis = f"""
            理想像: {json.dumps(dictionary_data, ensure_ascii=False)}
            実際の生成AIデータ: {csv_context}
            
            [指示]
            1. 診断ストーリー: 3つのカテゴリ(一致/ポジティブ乖離/ネガティブ乖離)について、各300文字程度の流れるような文章で記述せよ。項目分けや箇条書きは禁止。
            2. 改善提言: ネガティブをポジティブに変え、ポジティブを公式化するための戦略方針を1行で、具体的アクションを5つ作成せよ。
            3. 詳細エビデンス: 上位20〜25のトピックについて、「公式の言及内容」と「生成AIの言及内容」を対比させて比較せよ。
               - official_content: 公式の戦略・理想像では、そのトピックについてどのように言及・企図されているか（記載がない場合はその旨）。簡潔に。
               - ai_content: 実際の生成AIのデータでは、そのトピックがどのような内容・文脈で言及されているか。簡潔に。
               - category: "一致", "乖離（ポジティブ）", "乖離（ネガティブ）" のいずれかに必ず分類せよ。
            """
            
            final_data = generate_with_retry(client, model_name, prompt_analysis, response_schema, "総合ギャップ分析")
            
            if final_data:
                final_data["dictionary"] = dictionary_data
                st.session_state.bas_result = final_data

        except Exception as e:
            st.error(f"致命的なエラーが発生しました: {e}")
            st.stop()


# --- 3. 結果表示 UI ---
if st.session_state.bas_result:
    res = st.session_state.bas_result
    
    df = pd.DataFrame(res.get("ranking_data", []))
    if df.empty:
        df = pd.DataFrame(columns=['official_content', 'ai_content', 'score', 'category'])
    
    if 'score' in df.columns:
        df['score'] = pd.to_numeric(df['score'], errors='coerce').fillna(0).astype(int)
    else:
        df['score'] = 0

    if debug_mode and res:
        with st.expander("🔧 デバッグデータ"):
            st.json(res)

    # === 【PDFレポート出力用HTMLの生成】 ===
    story = res.get("diagnosis_story", {})
    
    # 印刷用テーブル行の組み立て（3列レイアウトに変更）
    table_rows_html = ""
    for _, row in df.iterrows():
        cat = row.get('category', '')
        row_cls = "row-match" if cat == "一致" else "row-pos" if "ポジティブ" in cat else "row-neg"
        table_rows_html += f"""
        <tr class="{row_cls}">
            <td style="border-bottom: 1px solid #e2e8f0; padding: 10px; font-size: 13px;">{row.get('official_content', '')}</td>
            <td style="border-bottom: 1px solid #e2e8f0; padding: 10px; font-size: 13px;">{row.get('ai_content', '')}</td>
            <td style="font-weight: bold; border-bottom: 1px solid #e2e8f0; padding: 10px; white-space: nowrap; font-size: 13px; text-align: center;">{cat}</td>
        </tr>
        """

    actions_html = "".join([f"<li style='margin-bottom: 8px;'>{action}</li>" for action in res.get("improvement_actions", [])])

    html_report = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>AI Strategy Scan Report</title>
        <style>
            body {{ font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif; color: #1e293b; line-height: 1.6; padding: 30px; max-width: 900px; margin: 0 auto; background-color: #ffffff; }}
            .header {{ background-color: #f8f9fa; padding: 20px; border-radius: 8px; border-left: 8px solid #007bff; margin-bottom: 30px; }}
            h1 {{ margin: 0; color: #1e293b; font-size: 24px; }}
            h2 {{ color: #1e293b; border-left: 5px solid #007bff; padding-left: 10px; font-size: 18px; margin-top: 35px; margin-bottom: 15px; page-break-after: avoid; }}
            .grid {{ display: table; width: 100%; table-layout: fixed; margin-bottom: 25px; }}
            .col {{ display: table-cell; width: 33.33%; padding: 15px; border-radius: 6px; box-sizing: border-box; vertical-align: top; }}
            .match {{ background-color: #e6f4ea; color: #137333; border: 1px solid #c3e6cb; }}
            .pos {{ background-color: #e8f0fe; color: #1a73e8; border: 1px solid #b8daff; }}
            .neg {{ background-color: #fce8e6; color: #c5221f; border: 1px solid #f5c6cb; }}
            .info-box {{ background-color: #e8f0fe; border-left: 5px solid #007bff; padding: 15px; border-radius: 4px; margin-bottom: 15px; font-weight: bold; color: #1a73e8; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th {{ background-color: #f1f5f9; font-weight: bold; padding: 12px 10px; border-bottom: 2px solid #cbd5e1; font-size: 14px; text-align: left; }}
            .row-match {{ background-color: #f2faf4; }}
            .row-pos {{ background-color: #f4f8ff; }}
            .row-neg {{ background-color: #fff5f5; }}
            @media print {{
                body {{ padding: 0; }}
                h2 {{ page-break-inside: avoid; }}
                tr {{ page-break-inside: avoid; }}
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>📊 AI Strategy Scan 分析結果レポート</h1>
            <p style="margin: 5px 0 0 0; color: #64748b; font-size: 14px;">対象ブランド: {brand_name} ({brand_url}) &nbsp;|&nbsp; 戦略シンクロ率: <strong>{res.get('consistency_score', 0)}%</strong></p>
        </div>
