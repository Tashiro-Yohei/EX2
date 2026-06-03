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
                                # ★ 4カラム構成へアップデート（reasonの追加）
                                "official_content": {"type": "string"},
                                "ai_content": {"type": "string"},
                                "category": {
                                    "type": "string",
                                    "enum": ["一致", "乖離（ポジティブ）", "乖離（ネガティブ）"]
                                },
                                "reason": {"type": "string"},
                                "score": {"type": "integer"}
                            },
                            "required": ["official_content", "ai_content", "category", "reason", "score"]
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
            3. 詳細エビデンス: 上位20〜25のトピックについて、「公式発信」と「生成AI発信」を対比させて比較せよ。
               - official_content: 公式の戦略意図や発信内容（要点のみを非常に簡潔に）
               - ai_content: 生成AIでの実際の言及内容（要点のみを非常に簡潔に）
               - category: "一致", "乖離（ポジティブ）", "乖離（ネガティブ）" のいずれかに必ず分類せよ。
               - reason: その判定に至った理由（なぜ一致か、なぜポジ/ネガな乖離か）を端的に簡潔に記載せよ。
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
        df = pd.DataFrame(columns=['official_content', 'ai_content', 'category', 'reason', 'score'])
    
    if 'score' in df.columns:
        df['score'] = pd.to_numeric(df['score'], errors='coerce').fillna(0).astype(int)
    else:
        df['score'] = 0

    if debug_mode and res:
        with st.expander("🔧 デバッグデータ"):
            st.json(res)

    # === 【UI/PDF共用】文字が途切れないHTMLテーブルの組み立て ===
    table_rows_html = ""
    for _, row in df.iterrows():
        cat = row.get('category', '')
        row_cls = "row-match" if cat == "一致" else "row-pos" if "ポジティブ" in cat else "row-neg"
        table_rows_html += f"""
        <tr class="{row_cls}">
            <td style="border-bottom: 1px solid #e2e8f0; padding: 12px; font-size: 13px; word-wrap: break-word;">{row.get('official_content', '')}</td>
            <td style="border-bottom: 1px solid #e2e8f0; padding: 12px; font-size: 13px; word-wrap: break-word;">{row.get('ai_content', '')}</td>
            <td style="font-weight: bold; border-bottom: 1px solid #e2e8f0; padding: 12px; font-size: 13px; text-align: center;">{cat}</td>
            <td style="border-bottom: 1px solid #e2e8f0; padding: 12px; font-size: 13px; word-wrap: break-word;">{row.get('reason', '')}</td>
        </tr>
        """

    # ① 現状診断
    story = res.get("diagnosis_story", {})
    st.markdown("### 🎯 ① 現状診断：イメージの一致と乖離")
    st.caption("公式のブランド戦略が生成AIの回答にどう反映されているか、一致点と乖離の物語を300文字ずつのストーリーで紐解きます。")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.html("<div style='background-color:#e6f4ea; padding:12px; border-radius:5px; border-left:5px solid #28a745; font-weight:bold; color:#137333; margin-bottom:10px;'>🟢 一致（狙い通り）</div>")
        st.write(story.get("match", "データなし"))
    with col2:
        st.html("<div style='background-color:#e8f0fe; padding:12px; border-radius:5px; border-left:5px solid #007bff; font-weight:bold; color:#1a73e8; margin-bottom:10px;'>🔵 乖離：ポジティブ（新発見）</div>")
        st.write(story.get("positive_gap", "データなし"))
    with col3:
        st.html("<div style='background-color:#fce8e6; padding:12px; border-radius:5px; border-left:5px solid #dc3545; font-weight:bold; color:#c5221f; margin-bottom:10px;'>🔴 乖離：ネガティブ（課題）</div>")
        st.write(story.get("negative_gap", "データなし"))
            
    st.divider()

    # ② 割合のグラフ
    st.markdown("### 📊 ② 一致・乖離（ポジティブ）・乖離（ネガティブ）の割合")
    
    m_val = df[df['category'] == '一致']['score'].sum() if not df.empty else 0
    p_val = df[df['category'] == '乖離（ポジティブ）']['score'].sum() if not df.empty else 0
    n_val = df[df['category'] == '乖離（ネガティブ）']['score'].sum() if not df.empty else 0
    total_v = m_val + p_val + n_val

    m_p = (m_val / total_v * 100) if total_v > 0 else 0
    p_p = (p_val / total_v * 100) if total_v > 0 else 0
    n_p = (n_val / total_v * 100) if total_v > 0 else 0

    p1, p2, p3 = m_p, m_p + p_p, 100.0
    col_chart, col_metric = st.columns([1, 1])
    with col_chart:
        st.html(f"""
        <div style="display: flex; justify-content: center; align-items: center; height: 220px;">
            <div style="width: 180px; height: 180px; border-radius: 50%; background: conic-gradient(#28a745 0% {p1}%, #007bff {p1}% {p2}%, #dc3545 {p2}% {p3}%); display: flex; justify-content: center; align-items: center;">
                <div style="width: 110px; height: 110px; border-radius: 50%; background-color: white; display: flex; flex-direction: column; justify-content: center; align-items: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                    <span style="font-size: 12px; color: #64748b; font-weight: bold;">シンクロ率</span>
                    <span style="font-size: 24px; color: #1e293b; font-weight: 800;">{res.get('consistency_score', 0)}%</span>
                </div>
            </div>
        </div>
        """)
    with col_metric:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"🟢 :green[**一致（狙い通り）**]: `{m_p:.1f}%`")
        st.markdown(f"🔵 :blue[**乖離（ポジティブ）**]: `{p_p:.1f}%`")
        st.markdown(f"🔴 :red[**乖離（ネガティブ）**]: `{n_p:.1f}%`")

    st.divider()

    # ③ 今後取り組むべきこと
    st.markdown("### 🚀 ③ 今後取り組むべきこと")
    st.info(f"**【最優先戦略方針】** {res.get('topline')}")
    for i, action in enumerate(res.get("improvement_actions", []), 1):
        st.markdown(f"**{i}.** {action}")
        
    st.divider()

    # ④ 詳細情報（カスタムHTML表を埋め込み：文字折り返し完全対応）
    st.markdown("### 📖 ④ 詳細情報（判定エビデンス対比表）")
    st.caption("公式の戦略意図と実際の生成AIの出力を横並びで比較し、その乖離理由を可視化しています。")
    
    ui_table_html = f"""
    <style>
        .custom-evidence-table {{ width: 100%; border-collapse: collapse; font-family: sans-serif; }}
        .custom-evidence-table th {{ background-color: #f1f5f9; padding: 12px; text-align: left; border-bottom: 2px solid #cbd5e1; position: sticky; top: 0; font-size: 14px; color: #1e293b; }}
        .custom-evidence-table td {{ color: #334155; }}
        .row-match {{ background-color: #f2faf4; }}
        .row-pos {{ background-color: #f4f8ff; }}
        .row-neg {{ background-color: #fff5f5; }}
    </style>
    <div style="max-height: 600px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 8px; margin-top: 15px;">
        <table class="custom-evidence-table">
            <thead>
                <tr>
                    <th style="width: 25%;">公式発信の内容</th>
                    <th style="width: 25%;">生成AI発信の内容</th>
                    <th style="width: 15%; text-align: center;">判定分類</th>
                    <th style="width: 35%;">判定の理由</th>
                </tr>
            </thead>
            <tbody>
                {table_rows_html}
            </tbody>
        </table>
    </div>
    """
    st.html(ui_table_html)


    # === 📄 レポート印刷・PDF保存エリア ===
    st.divider()
    st.markdown("### 📥 レポートの出力（PDF保存）")
    st.caption("表示されている解析結果を、崩れのない美しいA4レイアウトのレポートとして出力します。")
    
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

        <h2>🎯 ① 現状診断：イメージの一致と乖離</h2>
        <div class="grid">
            <div class="col match" style="margin-right: 10px;">
                <div style="font-weight: bold; margin-bottom: 8px;">🟢 一致（狙い通り）</div>
                <div style="font-size: 13px;">{story.get('match', '')}</div>
            </div>
            <div class="col pos" style="margin-right: 10px; margin-left: 10px;">
                <div style="font-weight: bold; margin-bottom: 8px;">🔵 乖離：ポジティブ（新発見）</div>
                <div style="font-size: 13px;">{story.get('positive_gap', '')}</div>
            </div>
            <div class="col neg" style="margin-left: 10px;">
                <div style="font-weight: bold; margin-bottom: 8px;">🔴 乖離：ネガティブ（課題）</div>
                <div style="font-size: 13px;">{story.get('negative_gap', '')}</div>
            </div>
        </div>

        <h2>🚀 ② 今後取り組むべきこと</h2>
        <div class="info-box">【最優先戦略方針】 {res.get('topline', '')}</div>
        <ol style="padding-left: 20px; margin-top: 10px;">
            {actions_html}
        </ol>

        <h2>📖 ③ 詳細情報（判定エビデンス対比表）</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">公式発信の内容</th>
                    <th style="width: 25%;">生成AI発信の内容</th>
                    <th style="width: 15%; text-align: center;">判定分類</th>
                    <th style="width: 35%;">判定の理由</th>
                </tr>
            </thead>
            <tbody>
                {table_rows_html}
            </tbody>
        </table>
    </body>
    </html>
    """
    
    st.download_button(
        label="📄 レポート（印刷用高画質HTML）をダウンロード",
        data=html_report,
        file_name=f"AI_Strategy_Scan_Report_{brand_name}.html",
        mime="text/html",
        use_container_width=True
    )
    st.info("💡 **【PDF保存の方法】**\nダウンロードしたファイルをダブルクリックしてブラウザで開き、キーボードの「**Ctrl + P**（Macは **Cmd + P**）」を押して、送信先を『**PDFに保存**』にするだけで、きれいにレイアウトされたA4サイズのPDFレポートが作成されます。")

    with st.expander("📄 参考：生成AI分析資料から抽出された理想キーワード一覧"):
        d = res.get("dictionary", {})
        c1, c2, c3 = st.columns(3)
        c1.write("**コア価値**\n" + ", ".join(d.get("core", [])))
        c2.write("**機能・効能**\n" + ", ".join(d.get("functional", [])))
        c3.write("**専門性**\n" + ", ".join(d.get("professional", [])))

else:
    st.html("""
        <div style="text-align:center; padding:100px 20px; color:#94a3b8;">
            <p style="font-size:40px; margin:0;">📥</p>
            <h4 style="margin:10px 0 0 0; color:#64748b;">データがセットされていません</h4>
            <p style="font-size:14px; margin:5px 0 0 0;">左側のサイドバーに必要な情報をセットして「分析を実行」を押してください。</p>
        </div>
    """)
