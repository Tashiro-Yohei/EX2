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
    uploaded_csv = st.file_uploader("❷ SNS言及データ(CSV)", type=["csv", "txt"])
    
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

# ★列数の制限を完全に撤廃し、あらゆる形式のCSV/TXTを確実に読み込む最強ロジック
def load_csv_data(file):
    encodings = ["utf-8", "shift_jis", "cp932", "utf-8-sig", "euc-jp", "iso-8859-1"]
    for encoding in encodings:
        try:
            file.seek(0)
            # sep=None, engine='python' により、カンマ・タブ・スペース等の区切り文字を自動判別
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
            # 短時間制限（429）の場合、画面には出さずに裏側で静かに待機する
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                match = re.search(r"retry in ([\d\.]+)s", err_msg)
                if match and attempt < 2:
                    wait_sec = int(float(match.group(1))) + 2
                    time.sleep(wait_sec)
                    continue
                else:
                    st.error("❌ APIの1日上限に達しました。明日までお待ちいただくか、有料APIキーへ切り替えてください。")
                    st.stop()
            # 混雑エラー（503）などの場合
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
                                "word": {"type": "string"},
                                "score": {"type": "integer"},
                                "category": {
                                    "type": "string",
                                    "enum": ["一致", "乖離（ポジティブ）", "乖離（ネガティブ）", "その他"]
                                },
                                "reason": {"type": "string"}
                            },
                            "required": ["word", "score", "category", "reason"]
                        }
                    },
                    "consistency_score": {"type": "integer"}
                },
                "required": ["diagnosis_story", "topline", "improvement_actions", "ranking_data", "consistency_score"]
            }

            prompt_analysis = f"""
            理想像: {json.dumps(dictionary_data, ensure_ascii=False)}
            実際のSNSデータ: {csv_context}
            
            [指示]
            1. 診断ストーリー: 3つのカテゴリ(一致/ポジティブ乖離/ネガティブ乖離)について、各300文字程度の流れるような文章で記述せよ。項目分けや箇条書きは禁止。
            2. 改善提言: ネガティブをポジティブに変え、ポジティブを公式化するための戦略方針を1行で、具体的アクションを5つ作成せよ。
            3. 詳細エビデンス: 上位20〜25語について、判定理由をロジカルに記述せよ。なぜ一致/乖離なのか、戦略のどの部分に関連するかを明記せよ。
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
        df = pd.DataFrame(columns=['word', 'score', 'category', 'reason'])
    
    if 'score' in df.columns:
        df['score'] = pd.to_numeric(df['score'], errors='coerce').fillna(0).astype(int)
    else:
        df['score'] = 0

    if debug_mode and res:
        with st.expander("🔧 デバッグデータ"):
            st.json(res)

    # ① 現状診断（300文字ストーリー）
    st.markdown("### 🎯 ① 現状診断：イメージの一致と乖離")
    st.caption("公式のブランド戦略が生成AIの回答にどう反映されているか、一致点と乖離の物語を300文字ずつのストーリーで紐解きます。")
    
    story = res.get("diagnosis_story", {})
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
    st.markdown("### 📊 ② 生成AI言言の割合（ポートフォリオ）")
    total_v = df['score'].sum()
    m_val = df[df['category'] == '一致']['score'].sum() if not df.empty else 0
    p_val = df[df['category'] == '乖離（ポジティブ）']['score'].sum() if not df.empty else 0
    n_val = df[df['category'] == '乖離（ネガティブ）']['score'].sum() if not df.empty else 0

    m_p = (m_val / total_v * 100) if total_v > 0 else 0
    p_p = (p_val / total_v * 100) if total_v > 0 else 0
    n_p = (n_val / total_v * 100) if total_v > 0 else 0
    o_p = max(0.0, 100 - (m_p + p_p + n_p)) if total_v > 0 else 100

    # ドーナツグラフの描画
    p1, p2, p3 = m_p, m_p + p_p, m_p + p_p + n_p
    col_chart, col_metric = st.columns([1, 1])
    with col_chart:
        st.html(f"""
        <div style="display: flex; justify-content: center; align-items: center; height: 220px;">
            <div style="width: 180px; height: 180px; border-radius: 50%; background: conic-gradient(#28a745 0% {p1}%, #007bff {p1}% {p2}%, #dc3545 {p2}% {p3}%, #e2e8f0 {p3}% 100%); display: flex; justify-content: center; align-items: center;">
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
        st.markdown(f"⚪ :gray[**その他**]: `{o_p:.1f}%`")

    st.divider()

    # ③ 今後取り組むべきこと
    st.markdown("### 🚀 ③ 今後取り組むべきこと")
    st.info(f"**【最優先戦略方針】** {res.get('topline')}")
    for i, action in enumerate(res.get("improvement_actions", []), 1):
        st.markdown(f"**{i}.** {action}")
        
    st.divider()

    # ④ 詳細情報
    st.markdown("### 📖 ④ 詳細情報（判定エビデンス一覧）")
    df_table = df[['word', 'category', 'reason']].rename(columns={'word': '注目キーワード','category': '判定分類','reason': '判定の具体的な理由・背景'})
    
    def color_rows(row):
        cat = row.get('判定分類', '')
        if cat == '一致': return ['background-color: #e6f4ea'] * 3
        if cat == '乖離（ポジティブ）': return ['background-color: #e8f0fe'] * 3
        if cat == '乖離（ネガティブ）': return ['background-color: #fce8e6'] * 3
        return [''] * 3

    st.dataframe(df_table.style.apply(color_rows, axis=1), use_container_width=True, height=550, hide_index=True)

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
