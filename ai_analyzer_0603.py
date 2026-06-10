import sys
import streamlit as st
from google import genai
from google.genai import types
import pandas as pd
import json
import re
import time
from pptx import Presentation
import docx
import plotly.graph_objects as go

# --- 1. アプリ設定 & UI ---
st.set_page_config(page_title="AI Analyzer", layout="wide", initial_sidebar_state="expanded")

# 画面の余白を最適化するカスタムCSS
st.html("""
    <style>
        [data-testid="stSidebarHeader"] {
            padding-top: 0.5rem !important;
            padding-bottom: 0rem !important;
            min-height: auto !important;
        }
        [data-testid="stSidebarUserContent"] {
            padding-top: 0rem !important;
        }
        .block-container {
            padding-top: 3.5rem !important; 
            padding-bottom: 2rem !important;
        }
    </style>
""")

# タイトルエリア
st.html("""
    <div style="background-color:#f8f9fa; padding:20px; border-radius:10px; border-left:8px solid #007bff; margin-bottom:25px;">
        <h2 style="margin:0; color:#1e293b;">📊 AI Analyzer Pro</h2>
        <p style="margin:5px 0 0 0; color:#64748b; font-size:15px;">自社のブランド戦略が生成AIにどこまで正しく伝わっているかを分析し、次の一手を見つけるダッシュボード</p>
    </div>
""")

# セッション状態の初期化
if "bas_result" not in st.session_state:
    st.session_state.bas_result = None

# サイドバー：設定エリア
with st.sidebar:
    st.markdown("### ⚙️ 解析設定")
    st.divider()
    
    api_key = st.text_input("🔑 Gemini API Key", type="password", help="Google AI Studioで発行したAPIキーを入力してください。")
    brand_name = st.text_input("🏷️ ブランド名", value="ディアナチュラ")
    brand_url = st.text_input("🌐 公式サイトURL", value="https://www.dear-natura.com/")
    
    st.divider()
    st.markdown("### 📂 データのアップロード")
    uploaded_pptx = st.file_uploader("❶ 生成AI分析資料（PPTX）", type=["pptx"])
    uploaded_csv = st.file_uploader("❷ 生成AIでの言及数データ(CSV)", type=["csv", "txt"])
    uploaded_docx = st.file_uploader("❸ 生成AIのブランド評価(DOCX)", type=["docx"])
    
    st.divider()
    if st.button("⏹️ 処理を中断 / 画面をリセット", type="secondary", use_container_width=True):
        st.session_state.bas_result = None
        st.rerun()
        
    st.divider()
    debug_mode = st.checkbox("🔧 開発者用ログを表示する", value=False)


# --- ユーティリティ関数 ---
def clean_and_parse_json(text: str) -> dict:
    json_match = re.search(r'\{.*\}', text, re.DOTALL)
    if not json_match:
        raise ValueError("JSONが見つかりません")
    text = json_match.group(0)
    text = re.sub(r'[\x00-\x1F\x7F]', '', text)
    return json.loads(text)

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
        return f"PPTXエラー: {e}"

def extract_text_from_docx(file) -> str:
    try:
        doc = docx.Document(file)
        text = [paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()]
        return "\n".join(text)[:3000]
    except Exception as e:
        return f"DOCXエラー: {e}"

def load_csv_data(file):
    encodings = ["utf-8", "shift_jis", "cp932", "utf-8-sig", "iso-8859-1"]
    delimiters = [",", "\t", ";"]
    for encoding in encodings:
        for delimiter in delimiters:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=encoding, sep=delimiter, engine="python", on_bad_lines="skip")
                if not df.empty and len(df.columns) >= 2:
                    return df
            except Exception:
                continue
    return None


# --- 2. 分析ロジック ---
if st.button("🚀 戦略ギャップ分析を実行", type="primary", use_container_width=True):
    if not api_key or not brand_url or not uploaded_pptx or not uploaded_csv or not uploaded_docx:
        st.error("左側のサイドバーで、APIキー、URL、および3つのファイルをすべてセットしてください。")
        st.stop()

    with st.spinner("AIがデータを解析し、経営・マーケティング向けのレポートを作成中..."):
        try:
            client = genai.Client(api_key=api_key)
            model_name = "gemini-2.5-flash"
            
            pptx_text = extract_text_from_pptx(uploaded_pptx)
            df_raw = load_csv_data(uploaded_csv)
            docx_text = extract_text_from_docx(uploaded_docx)
            
            if df_raw is None:
                st.error("CSVファイルの読み込みに失敗しました。")
                st.stop()
                
            csv_context = df_raw.head(35).to_csv(index=False)
            
            # Phase 1: オウンドメディアからのキーワード抽出
            prompt_dict = f"""
            ブランド "{brand_name}" (公式URL: {brand_url}) の戦略資料からオウンドメディアが目指すブランド像を示すキーワードを5つずつ抽出してください。
            
            [戦略資料テキスト]
            {pptx_text}
            """
            
            dict_schema = {
                "type": "object",
                "properties": {
                    "core": {"type": "array", "items": {"type": "string"}},
                    "functional": {"type": "array", "items": {"type": "string"}},
                    "professional": {"type": "array", "items": {"type": "string"}}
                },
                "required": ["core", "functional", "professional"]
            }

            dictionary_data = None
            for attempt in range(3):
                try:
                    res_dict = client.models.generate_content(
                        model=model_name,
                        contents=[prompt_dict],
                        config=types.GenerateContentConfig(temperature=0.1, response_mime_type="application/json", response_schema=dict_schema)
                    )
                    dictionary_data = json.loads(res_dict.text)
                    break
                except Exception as e:
                    if "503" in str(e) and attempt < 2:
                        time.sleep(3)
                        continue
                    raise e

            if not dictionary_data:
                st.error("AIサーバーが混雑しています。少し時間を置いて再度お試しください。")
                st.stop()

            # Phase 2: 戦略的ギャップ分析とスコアリング
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
                    "radar_scores": {
                        "type": "object",
                        "properties": {
                            "brand_philosophy": {"type": "integer"},
                            "functional_value": {"type": "integer"},
                            "emotional_engagement": {"type": "integer"},
                            "safety_reputation": {"type": "integer"},
                            "competitive_priority": {"type": "integer"}
                        },
                        "required": ["brand_philosophy", "functional_value", "emotional_engagement", "safety_reputation", "competitive_priority"]
                    },
                    "radar_reasons": {
                        "type": "object",
                        "properties": {
                            "brand_philosophy": {"type": "string"},
                            "functional_value": {"type": "string"},
                            "emotional_engagement": {"type": "string"},
                            "safety_reputation": {"type": "string"},
                            "competitive_priority": {"type": "string"}
                        },
                        "required": ["brand_philosophy", "functional_value", "emotional_engagement", "safety_reputation", "competitive_priority"]
                    }
                },
                "required": ["diagnosis_story", "topline", "improvement_actions", "radar_scores", "radar_reasons"]
            }

            prompt_analysis = f"""
            Analyze Generative AI output data for "{brand_name}" (Official URL: {brand_url}) against the owned media keywords.
            
            [OWNED MEDIA KEYWORDS]
            {json.dumps(dictionary_data, ensure_ascii=False)}
            
            [GENERATIVE AI RANKING DATA (CSV)]
            {csv_context}
            
            [GENERATIVE AI BRAND EVALUATION (DOCX)]
            {docx_text}
            
            TASK:
            1. "diagnosis_story": Write 3 fluent narrative paragraphs in Japanese (EACH strictly around 200-250 characters) aimed at business executives. Use clear, accessible language.
               - "match": What aspects of the company's message are correctly understood by the AI?
               - "positive_gap": What unexpected strengths or positive perceptions did the AI find that the company isn't heavily promoting?
               - "negative_gap": What misconceptions or weak points exist in the AI's understanding compared to the company's intent?
            2. "topline": Write a single-sentence summary strategy for executives.
            3. "improvement_actions": Provide EXACTLY 5 clear, actionable marketing steps to turn weaknesses into strengths or to leverage unexpected positive perceptions.
            4. "radar_scores" & "radar_reasons": Score the Generative AI's perception on a scale of 0-100 for the following 5 criteria. Then, in "radar_reasons", provide a specific business reason (approx. 80-100 characters in Japanese) explaining WHY it got this score based on the data.
               - "brand_philosophy": ブランド理念の浸透度
               - "functional_value": 機能価値の伝達度
               - "emotional_engagement": 情緒的エンゲージメント（顧客の愛着度）
               - "safety_reputation": ブランドの安全性と評判
               - "competitive_priority": 対競合優先度（他社と比べた選ばれやすさ）
            Return JSON in Japanese.
            """

            final_data = None
            for attempt in range(3):
                try:
                    res_analysis = client.models.generate_content(
                        model=model_name,
                        contents=[prompt_analysis],
                        config=types.GenerateContentConfig(temperature=0.1, response_mime_type="application/json", response_schema=response_schema)
                    )
                    final_data = json.loads(res_analysis.text)
                    break
                except Exception as e:
                    if "503" in str(e) and attempt < 2:
                        time.sleep(3)
                        continue
                    raise e

            if final_data:
                final_data["dictionary"] = dictionary_data
                st.session_state.bas_result = final_data
            else:
                st.error("AIサーバーが混雑しています。少し時間を置いて再度お試しください。")
                st.stop()

        except Exception as e:
            st.error(f"分析エラーが発生しました: {e}")
            st.stop()


# --- 3. 結果表示 UI ---
if st.session_state.bas_result:
    res = st.session_state.bas_result

    if debug_mode and res:
        with st.expander("🔧 開発者用デバッグログ"):
            st.json(res)

    # ==========================================
    # ① 現状診断：自社発信とAIの認識ギャップ
    # ==========================================
    st.markdown("### 🎯 ① 現状の診断：自社発信とAIの認識ギャップ")
    st.caption("自社で発信しているメッセージがAIにどう伝わっているか、「狙い通りな点」と「意外なズレ」をわかりやすく解説します。")
    
    story = res.get("diagnosis_story", {})
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.html("<div style='background-color:#e6f4ea; padding:12px; border-radius:5px; border-left:5px solid #28a745; font-weight:bold; color:#137333; margin-bottom:10px;'>🟢 狙い通りに伝わっている強み</div>")
        st.write(story.get("match", "分析データなし"))
        
    with col2:
        st.html("<div style='background-color:#e8f0fe; padding:12px; border-radius:5px; border-left:5px solid #007bff; font-weight:bold; color:#1a73e8; margin-bottom:10px;'>🔵 AIが見つけた「意外な強み」</div>")
        st.write(story.get("positive_gap", "分析データなし"))
        
    with col3:
        st.html("<div style='background-color:#fce8e6; padding:12px; border-radius:5px; border-left:5px solid #dc3545; font-weight:bold; color:#c5221f; margin-bottom:10px;'>🔴 AIの誤解や情報不足（課題）</div>")
        st.write(story.get("negative_gap", "分析データなし"))
            
    st.divider()

    # ==========================================
    # ② 生成AIからのブランド評価（5つの重要指標）
    # ==========================================
    st.markdown("### 📊 ② 生成AIからのブランド評価（5つの重要指標）")
    st.caption("自社のブランド戦略がAIにどの程度評価され、浸透しているかを5つの軸で採点した結果とその理由です。")
    
    radar = res.get("radar_scores", {})
    reasons = res.get("radar_reasons", {})
    
    s_bp = radar.get("brand_philosophy", 0)
    s_fv = radar.get("functional_value", 0)
    s_ee = radar.get("emotional_engagement", 0)
    s_sr = radar.get("safety_reputation", 0)
    s_cp = radar.get("competitive_priority", 0)
    
    # 総合スコア（平均）
    overall = sum([s_bp, s_fv, s_ee, s_sr, s_cp]) / 5.0
    
    def get_color_style(score):
        if score >= 75:
            return "background-color:#d4edda; border:1px solid #c3e6cb; color:#155724;" # 緑
        elif score >= 60:
            return "background-color:#fff3cd; border:1px solid #ffeeba; color:#856404;" # 黄
        else:
            return "background-color:#f8d7da; border:1px solid #f5c6cb; color:#721c24;" # 赤

    col_radar, col_metrics = st.columns([1, 1])
    
    with col_radar:
        categories = ['ブランド理念の<br>浸透度', '機能価値の<br>伝達度', '情緒的<br>エンゲージメント', 'ブランドの<br>安全性と評判', '対競合優先度']
        scores_list = [s_bp, s_fv, s_ee, s_sr, s_cp]
        
        categories_closed = categories + [categories[0]]
        scores_closed = scores_list + [scores_list[0]]
        
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(
            r=scores_closed,
            theta=categories_closed,
            fill='toself',
            name='Score',
            line_color='#c55a11',
            fillcolor='rgba(197, 90, 17, 0.2)'
        ))
        
        fig.update_layout(
            polar=dict(
                radialaxis=dict(visible=True, range=[0, 100], showticklabels=False),
            ),
            showlegend=False,
            margin=dict(l=60, r=60, t=40, b=40),
            height=380
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_metrics:
        html_cards = f"""
        <div style="margin-bottom:15px; padding:15px; border-radius:5px; {get_color_style(overall)}">
            <div style="font-size:12px; margin-bottom:5px;">総合スコア</div>
            <div style="font-size:32px; font-weight:bold;">{overall:.1f}</div>
        </div>
        
        <div style="display:flex; gap:10px; margin-bottom:10px;">
            <div style="flex:1; padding:15px; border-radius:5px; {get_color_style(s_bp)}">
                <div style="font-size:12px; margin-bottom:5px;">ブランド理念の浸透度</div>
                <div style="font-size:24px; font-weight:bold;">{s_bp:.1f}</div>
            </div>
            <div style="flex:1; padding:15px; border-radius:5px; {get_color_style(s_fv)}">
                <div style="font-size:12px; margin-bottom:5px;">機能価値の伝達度</div>
                <div style="font-size:24px; font-weight:bold;">{s_fv:.1f}</div>
            </div>
        </div>
        
        <div style="display:flex; gap:10px; margin-bottom:10px;">
            <div style="flex:1; padding:15px; border-radius:5px; {get_color_style(s_ee)}">
                <div style="font-size:12px; margin-bottom:5px;">情緒的エンゲージメント</div>
                <div style="font-size:24px; font-weight:bold;">{s_ee:.1f}</div>
            </div>
            <div style="flex:1; padding:15px; border-radius:5px; {get_color_style(s_sr)}">
                <div style="font-size:12px; margin-bottom:5px;">ブランドの安全性と評判</div>
                <div style="font-size:24px; font-weight:bold;">{s_sr:.1f}</div>
            </div>
        </div>
        
        <div style="display:flex; gap:10px;">
            <div style="flex:0.49; padding:15px; border-radius:5px; {get_color_style(s_cp)}">
                <div style="font-size:12px; margin-bottom:5px;">対競合優先度</div>
                <div style="font-size:24px; font-weight:bold;">{s_cp:.1f}</div>
            </div>
            <div style="flex:0.51;"></div>
        </div>
        
        <div style="margin-top:10px; font-size:11px; color:#6c757d;">
            色分け: 75点以上は緑（優秀）、60-74点は黄（標準）、59点以下は赤（改善が必要）を示します。
        </div>
        """
        st.html(html_cards)

    # 評価理由の表示エリア
    st.markdown("#### 📝 各スコアの評価理由（なぜこの点数になったのか）")
    st.markdown(f"""
    * **ブランド理念の浸透度 ({s_bp}点):** {reasons.get('brand_philosophy', 'データなし')}
    * **機能価値の伝達度 ({s_fv}点):** {reasons.get('functional_value', 'データなし')}
    * **情緒的エンゲージメント ({s_ee}点):** {reasons.get('emotional_engagement', 'データなし')}
    * **ブランドの安全性と評判 ({s_sr}点):** {reasons.get('safety_reputation', 'データなし')}
    * **対競合優先度 ({s_cp}点):** {reasons.get('competitive_priority', 'データなし')}
    """)

    st.divider()

    # ==========================================
    # ③ 今後、マーケティングで打つべき具体策
    # ==========================================
    st.markdown("### 🚀 ③ 今後、マーケティングで打つべき具体策")
    st.caption("AIの誤解を解き、意外な強みを自社のPRに活かすための5つのアクションプランです。")
    
    st.info(f"**【戦略方針】** {res.get('topline')}")
    
    st.markdown("#### 🛠️ 具体的な5つのアクション")
    for i, action in enumerate(res.get("improvement_actions", []), 1):
        st.markdown(f"**{i}.** {action}")
        
    st.divider()

    # 参考情報
    with st.expander("📄 参考：AIが分析のベースにした自社の公式キーワード"):
        d = res.get("dictionary", {})
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**💡 コアとなる価値**")
            st.write(", ".join(d.get("core", [])) if d.get("core") else "抽出なし")
        with c2:
            st.markdown("**✨ 機能や効能**")
            st.write(", ".join(d.get("functional", [])) if d.get("functional") else "抽出なし")
        with c3:
            st.markdown("**🛡️ 信頼感や専門性**")
            st.write(", ".join(d.get("professional", [])) if d.get("professional") else "抽出なし")

else:
    st.html("""
        <div style="text-align:center; padding:100px 20px; color:#94a3b8;">
            <p style="font-size:40px; margin:0;">📥</p>
            <h4 style="margin:10px 0 0 0; color:#64748b;">データがセットされていません</h4>
            <p style="font-size:14px; margin:5px 0 0 0;">左側のサイドバーにAPIキーを入力し、3つのファイルをセットして「分析を開始する」を押してください。</p>
        </div>
    """)
