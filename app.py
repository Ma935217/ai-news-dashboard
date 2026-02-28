"""
AIニュース収集ダッシュボード
- Google News RSS から feedparser でニュースを取得
- Streamlit でカード型デザインのダッシュボードを構築
- 介護福祉×AI ニュースに特化したプリセット検索ワード付き
"""

import streamlit as st
import feedparser
import urllib.parse
from datetime import datetime
import time
import re
import html


# ─────────────────────────────────────
# ページ設定
# ─────────────────────────────────────
st.set_page_config(
    page_title="AI News Dashboard",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ─────────────────────────────────────
# カスタム CSS
# ─────────────────────────────────────
st.markdown("""
<style>
/* ---------- Google Fonts ---------- */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Noto+Sans+JP:wght@300;400;500;600;700&display=swap');

/* ---------- Global ---------- */
html, body, [class*="css"] {
    font-family: 'Inter', 'Noto Sans JP', sans-serif;
}

.stApp {
    background: linear-gradient(135deg, #0f0c29 0%, #1a1a40 40%, #24243e 100%);
}

/* ---------- Sidebar ---------- */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a1a40 0%, #0f0c29 100%);
    border-right: 1px solid rgba(255, 255, 255, 0.06);
}

section[data-testid="stSidebar"] .stMarkdown h1,
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3 {
    color: #e0e0ff !important;
}

section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown label {
    color: #b0b0d0 !important;
}

/* ---------- Header ---------- */
.main-header {
    text-align: center;
    padding: 2rem 1rem 1.5rem;
    margin-bottom: 1rem;
}

.main-header h1 {
    font-size: 2.6rem;
    font-weight: 800;
    background: linear-gradient(135deg, #00d2ff 0%, #7b68ee 50%, #ff6ec7 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 0.3rem;
    letter-spacing: -0.02em;
}

.main-header .subtitle {
    font-size: 1rem;
    color: #8888bb;
    font-weight: 300;
    letter-spacing: 0.08em;
}

/* ---------- Status Bar ---------- */
.status-bar {
    display: flex;
    justify-content: center;
    gap: 2rem;
    margin-bottom: 2rem;
    flex-wrap: wrap;
}

.status-item {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px;
    padding: 0.8rem 1.5rem;
    text-align: center;
    backdrop-filter: blur(10px);
    min-width: 160px;
}

.status-item .label {
    font-size: 0.75rem;
    color: #7b68ee;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    font-weight: 600;
}

.status-item .value {
    font-size: 1.3rem;
    color: #e0e0ff;
    font-weight: 700;
    margin-top: 0.2rem;
}

/* ---------- News Card ---------- */
.news-card {
    background: rgba(255, 255, 255, 0.04);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 16px;
    padding: 1.8rem;
    margin-bottom: 1.2rem;
    backdrop-filter: blur(12px);
    transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
    position: relative;
    overflow: hidden;
}

.news-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    width: 4px;
    height: 100%;
    background: linear-gradient(180deg, #00d2ff, #7b68ee);
    border-radius: 4px 0 0 4px;
    opacity: 0;
    transition: opacity 0.35s ease;
}

.news-card:hover {
    transform: translateY(-3px);
    border-color: rgba(123, 104, 238, 0.3);
    box-shadow: 0 12px 40px rgba(123, 104, 238, 0.12),
                0 4px 12px rgba(0, 0, 0, 0.2);
}

.news-card:hover::before {
    opacity: 1;
}

.news-card .card-date {
    font-size: 0.78rem;
    color: #7b68ee;
    font-weight: 500;
    letter-spacing: 0.04em;
    margin-bottom: 0.6rem;
    display: flex;
    align-items: center;
    gap: 0.4rem;
}

.news-card .card-title {
    font-size: 1.15rem;
    font-weight: 700;
    color: #e8e8ff;
    line-height: 1.5;
    margin-bottom: 0.8rem;
    display: -webkit-box;
    -webkit-line-clamp: 3;
    -webkit-box-orient: vertical;
    overflow: hidden;
}

.news-card .card-summary {
    font-size: 0.88rem;
    color: #9999cc;
    line-height: 1.7;
    margin-bottom: 1.2rem;
    display: -webkit-box;
    -webkit-line-clamp: 4;
    -webkit-box-orient: vertical;
    overflow: hidden;
}

.news-card .card-source {
    font-size: 0.75rem;
    color: #6666aa;
    margin-bottom: 1rem;
    font-weight: 400;
}

.news-card .read-more {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    padding: 0.55rem 1.4rem;
    background: linear-gradient(135deg, rgba(123, 104, 238, 0.2), rgba(0, 210, 255, 0.15));
    border: 1px solid rgba(123, 104, 238, 0.3);
    border-radius: 8px;
    color: #b0b0ff;
    font-size: 0.82rem;
    font-weight: 600;
    text-decoration: none;
    transition: all 0.3s ease;
    letter-spacing: 0.03em;
}

.news-card .read-more:hover {
    background: linear-gradient(135deg, rgba(123, 104, 238, 0.35), rgba(0, 210, 255, 0.25));
    color: #ffffff;
    transform: translateX(3px);
    box-shadow: 0 4px 15px rgba(123, 104, 238, 0.25);
}

/* ---------- Preset Buttons ---------- */
.preset-container {
    display: flex;
    flex-wrap: wrap;
    gap: 0.4rem;
    margin-top: 0.5rem;
}

.preset-btn {
    background: rgba(123, 104, 238, 0.1);
    border: 1px solid rgba(123, 104, 238, 0.25);
    border-radius: 20px;
    padding: 0.35rem 0.9rem;
    color: #b0b0ff;
    font-size: 0.75rem;
    cursor: pointer;
    transition: all 0.25s ease;
    text-decoration: none;
    display: inline-block;
}

.preset-btn:hover {
    background: rgba(123, 104, 238, 0.25);
    color: #ffffff;
}

/* ---------- Loading animation ---------- */
.loading-container {
    text-align: center;
    padding: 3rem;
}

.loading-text {
    color: #7b68ee;
    font-size: 1rem;
    font-weight: 500;
    animation: pulse 1.5s ease-in-out infinite;
}

@keyframes pulse {
    0%, 100% { opacity: 0.5; }
    50% { opacity: 1; }
}

/* ---------- No Results ---------- */
.no-results {
    text-align: center;
    padding: 4rem 2rem;
    color: #8888bb;
}

.no-results .icon {
    font-size: 3rem;
    margin-bottom: 1rem;
}

.no-results .message {
    font-size: 1.1rem;
    font-weight: 500;
    margin-bottom: 0.5rem;
    color: #aaaadd;
}

.no-results .hint {
    font-size: 0.85rem;
    color: #7777aa;
}

/* ---------- Footer ---------- */
.footer {
    text-align: center;
    padding: 2rem 1rem;
    margin-top: 2rem;
    border-top: 1px solid rgba(255,255,255,0.06);
    color: #555588;
    font-size: 0.78rem;
}

/* ---------- Divider ---------- */
.gradient-divider {
    height: 2px;
    background: linear-gradient(90deg, transparent, rgba(123,104,238,0.3), transparent);
    margin: 1rem 0 2rem;
    border: none;
}

/* ---------- Hide Streamlit defaults ---------- */
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────
# ヘルパー関数
# ─────────────────────────────────────
def build_google_news_rss_url(query: str, lang: str = "ja", country: str = "JP") -> str:
    """Google News RSS の URL を検索ワードから動的生成"""
    encoded = urllib.parse.quote(query)
    return (
        f"https://news.google.com/rss/search?"
        f"q={encoded}&hl={lang}&gl={country}&ceid={country}:{lang}"
    )


def clean_html(raw_html: str) -> str:
    """HTML タグを除去してプレーンテキストに変換"""
    text = re.sub(r'<[^>]+>', '', raw_html)
    text = html.unescape(text)
    return text.strip()


def format_date(entry) -> str:
    """フィードエントリの日時を読みやすい形式に変換"""
    if hasattr(entry, 'published_parsed') and entry.published_parsed:
        try:
            dt = datetime(*entry.published_parsed[:6])
            return dt.strftime("%Y年%m月%d日 %H:%M")
        except Exception:
            pass
    if hasattr(entry, 'published'):
        return entry.published
    return "日時不明"


def extract_source(entry) -> str:
    """ニュースソース（メディア名）を抽出"""
    if hasattr(entry, 'source') and hasattr(entry.source, 'title'):
        return entry.source.title
    title = entry.get('title', '')
    if ' - ' in title:
        return title.rsplit(' - ', 1)[-1]
    return ""


def extract_summary(entry) -> str:
    """要約テキストを抽出"""
    summary = entry.get('summary', '') or entry.get('description', '')
    return clean_html(summary)


@st.cache_data(ttl=300, show_spinner=False)
def fetch_news(query: str, lang: str, country: str):
    """ニュースフィードを取得（5 分キャッシュ）"""
    url = build_google_news_rss_url(query, lang, country)
    feed = feedparser.parse(url)
    articles = []
    for entry in feed.entries:
        source = extract_source(entry)
        title_clean = entry.get('title', '')
        if ' - ' in title_clean and source:
            title_clean = title_clean.rsplit(' - ', 1)[0]

        articles.append({
            'title': title_clean,
            'link': entry.get('link', '#'),
            'published': format_date(entry),
            'summary': extract_summary(entry),
            'source': source,
        })
    return articles, url


# ─────────────────────────────────────
# サイドバー
# ─────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔍 検索設定")
    st.markdown("---")

    search_query = st.text_input(
        "検索キーワード",
        value="Artificial Intelligence",
        help="Google News で検索するキーワードを入力してください",
        placeholder="例: AI 介護 福祉",
    )

    st.markdown("#### 🏷️ おすすめプリセット")
    st.caption("クリックすると検索ワードに反映されます")

    preset_keywords = [
        "AI 介護",
        "AI 福祉",
        "介護ロボット",
        "生成AI 介護現場",
        "AI 見守り 高齢者",
        "ChatGPT 介護",
        "AI ケアプラン",
        "介護DX",
        "AI 認知症",
        "Artificial Intelligence",
    ]

    cols = st.columns(2)
    for i, keyword in enumerate(preset_keywords):
        with cols[i % 2]:
            if st.button(keyword, key=f"preset_{i}", use_container_width=True):
                st.session_state["preset_query"] = keyword
                st.rerun()

    # プリセットボタンが押された場合に検索ワードを上書き
    if "preset_query" in st.session_state:
        search_query = st.session_state.pop("preset_query")

    st.markdown("---")

    st.markdown("#### 🌐 地域 / 言語")
    lang_options = {"日本語": ("ja", "JP"), "English": ("en", "US")}
    selected_lang = st.selectbox("表示言語", list(lang_options.keys()), index=0)
    lang_code, country_code = lang_options[selected_lang]

    st.markdown("---")
    st.markdown(
        "<div style='text-align:center; color:#555588; font-size:0.75rem;'>"
        "Powered by Google News RSS<br>& Streamlit 🚀"
        "</div>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────
# メインコンテンツ
# ─────────────────────────────────────

# ヘッダー
st.markdown(
    """
    <div class="main-header">
        <h1>🤖 AI News Dashboard</h1>
        <div class="subtitle">AI × 介護福祉 最新ニュースをリアルタイム収集</div>
    </div>
    <div class="gradient-divider"></div>
    """,
    unsafe_allow_html=True,
)

# ニュース取得
with st.spinner(""):
    st.markdown(
        '<div class="loading-container"><div class="loading-text">📡 ニュースを取得中...</div></div>',
        unsafe_allow_html=True,
    )
    articles, feed_url = fetch_news(search_query, lang_code, country_code)
    # placeholderを消すために再描画
    st.empty()

# ステータスバー
now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
st.markdown(
    f"""
    <div class="status-bar">
        <div class="status-item">
            <div class="label">検索ワード</div>
            <div class="value">{search_query}</div>
        </div>
        <div class="status-item">
            <div class="label">取得件数</div>
            <div class="value">{len(articles)} 件</div>
        </div>
        <div class="status-item">
            <div class="label">最終更新</div>
            <div class="value">{now_str}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# 記事一覧
if not articles:
    st.markdown(
        """
        <div class="no-results">
            <div class="icon">🔎</div>
            <div class="message">該当するニュースが見つかりませんでした</div>
            <div class="hint">別のキーワードで検索してみてください</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    # 2 カラムレイアウト
    col1, col2 = st.columns(2, gap="medium")

    for idx, article in enumerate(articles):
        with (col1 if idx % 2 == 0 else col2):
            source_html = (
                f'<div class="card-source">📰 {article["source"]}</div>'
                if article["source"]
                else ""
            )
            summary_html = (
                f'<div class="card-summary">{article["summary"]}</div>'
                if article["summary"]
                else ""
            )

            st.markdown(
                f"""
                <div class="news-card">
                    <div class="card-date">🕐 {article['published']}</div>
                    <div class="card-title">{article['title']}</div>
                    {summary_html}
                    {source_html}
                    <a class="read-more" href="{article['link']}" target="_blank" rel="noopener noreferrer">
                        記事を読む →
                    </a>
                </div>
                """,
                unsafe_allow_html=True,
            )

# フッター
st.markdown(
    f"""
    <div class="footer">
        AI News Dashboard v1.0 &nbsp;|&nbsp; Data from Google News RSS &nbsp;|&nbsp;
        Feed URL: <a href="{feed_url}" target="_blank" style="color:#7b68ee;">RSS</a>
    </div>
    """,
    unsafe_allow_html=True,
)
