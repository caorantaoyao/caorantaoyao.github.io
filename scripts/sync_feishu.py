#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书知识库 (Wiki) → 静态站点 同步脚本

在 GitHub Actions 中定时运行：
  1. 用 app_id / app_secret 换取 tenant_access_token
  2. 解析 Wiki 节点 token → space_id
  3. 递归遍历知识库节点树（父节点标题 = 分类目录 = 标签）
  4. 逐篇拉取 docx blocks，转成富文本 HTML
  5. 生成 data/articles.json（列表）与 articles/<slug>.html（每篇详情页）

设计约定：
  - 标签来源：文档所在的父节点（分类目录）标题，自动映射到配色 class
  - 只处理 obj_type == "docx" 的节点；容器/其它类型仅作为分类目录遍历

环境变量：
  FEISHU_APP_ID, FEISHU_APP_SECRET, FEISHU_WIKI_TOKEN
"""

import os
import re
import sys
import json
import html
import time
import hashlib
import datetime as dt

import requests

BASE = "https://open.feishu.cn/open-apis"
TIMEOUT = 30

# 输出目录（相对仓库根）
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
ARTICLES_DIR = os.path.join(ROOT, "articles")

# 分类目录标题 → 配色 class（沿用 index.html 的 .chemistry/.cloud/.security）
# 关键词命中即着色；未命中的分类默认走 accent(cloud) 色。
TAG_COLOR_RULES = [
    (("化学", "分子", "计算化学", "chem", "md", "fep"), "chemistry"),
    (("云", "云原生", "工程", "架构", "cloud", "devops", "k8s", "vke"), "cloud"),
    (("安全", "逆向", "漏洞", "security", "sec", "pwn", "ctf"), "security"),
]


def tag_color(category: str) -> str:
    low = (category or "").lower()
    for keywords, color in TAG_COLOR_RULES:
        for kw in keywords:
            if kw in low:
                return color
    return "cloud"


# ─────────────────────────────  飞书 API 封装  ─────────────────────────────

class Feishu:
    def __init__(self, app_id: str, app_secret: str):
        self.app_id = app_id
        self.app_secret = app_secret
        self.token = None
        self.session = requests.Session()

    def _auth(self):
        r = self.session.post(
            f"{BASE}/auth/v3/tenant_access_token/internal",
            json={"app_id": self.app_id, "app_secret": self.app_secret},
            timeout=TIMEOUT,
        )
        r.raise_for_status()
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"鉴权失败: {data}")
        self.token = data["tenant_access_token"]

    def _get(self, path: str, params: dict = None, _retry: int = 0):
        if not self.token:
            self._auth()
        headers = {"Authorization": f"Bearer {self.token}"}
        r = self.session.get(f"{BASE}{path}", headers=headers,
                             params=params or {}, timeout=TIMEOUT)
        # 频控退避
        if r.status_code == 400 and _retry < 5:
            try:
                if r.json().get("code") == 99991400:
                    time.sleep(2 ** _retry)
                    return self._get(path, params, _retry + 1)
            except Exception:
                pass
        r.raise_for_status()
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"GET {path} 失败: {data}")
        return data["data"]

    # 解析 wiki 节点 → space_id + node 信息
    def get_node(self, token: str) -> dict:
        return self._get("/wiki/v2/spaces/get_node", {"token": token})["node"]

    # 分页列出某父节点下的子节点
    def list_children(self, space_id: str, parent_token: str = None):
        items, page_token = [], None
        while True:
            params = {"page_size": 50}
            if parent_token:
                params["parent_node_token"] = parent_token
            if page_token:
                params["page_token"] = page_token
            data = self._get(f"/wiki/v2/spaces/{space_id}/nodes", params)
            items.extend(data.get("items", []))
            if data.get("has_more") and data.get("page_token"):
                page_token = data["page_token"]
            else:
                break
        return items

    # 拉取 docx 全部 blocks（分页）
    def get_blocks(self, document_id: str):
        items, page_token = [], None
        while True:
            params = {"page_size": 500, "document_revision_id": -1}
            if page_token:
                params["page_token"] = page_token
            data = self._get(f"/docx/v1/documents/{document_id}/blocks", params)
            items.extend(data.get("items", []))
            if data.get("has_more") and data.get("page_token"):
                page_token = data["page_token"]
            else:
                break
        return items


# ─────────────────────────────  docx blocks → HTML  ─────────────────────────────

# 代码块 language 枚举（飞书 docx 常见值 → 高亮 class 名）
CODE_LANG = {
    1: "plaintext", 8: "c", 9: "csharp", 10: "cpp", 12: "css", 22: "go",
    23: "html", 26: "java", 27: "javascript", 28: "json", 30: "kotlin",
    36: "markdown", 43: "objectivec", 49: "php", 51: "powershell",
    52: "python", 53: "r", 55: "ruby", 56: "rust", 58: "shell",
    60: "sql", 61: "swift", 63: "typescript", 65: "xml", 66: "yaml",
}


def esc(s: str) -> str:
    return html.escape(s or "", quote=True)


def render_elements(elements) -> str:
    """把一个 block 的 elements（text_run 数组）渲染为内联 HTML"""
    out = []
    for el in elements or []:
        tr = el.get("text_run")
        if not tr:
            # 忽略 mention_user / equation 等，尽量取其纯文本
            eq = el.get("equation")
            if eq and eq.get("content"):
                out.append(f"<code>{esc(eq['content'])}</code>")
            continue
        content = esc(tr.get("content", ""))
        style = tr.get("text_element_style", {}) or {}
        if style.get("inline_code"):
            content = f"<code>{content}</code>"
        if style.get("bold"):
            content = f"<strong>{content}</strong>"
        if style.get("italic"):
            content = f"<em>{content}</em>"
        if style.get("strikethrough"):
            content = f"<s>{content}</s>"
        if style.get("underline"):
            content = f"<u>{content}</u>"
        link = style.get("link")
        if link and link.get("url"):
            url = requests.utils.unquote(link["url"])
            content = f'<a href="{esc(url)}" target="_blank" rel="noopener">{content}</a>'
        out.append(content)
    return "".join(out)


def blocks_to_html(blocks) -> tuple:
    """
    把 docx blocks（文档顺序的扁平列表）转成 HTML 正文。
    返回 (html_body, first_text) —— first_text 用于生成摘要。
    """
    html_parts = []
    first_text = ""
    # 列表分组缓冲
    list_buf, list_tag = [], None

    def flush_list():
        nonlocal list_buf, list_tag
        if list_buf:
            items = "".join(f"<li>{x}</li>" for x in list_buf)
            html_parts.append(f"<{list_tag}>{items}</{list_tag}>")
            list_buf, list_tag = [], None

    def block_text(b, key):
        return (b.get(key) or {}).get("elements", [])

    for b in blocks:
        bt = b.get("block_type")

        # 段落
        if bt == 2:
            flush_list()
            inner = render_elements(block_text(b, "text"))
            if inner.strip():
                html_parts.append(f"<p>{inner}</p>")
                if not first_text:
                    first_text = re.sub(r"<[^>]+>", "", inner)
        # 标题 h1~h9 → 页面里降一级，h1 留给文章标题
        elif bt in (3, 4, 5, 6, 7, 8, 9, 10, 11):
            flush_list()
            level = min(bt - 1, 6)  # h2..h6
            key = f"heading{bt - 2}"
            inner = render_elements(block_text(b, key))
            html_parts.append(f"<h{level}>{inner}</h{level}>")
        # 无序列表
        elif bt == 12:
            if list_tag not in (None, "ul"):
                flush_list()
            list_tag = "ul"
            list_buf.append(render_elements(block_text(b, "bullet")))
        # 有序列表
        elif bt == 13:
            if list_tag not in (None, "ol"):
                flush_list()
            list_tag = "ol"
            list_buf.append(render_elements(block_text(b, "ordered")))
        # 代码块
        elif bt == 14:
            flush_list()
            code = b.get("code") or {}
            lang = CODE_LANG.get((code.get("style") or {}).get("language"), "plaintext")
            raw = "".join(
                (el.get("text_run") or {}).get("content", "")
                for el in code.get("elements", [])
            )
            html_parts.append(
                f'<pre><code class="language-{lang}">{esc(raw)}</code></pre>'
            )
        # 引用
        elif bt == 15:
            flush_list()
            inner = render_elements(block_text(b, "quote"))
            html_parts.append(f"<blockquote>{inner}</blockquote>")
        # 待办
        elif bt == 17:
            flush_list()
            todo = b.get("todo") or {}
            done = (todo.get("style") or {}).get("done")
            mark = "☑" if done else "☐"
            inner = render_elements(todo.get("elements", []))
            html_parts.append(f'<p class="todo">{mark} {inner}</p>')
        # 高亮块 callout
        elif bt == 19:
            flush_list()
            inner = render_elements(block_text(b, "callout"))
            if inner.strip():
                html_parts.append(f'<div class="callout">{inner}</div>')
        # 分割线
        elif bt == 22:
            flush_list()
            html_parts.append("<hr>")
        # 图片
        elif bt == 27:
            flush_list()
            # 图片需二次下载素材, 此处先占位, 后续可扩展 drive 下载
            html_parts.append('<p class="img-placeholder">[图片]</p>')
        # 其它类型忽略

    flush_list()
    return "\n".join(html_parts), first_text.strip()


# ─────────────────────────────  详情页模板  ─────────────────────────────

DETAIL_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title} — 曹然</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@400;500;600;700;900&family=Spectral:ital,wght@0,400;0,500;0,600;0,700;1,400&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg:#f5f1e8; --surface:#fbf8f1; --surface-raised:#efe8da; --border:#ddd3c1; --border-soft:#e6ddcc;
    --text:#29241f; --text-muted:#574f45; --text-dim:#928979;
    --accent:#c0553b; --accent-bright:#9c4229; --accent-glow:rgba(192,85,59,0.09); --accent-dim:#b98a3e;
    --green:#2f5d4e; --green-dim:#5c7d70; --red:#b06a6a; --blue:#7f96b0;
    --font-display:'Spectral','Noto Serif SC',serif;
    --font-serif:'Noto Serif SC','Source Han Serif SC','Songti SC',serif;
    --font-mono:'JetBrains Mono','SF Mono',monospace;
    --max-width:760px; --ease:cubic-bezier(0.22,1,0.36,1);
  }}
  *,*::before,*::after {{ margin:0; padding:0; box-sizing:border-box; }}
  html {{ scroll-behavior:smooth; }}
  body {{ background:var(--bg); color:var(--text); font-family:var(--font-serif);
    line-height:1.9; -webkit-font-smoothing:antialiased; }}
  ::selection {{ background:var(--accent); color:var(--bg); }}
  body::after {{ content:''; position:fixed; inset:0; z-index:0; pointer-events:none; opacity:0.5;
    background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='140' height='140'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.85' numOctaves='3' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.045'/%3E%3C/svg%3E"); }}
  #progress {{ position:fixed; top:0; left:0; height:2px; width:0; z-index:200;
    background:linear-gradient(90deg,var(--accent-dim),var(--accent-bright)); transition:width 0.1s linear; }}
  .container {{ position:relative; z-index:1; max-width:var(--max-width); margin:0 auto; padding:0 1.5rem; }}
  nav {{ position:sticky; top:0; z-index:100; padding:1.2rem 0;
    backdrop-filter:blur(20px) saturate(160%); -webkit-backdrop-filter:blur(20px) saturate(160%);
    background:rgba(245,241,232,0.85); border-bottom:1px solid var(--border-soft); }}
  nav .container {{ display:flex; justify-content:space-between; align-items:center; }}
  .nav-logo {{ font-family:var(--font-mono); font-size:0.85rem; letter-spacing:0.25em;
    color:var(--accent); text-decoration:none; text-transform:uppercase; display:inline-flex; align-items:center; gap:0.5rem; }}
  .nav-logo .dot {{ width:6px; height:6px; border-radius:50%; background:var(--accent); }}
  .nav-back {{ font-family:var(--font-mono); font-size:0.72rem; letter-spacing:0.08em;
    color:var(--text-muted); text-decoration:none; text-transform:uppercase; transition:color 0.2s; }}
  .nav-back:hover {{ color:var(--accent); }}
  article {{ padding:clamp(3rem,8vh,5rem) 0 4rem; opacity:0; transform:translateY(20px);
    animation:fadeUp 0.8s var(--ease) 0.1s forwards; }}
  @keyframes fadeUp {{ to {{ opacity:1; transform:translateY(0); }} }}
  .article-meta {{ display:flex; gap:1rem; align-items:center; margin-bottom:1.5rem;
    font-family:var(--font-mono); font-size:0.7rem; letter-spacing:0.05em; }}
  .article-date {{ color:var(--text-dim); }}
  .article-tag {{ text-transform:uppercase; padding:0.2rem 0.6rem; border:1px solid var(--border);
    border-radius:2px; color:var(--text-muted); }}
  .article-tag.chemistry {{ color:var(--green); border-color:var(--green-dim); }}
  .article-tag.cloud {{ color:var(--accent); border-color:var(--accent-dim); }}
  .article-tag.security {{ color:var(--red); border-color:rgba(176,106,106,0.4); }}
  h1.article-title {{ font-family:var(--font-display); font-size:clamp(1.8rem,4vw,2.6rem); font-weight:700; line-height:1.3;
    letter-spacing:-0.02em; margin-bottom:2.5rem; }}
  .article-body {{ font-size:1.02rem; color:var(--text-muted); }}
  .article-body h2 {{ font-family:var(--font-display); font-size:1.5rem; font-weight:600; margin:2.5rem 0 1rem;
    color:var(--text); letter-spacing:-0.01em; }}
  .article-body h3 {{ font-family:var(--font-display); font-size:1.2rem; font-weight:600; margin:2rem 0 0.8rem; color:var(--text); }}
  .article-body h4,.article-body h5,.article-body h6 {{ font-size:1.05rem; font-weight:600;
    margin:1.5rem 0 0.6rem; color:var(--text); }}
  .article-body p {{ margin-bottom:1.3rem; }}
  .article-body a {{ color:var(--accent); text-decoration:none; border-bottom:1px solid var(--accent-dim); }}
  .article-body a:hover {{ border-bottom-color:var(--accent); }}
  .article-body ul,.article-body ol {{ margin:0 0 1.3rem 1.5rem; }}
  .article-body li {{ margin-bottom:0.5rem; }}
  .article-body code {{ font-family:var(--font-mono); font-size:0.85em;
    background:var(--surface-raised); padding:0.15em 0.4em; border-radius:3px; color:var(--accent-bright); }}
  .article-body pre {{ background:var(--surface); border:1px solid var(--border); border-radius:4px;
    padding:1.2rem; overflow-x:auto; margin-bottom:1.3rem; }}
  .article-body pre code {{ background:none; padding:0; color:var(--text); font-size:0.82rem; line-height:1.7; }}
  .article-body blockquote {{ border-left:2px solid var(--accent-dim); padding-left:1.2rem;
    margin:0 0 1.3rem; color:var(--text-muted); font-style:italic; }}
  .article-body .callout {{ background:var(--surface); border:1px solid var(--border);
    border-left:2px solid var(--accent); border-radius:4px; padding:1rem 1.2rem; margin-bottom:1.3rem; }}
  .article-body hr {{ border:none; border-top:1px solid var(--border); margin:2.5rem 0; }}
  .article-body .todo {{ color:var(--text-muted); }}
  .article-body .img-placeholder {{ color:var(--text-dim); font-family:var(--font-mono);
    font-size:0.8rem; text-align:center; padding:1rem; border:1px dashed var(--border); }}
  footer {{ border-top:1px solid var(--border); padding:2rem 0; margin-top:3rem;
    font-family:var(--font-mono); font-size:0.68rem; color:var(--text-dim); text-align:center; }}
</style>
</head>
<body>
<div id="progress"></div>
<nav>
  <div class="container">
    <a href="/" class="nav-logo"><span class="dot"></span>曹然</a>
    <a href="/library.html" class="nav-back">← 返回知识库</a>
  </div>
</nav>
<main class="container">
  <article>
    <div class="article-meta">
      <span class="article-date">{date}</span>
      <span class="article-tag {tag_color}">{category}</span>
    </div>
    <h1 class="article-title">{title}</h1>
    <div class="article-body">
{body}
    </div>
  </article>
</main>
<footer>&copy; 2026 曹然 · 技术札记与深度思考</footer>
<script>
  var p=document.getElementById('progress');
  addEventListener('scroll',function(){{
    var h=document.documentElement,max=h.scrollHeight-h.clientHeight;
    p.style.width=(max>0?(scrollY/max*100):0)+'%';
  }},{{passive:true}});
</script>
</body>
</html>
"""


def slugify(title: str, token: str) -> str:
    """生成稳定、URL 安全的文件名：优先中文转拼音不可得，用 token 短哈希兜底"""
    base = re.sub(r"[^\w\u4e00-\u9fff]+", "-", title).strip("-")
    h = hashlib.md5(token.encode("utf-8")).hexdigest()[:8]
    if base:
        # 保留可读的 ascii 部分, 否则纯用 hash
        ascii_part = re.sub(r"[^a-zA-Z0-9\-]+", "", base).strip("-")
        return f"{ascii_part}-{h}" if ascii_part else h
    return h


def ts_to_date(ts) -> str:
    try:
        return dt.datetime.fromtimestamp(int(ts), dt.timezone.utc).astimezone().strftime("%Y.%m.%d")
    except Exception:
        return dt.datetime.now().strftime("%Y.%m.%d")


# ─────────────────────────────  主流程  ─────────────────────────────

def walk(fs: Feishu, space_id: str, parent_token: str, category: str, articles: list,
         skip_token: str = None):
    """
    递归遍历知识库节点。
    category = 当前父节点标题（分类目录）；顶层节点的 category 由其自身标题决定。
    skip_token = 需跳过的节点（如知识库根节点/首页），仅跳过其自身内容，仍下钻其子节点。
    """
    children = fs.list_children(space_id, parent_token)
    for node in children:
        title = node.get("title") or "(未命名)"
        obj_type = node.get("obj_type")
        node_token = node.get("node_token")
        has_child = node.get("has_child")

        # 跳过根节点/首页自身：不收录为文章，但仍继续遍历其子节点
        if skip_token and node_token == skip_token:
            if has_child:
                walk(fs, space_id, node_token, category, articles, skip_token)
            continue

        # 顶层节点（category 为空）以自身标题作为分类；下层沿用父目录分类
        node_category = category or title

        if obj_type == "docx" and node.get("obj_token"):
            try:
                blocks = fs.get_blocks(node["obj_token"])
                body_html, first_text = blocks_to_html(blocks)
                slug = slugify(title, node_token)
                date = ts_to_date(node.get("obj_edit_time") or node.get("obj_create_time"))
                color = tag_color(node_category)
                summary = (first_text[:120] + "…") if len(first_text) > 120 else first_text

                # 写详情页
                page = DETAIL_TEMPLATE.format(
                    title=esc(title), date=date, category=esc(node_category or "文章"),
                    tag_color=color, body=body_html or "<p>（暂无正文）</p>",
                )
                os.makedirs(ARTICLES_DIR, exist_ok=True)
                with open(os.path.join(ARTICLES_DIR, f"{slug}.html"), "w", encoding="utf-8") as f:
                    f.write(page)

                articles.append({
                    "title": title,
                    "date": date,
                    "category": node_category or "文章",
                    "tag_color": color,
                    "summary": summary or title,
                    "url": f"/articles/{slug}.html",
                    "sort_ts": int(node.get("obj_edit_time") or node.get("obj_create_time") or 0),
                })
                print(f"  ✓ [{node_category}] {title} -> {slug}.html")
            except Exception as e:
                print(f"  ✗ 跳过 {title}: {e}", file=sys.stderr)

        # 若该节点还有子节点，则它本身作为分类目录继续下钻
        if has_child:
            # 分类名：容器节点用自身标题；docx 节点仍沿用上层分类
            next_category = title if obj_type != "docx" else node_category
            walk(fs, space_id, node_token, next_category, articles)


def main():
    app_id = os.environ.get("FEISHU_APP_ID")
    app_secret = os.environ.get("FEISHU_APP_SECRET")
    wiki_token = os.environ.get("FEISHU_WIKI_TOKEN")
    if not (app_id and app_secret and wiki_token):
        print("缺少环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET / FEISHU_WIKI_TOKEN", file=sys.stderr)
        sys.exit(1)

    fs = Feishu(app_id, app_secret)
    print("· 解析知识库根节点 …")
    root = fs.get_node(wiki_token)
    space_id = root["space_id"]
    root_token = root.get("node_token")
    print(f"· space_id = {space_id}, 根节点 = {root.get('title')}")

    articles = []
    # 顶层节点是"知识空间"的直接子节点，必须不传 parent_node_token 才能列出；
    # 传入根 docx 节点的 token 只会得到 0 条（它自身没有 wiki 子节点）。
    # skip_token=root_token：跳过知识库首页/根节点自身，不将其收录为文章。
    walk(fs, space_id, None, None, articles, skip_token=root_token)

    # 按更新时间倒序
    articles.sort(key=lambda a: a["sort_ts"], reverse=True)
    for a in articles:
        a.pop("sort_ts", None)

    os.makedirs(DATA_DIR, exist_ok=True)
    # 始终确保 articles/ 目录存在，避免 0 篇时工作流 `git add articles/` 报错
    os.makedirs(ARTICLES_DIR, exist_ok=True)
    out = {
        "generated_at": dt.datetime.now(dt.timezone.utc).astimezone().isoformat(),
        "count": len(articles),
        "articles": articles,
    }
    with open(os.path.join(DATA_DIR, "articles.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    # 全量覆盖：删除本次结果不再引用的旧详情页（孤儿文件）。
    # 此步骤在文章已全部成功生成之后执行，保证原子性——同步失败时不会误删。
    referenced = {os.path.basename(a["url"]) for a in articles}
    removed = 0
    for fn in os.listdir(ARTICLES_DIR):
        if fn.endswith(".html") and fn not in referenced:
            try:
                os.remove(os.path.join(ARTICLES_DIR, fn))
                removed += 1
                print(f"  - 清理孤儿页 {fn}")
            except OSError as e:
                print(f"  ✗ 清理 {fn} 失败: {e}", file=sys.stderr)

    print(f"· 完成，共 {len(articles)} 篇文章 -> data/articles.json"
          + (f"（清理 {removed} 个孤儿页）" if removed else ""))


if __name__ == "__main__":
    main()
