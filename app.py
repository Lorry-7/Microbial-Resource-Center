# app.py
# 运行：streamlit run app.py

import math
import os
import re
import tempfile
from io import BytesIO
from typing import Dict, List, Optional, Tuple
import html as _html

import pandas as pd
import streamlit as st

# 用于提取 Excel 嵌入图片
from openpyxl import load_workbook
from PIL import Image


# ---------------------------
# 基础配置
# ---------------------------
st.set_page_config(page_title="微生物资源目录", layout="wide")

SANDBOX_XLSX_PATH = "/mnt/data/Cqut_Microbial_Resource_Center_Data.xlsx"
LOCAL_FALLBACK_XLSX = "Cqut_Microbial_Resource_Center_Data.xlsx"

ID_COL_CANDIDATES = ["菌种编号", "编号", "ID", "id", "strain_id"]
IMAGE_COL_CANDIDATES = ["菌种照片", "图片", "照片", "image", "img", "photo", "picture"]

# 导出嵌入图片的目录：部署环境通常只有 /tmp 可写，所以用 tempfile.gettempdir()
EXPORT_ROOT = os.path.join(tempfile.gettempdir(), "_extracted_images")


# ---------------------------
# CSS：更克制、更科研门户风（统一间距/层级/控件/表格）
# ---------------------------
CSS = """
<style>
:root{
  --bg: #f5f7fb;
  --card: #ffffff;
  --text: #0f172a;
  --muted: #475569;
  --muted2:#64748b;
  --border:#e6eef7;
  --border2:#dbe7f6;
  --head:#0d3b66;
  --head2:#0b4f8a;
  --head3:#0a5aa0;
  --soft:#f1f5fb;
  --hover:#f8fbff;
  --shadow: 0 10px 25px rgba(15, 23, 42, .06);
  --radius: 14px;
  --radius2: 12px;
}

html, body, [class*="css"]{
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB",
               "Microsoft YaHei", "Noto Sans CJK SC", Arial, sans-serif;
  color: var(--text);
}

.stApp{ background: var(--bg); }

/* 主内容最大宽度与留白节奏 */
.main .block-container{
  padding-top: 1.1rem;
  padding-bottom: 2.0rem;
  max-width: 1400px;
}

/* 顶部横幅：更克制的圆角与阴影 */
.nimr-topbar{
  background: linear-gradient(90deg,var(--head) 0%,var(--head2) 45%,var(--head3) 100%);
  border: 1px solid rgba(255,255,255,0.18);
  padding: 18px 22px;
  border-radius: var(--radius);
  box-shadow: var(--shadow);
}
.nimr-topbar .row{display:flex;gap:16px;align-items:center;}
.nimr-logo{
  width:54px;height:54px;border-radius: 14px;
  background: rgba(255,255,255,0.16);
  display:flex;align-items:center;justify-content:center;
  color:#fff;font-weight:900;letter-spacing:1px;
  border:1px solid rgba(255,255,255,0.22);
}
.nimr-title{color:#fff;font-size:20px;font-weight:900;line-height:1.18;}
.nimr-subtitle{color:rgba(255,255,255,0.86);margin-top:4px;font-size:12.5px;}

/* 面包屑 */
.nimr-breadcrumb{
  margin-top: 14px;
  font-size:13px;color: var(--muted);
  padding:10px 12px;background: var(--card);
  border:1px solid var(--border);
  border-radius: var(--radius2);
  box-shadow: 0 6px 18px rgba(15,23,42,.03);
}
.nimr-breadcrumb b{color: var(--text);}

/* 标题条：左侧强调线 + 更紧凑 */
.nimr-section-title{
  margin-top: 12px;
  background: var(--card);
  border:1px solid var(--border);
  border-left: 5px solid var(--head2);
  padding:10px 12px;
  border-radius: var(--radius2);
  font-weight:900;
  color: var(--text);
  box-shadow: 0 6px 18px rgba(15,23,42,.03);
}

/* 卡片：统一阴影、边框、圆角 */
.card{
  margin-top:12px;
  background: var(--card);
  border:1px solid var(--border);
  border-radius: var(--radius2);
  padding: 14px 14px 12px 14px;
  box-shadow: 0 10px 25px rgba(15,23,42,.04);
}

/* 让 Streamlit 的输入控件更像门户表单 */
div[data-testid="stTextInput"] label{
  color: var(--muted);
  font-size: 12.5px;
  font-weight: 700;
}
div[data-testid="stTextInput"] input{
  border-radius: 10px !important;
  border: 1px solid var(--border2) !important;
  background: #fff !important;
}
div[data-testid="stTextInput"] input:focus{
  outline: none !important;
  border-color: rgba(11,79,138,.55) !important;
  box-shadow: 0 0 0 3px rgba(11,79,138,.12) !important;
}

/* 按钮：统一圆角与轻阴影 */
div.stButton > button{
  border-radius: 12px !important;
  border: 1px solid var(--border2) !important;
  background: #fff !important;
  color: var(--text) !important;
  font-weight: 800 !important;
  box-shadow: 0 8px 18px rgba(15,23,42,.04);
}
div.stButton > button:hover{
  background: var(--hover) !important;
  border-color: rgba(11,79,138,.35) !important;
}
div.stButton > button:disabled{
  opacity: .55 !important;
  box-shadow: none !important;
}

/* 表格容器 */
.table-wrap{
  margin-top:12px;
  background: var(--card);
  border:1px solid var(--border);
  border-radius: var(--radius2);
  padding: 10px 12px 12px 12px;
  box-shadow: 0 10px 25px rgba(15,23,42,.04);
}

/* 表格本体（自渲染 HTML） */
.nimr-table{
  width:100%;
  border-collapse: separate;
  border-spacing: 0;
  font-size: 13px;
  overflow: hidden;
  border-radius: 12px;
}
.nimr-table thead th{
  background: var(--soft);
  color: var(--text);
  font-weight: 900;
  text-align: center;
  padding: 10px 10px;
  border-bottom: 1px solid var(--border2);
}
.nimr-table tbody td{
  padding: 10px 10px;
  border-bottom: 1px solid var(--border2);
  color: #111827;
  vertical-align: top;
  background: #fff;
}
.nimr-table tbody tr:hover td{ background: var(--hover); }
.nimr-table tbody tr:last-child td{ border-bottom: none; }

/* 单元格对齐 */
.td-center{ text-align:center; white-space:nowrap; }
.td-left{ text-align:left; }

/* 分页信息 */
.pager{
  margin-top:10px;
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap: 10px;
  color: var(--muted);
  font-size: 13px;
}
.pill{
  background: var(--soft);
  border:1px solid var(--border2);
  padding:6px 10px;
  border-radius:999px;
  white-space: nowrap;
}

/* 详情：KV 表 */
.kv{width:100%;border-collapse:separate;border-spacing:0;font-size:13px;overflow:hidden;border-radius:12px;}
.kv td{border-bottom:1px solid var(--border2);padding:10px 10px;vertical-align:top;background:#fff;}
.kv tr:last-child td{border-bottom:none;}
.kv .k{
  width: 190px;
  background: var(--soft);
  font-weight: 900;
  color: var(--text);
  border-right: 1px solid var(--border2);
}

/* 链接 */
a.nimr-link{color: var(--head2);text-decoration:none;font-weight:900;}
a.nimr-link:hover{text-decoration:underline;}

/* 顶部 Streamlit 的多余元素可隐藏（不算新增内容，只是更干净） */
#MainMenu{visibility:hidden;}
footer{visibility:hidden;}
header{visibility:hidden;}
</style>
"""


# ---------------------------
# 数据加载与清洗
# ---------------------------
@st.cache_data(show_spinner=False)
def load_excel(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.replace({pd.NA: "", None: ""}).fillna("")
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].astype(str).map(lambda x: re.sub(r"[ \t]+", " ", x).strip())
    return df


def pick_excel_path() -> str:
    if os.path.exists(SANDBOX_XLSX_PATH):
        return SANDBOX_XLSX_PATH
    if os.path.exists(LOCAL_FALLBACK_XLSX):
        return LOCAL_FALLBACK_XLSX
    raise FileNotFoundError("未找到 Excel 文件：请确认已上传或放在同目录。")


def detect_id_col(df: pd.DataFrame) -> str:
    for c in ID_COL_CANDIDATES:
        if c in df.columns:
            return c
    return df.columns[0]


def detect_image_col(df: pd.DataFrame) -> Optional[str]:
    for c in IMAGE_COL_CANDIDATES:
        if c in df.columns:
            return c
    return None


# ---------------------------
# URL Query Params
# ---------------------------
def get_query_id() -> str:
    return str(st.query_params.get("id", "")).strip()


def set_query_id(val: str):
    if val:
        st.query_params["id"] = val
    else:
        st.query_params.clear()


# ---------------------------
# 头部/面包屑
# ---------------------------
def render_header():
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="nimr-topbar">
          <div class="row">
            <div class="nimr-logo">MR</div>
            <div>
              <div class="nimr-title">Lab 106 · 微生物资源中心（资源目录）</div>
              <div class="nimr-subtitle">Lab 106 · Microbial Resource Center (Resource Catalog)</div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_breadcrumb(labels: List[Tuple[str, bool]]):
    parts = []
    for lab, cur in labels:
        parts.append(f"<b>{lab}</b>" if cur else lab)
    st.markdown(
        f'<div class="nimr-breadcrumb">当前位置： {" &nbsp;›&nbsp; ".join(parts)}</div>',
        unsafe_allow_html=True,
    )


# ---------------------------
# 搜索（全字段）
# ---------------------------
def build_global_search_mask(df: pd.DataFrame, query: str) -> pd.Series:
    q = (query or "").strip().lower()
    if not q:
        return pd.Series([True] * len(df), index=df.index)
    all_text = df.astype(str).agg(" | ".join, axis=1).str.lower()
    return all_text.str.contains(re.escape(q), na=False)


def short_text(s: str, n: int = 60) -> str:
    s = str(s or "").replace("\n", " / ").strip()
    return (s[:n] + "…") if len(s) > n else s


# ---------------------------
# 分页（稳定：session_state）
# ---------------------------
def ensure_pagination_state(total_pages: int):
    if "page" not in st.session_state:
        st.session_state.page = 1
    st.session_state.page = max(1, min(int(st.session_state.page), total_pages))


# ---------------------------
# 图片：支持单元格文本路径/URL + Excel 嵌入图片对象
# ---------------------------
def split_image_tokens(s: str) -> List[str]:
    if not s:
        return []
    s = str(s).strip()
    if not s or s.lower() in ("nan", "none"):
        return []
    tokens = re.split(r"[;\n,|]+", s)
    return [t.strip().strip('"').strip("'") for t in tokens if t.strip()]


def resolve_image_path(token: str, excel_dir: str) -> Optional[str]:
    if not token:
        return None
    t = token.strip()
    if re.match(r"^https?://", t, re.IGNORECASE):
        return t
    if t.lower().startswith("file://"):
        t = t[7:].strip()

    # 绝对路径
    if os.path.isabs(t) and os.path.exists(t):
        return t

    # 相对路径/文件名：多目录尝试
    candidates = [
        os.path.join(excel_dir, t),
        os.path.join(tempfile.gettempdir(), t),
        os.path.join(os.getcwd(), t),
        t,
    ]
    for p in candidates:
        try:
            if os.path.exists(p):
                return p
        except Exception:
            continue
    return None


@st.cache_data(show_spinner=False)
def extract_embedded_images(excel_path: str) -> Dict[int, List[str]]:
    """
    从 Excel 工作表中提取嵌入图片对象，按“图片锚点所在行号(Excel行号)”索引。
    返回：
      { excel_row_number: [saved_image_path1, saved_image_path2, ...], ... }
    注意：Excel 行号从 1 开始。
    """
    wb = load_workbook(excel_path)
    ws = wb.worksheets[0]

    # 为避免旧文件残留：按文件特征创建导出目录
    st_info = os.stat(excel_path)
    tag = f"mtime{int(st_info.st_mtime)}_size{st_info.st_size}"
    out_dir = os.path.join(EXPORT_ROOT, tag)

    # 部署环境可能只读：无法落盘则直接返回空映射，不影响详情页
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception:
        return {}

    mapping: Dict[int, List[str]] = {}
    images = getattr(ws, "_images", [])

    for idx, img in enumerate(images, start=1):
        try:
            anchor = img.anchor
            row0 = anchor._from.row  # type: ignore
            excel_row = int(row0) + 1
        except Exception:
            continue

        try:
            raw = img._data()  # type: ignore
        except Exception:
            try:
                raw = img._data  # type: ignore
            except Exception:
                continue

        try:
            im = Image.open(BytesIO(raw))
            save_path = os.path.join(out_dir, f"row{excel_row}_img{idx}.png")
            im.save(save_path, format="PNG")
            mapping.setdefault(excel_row, []).append(save_path)
        except Exception:
            continue

    return mapping


def get_images_for_record(
    df: pd.DataFrame,
    excel_path: str,
    df_row_index: int,
    img_col: Optional[str],
) -> List[str]:
    """
    聚合两类图片来源：
    1) 单元格文本（URL/路径） -> resolved paths
    2) Excel 嵌入图片对象 -> extracted png paths
    """
    excel_dir = os.path.dirname(os.path.abspath(excel_path))
    results: List[str] = []

    # 1) 文本列
    if img_col and img_col in df.columns:
        tokens = split_image_tokens(df.loc[df_row_index, img_col])
        for t in tokens:
            p = resolve_image_path(t, excel_dir)
            if p:
                results.append(p)

    # 2) 嵌入图片：将 df 行号映射到 Excel 行号（df第0行≈Excel第2行，Excel第1行是表头）
    excel_row = int(df_row_index) + 2
    embedded_map = extract_embedded_images(excel_path)
    for p in embedded_map.get(excel_row, []):
        results.append(p)

    # 去重
    dedup = []
    for p in results:
        if p not in dedup:
            dedup.append(p)
    return dedup


# ---------------------------
# 列表表格：自渲染 HTML（更可控、更美观）
# ---------------------------
def _render_table_html(df_show: pd.DataFrame, center_cols: Optional[set] = None) -> str:
    center_cols = center_cols or set()

    cols = list(df_show.columns)
    thead = "".join([f"<th>{_html.escape(str(c))}</th>" for c in cols])

    body_rows = []
    for _, r in df_show.iterrows():
        tds = []
        for c in cols:
            v = r.get(c, "")
            if c == "操作":
                # 允许 HTML（链接）
                cell = str(v)
            else:
                cell = _html.escape(str(v))
            cls = "td-center" if c in center_cols else "td-left"
            tds.append(f'<td class="{cls}">{cell if cell else "&nbsp;"}</td>')
        body_rows.append("<tr>" + "".join(tds) + "</tr>")

    return f"""
    <table class="nimr-table">
      <thead><tr>{thead}</tr></thead>
      <tbody>
        {''.join(body_rows)}
      </tbody>
    </table>
    """


# ---------------------------
# 列表页
# ---------------------------
def render_list(df: pd.DataFrame, id_col: str):
    render_breadcrumb([("首页", False), ("资源目录", True)])
    st.markdown('<div class="nimr-section-title">微生物资源目录</div>', unsafe_allow_html=True)

    # 固定每页 10 条
    page_size = 10

    # 过滤器卡片
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("**菌种检索**")

    # 选择特定字段作为检索条件
    search_cols = ["菌种编号", "菌种命名", "菌种来源", "申请人"]
    search_cols = [col for col in search_cols if col in df.columns]

    search_conditions = {}

    cols = st.columns(2, gap="large")
    for i, col in enumerate(search_cols):
        with cols[i % 2]:
            search_conditions[col] = st.text_input(
                f"{col}",
                value="",
                placeholder=f"输入{col}关键词",
                key=f"search_{col}",
            )

    filtered = df.copy()
    for col, value in search_conditions.items():
        if value.strip():
            filtered = filtered[filtered[col].astype(str).str.contains(value.strip(), case=False, na=False)]

    filtered = filtered.reset_index(drop=True)
    st.markdown("</div>", unsafe_allow_html=True)

    total = len(filtered)
    total_pages = max(1, math.ceil(total / page_size))
    ensure_pagination_state(total_pages)

    st.markdown('<div class="table-wrap">', unsafe_allow_html=True)

    start = (st.session_state.page - 1) * page_size
    end = start + page_size
    page_df = filtered.iloc[start:end].copy()

    preferred = [id_col, "菌种命名", "属、种", "保藏日期", "菌种来源"]
    show_cols = [c for c in preferred if c in page_df.columns]
    if len(show_cols) < 3:
        show_cols = list(page_df.columns[: min(5, len(page_df.columns))])

    view_links = []
    for _, r in page_df.iterrows():
        rid = str(r.get(id_col, "")).strip()
        view_links.append(f'<a class="nimr-link" href="?id={_html.escape(rid)}">查看</a>' if rid else "-")
    page_df["操作"] = view_links

    display_df = page_df[show_cols + ["操作"]].copy()
    for c in show_cols:
        display_df[c] = display_df[c].map(lambda x: short_text(x, 80 if c == "属、种" else 60))

    center_cols = {id_col, "保藏日期", "操作"}
    table_html = _render_table_html(display_df, center_cols=center_cols)
    st.markdown(table_html, unsafe_allow_html=True)

    b1, b2, b3, b4 = st.columns([1, 1, 1, 1], gap="small", vertical_alignment="center")
    with b1:
        if st.button("⏮ 首页", use_container_width=True, disabled=(st.session_state.page == 1)):
            st.session_state.page = 1
            st.rerun()
    with b2:
        if st.button("◀ 上一页", use_container_width=True, disabled=(st.session_state.page == 1)):
            st.session_state.page -= 1
            st.rerun()
    with b3:
        if st.button("下一页 ▶", use_container_width=True, disabled=(st.session_state.page >= total_pages)):
            st.session_state.page += 1
            st.rerun()
    with b4:
        if st.button("末页 ⏭", use_container_width=True, disabled=(st.session_state.page >= total_pages)):
            st.session_state.page = total_pages
            st.rerun()

    st.markdown(
        f"""
        <div class="pager">
          <div class="pill">共 {total} 条记录</div>
          <div class="pill">第 {st.session_state.page} / {total_pages} 页（当前显示 {start+1 if total>0 else 0}-{min(end,total)}）</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------
# 详情页（左：KV；右：图片固定区）
# ---------------------------
def render_detail(df: pd.DataFrame, id_col: str, rid: str, excel_path: str):
    render_breadcrumb([("首页", False), ("资源目录", False), (f"详情：{rid}", True)])

    st.markdown(
        f'''
        <div class="nimr-section-title" style="display:flex;justify-content:space-between;align-items:center;gap:12px;">
          <div>资源详情</div>
          <a href="?" class="nimr-link"
             style="padding:7px 12px;border-radius:12px;background:var(--soft);border:1px solid var(--border2);">
             ← 返回资源目录
          </a>
        </div>
        ''',
        unsafe_allow_html=True,
    )

    hit = df[df[id_col].astype(str).str.strip() == rid.strip()]
    if hit.empty:
        st.warning(f"未找到记录：{id_col} = {rid}")
        return

    df_row_index = int(hit.index[0])
    row = hit.iloc[0].to_dict()

    img_col = detect_image_col(df)
    images = get_images_for_record(df, excel_path, df_row_index, img_col)

    left, right = st.columns([1.85, 1.0], gap="large", vertical_alignment="top")

    with left:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**基本信息**")
        exclude_cols = [img_col] if img_col else []
        st.markdown(_kv_html(df.columns.tolist(), row, exclude_cols), unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**菌种图片**")

        if images:
            for p in images:
                try:
                    st.image(p, use_container_width=True)
                except Exception:
                    st.warning(f"无法加载图片：{p}")
        else:
            st.info("未检测到图片：\n- 若 Excel 是“插入图片对象”，本程序会自动提取；\n- 若图片在列中以路径/URL存储，请确认可访问。")

        st.markdown("</div>", unsafe_allow_html=True)


def _kv_html(columns: List[str], row: dict, exclude_cols: List[str] = []) -> str:
    rows_html = []
    for c in columns:
        if c in exclude_cols:
            continue
        v = row.get(c, "")
        v = "" if v is None else str(v)
        if v.lower() in ("nan", "none"):
            v = ""

        is_parentheses_start = v.strip().startswith("(")
        v_html = (
            v.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace("\n", "<br/>")
        )
        if is_parentheses_start:
            v_html = f'<span style="color:#94a3b8;">{v_html}</span>'

        rows_html.append(
            f'<tr><td class="k">{_html.escape(str(c))}</td><td>{v_html if v_html else "&nbsp;"}</td></tr>'
        )
    return f'<table class="kv">{"".join(rows_html)}</table>'


# ---------------------------
# 主程序
# ---------------------------
def main():
    render_header()

    excel_path = pick_excel_path()
    df = load_excel(excel_path)
    id_col = detect_id_col(df)

    rid = get_query_id()
    if rid:
        render_detail(df, id_col, rid, excel_path)
    else:
        render_list(df, id_col)


if __name__ == "__main__":
    main()
