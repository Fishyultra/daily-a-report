import os
import glob
import time
import requests
import pandas as pd
import matplotlib
matplotlib.use('Agg')               # 无GUI服务器环境必需
import matplotlib.pyplot as plt
from datetime import datetime
from docx import Document

# ==================== 配置 ====================
TOKEN = os.environ.get("PUSHPLUS_TOKEN")
REPO  = os.environ.get("GITHUB_REPOSITORY", "")   # owner/repo
TOPIC = os.environ.get("PUSHPLUS_TOPIC", "")      # 如需发群组，填topic编码
DATA_DIR = "data"
OUTPUT_DIR = "output"

# ==================== 工具函数 ====================
def get_latest_file(exts):
    """抓取 data/ 目录下指定扩展名的最新文件"""
    files = []
    for ext in exts:
        files.extend(glob.glob(os.path.join(DATA_DIR, f"*.{ext})))
    return max(files, key=os.path.getmtime) if files else None

def read_text(path):
    """读取文字描述，支持 txt/md/docx"""
    if not path:
        return "今日无文字描述。"
    if path.endswith(".docx"):
        doc = Document(path)
        return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip()

def generate_charts(excel_path):
    """为Excel每个Sheet生成一张图，标题即Sheet名"""
    date_str = datetime.now().strftime("%Y-%m-%d")
    out_dir = os.path.join(OUTPUT_DIR, date_str)
    os.makedirs(out_dir, exist_ok=True)

    xl = pd.ExcelFile(excel_path)
    charts = []   # [(sheet_name, safe_name), ...]

    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if df.empty:
            continue

        # 安全文件名
        safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in sheet).strip("_")
        png_path = os.path.join(out_dir, f"{safe_name}.png")

        fig, ax = plt.subplots(figsize=(8, 3.5), dpi=100)
        plotted = False

        # ---------- 启发1：宽格式（如"连板个股"，行少列多，第一列是指标名） ----------
        if len(df) <= 6 and len(df.columns) > 10:
            try:
                idx_col = df.columns[0]
                df_t = df.set_index(idx_col).T
                # 列名是Excel日期序列号，转日期
                df_t.index = pd.to_datetime(df_t.index, unit='D', origin='1899-12-30', errors='coerce')
                df_t = df_t.dropna()
                num_cols = [c for c in df_t.columns if pd.api.types.is_numeric_dtype(df_t[c])]
                for col in num_cols[:3]:
                    ax.plot(df_t.index, df_t[col].astype(float), label=str(col), linewidth=1.2)
                if num_cols:
                    ax.legend(fontsize=8)
                    plotted = True
            except Exception as e:
                print(f"[{sheet}] 宽格式解析失败: {e}")

        # ---------- 启发2：长格式，找日期/时间列 ----------
        if not plotted:
            date_col = None
            for col in df.columns:
                if any(k in str(col).lower() for k in ["date", "日期", "时间", "day"]):
                    date_col = col
                    break

            if date_col is not None:
                try:
                    if pd.api.types.is_numeric_dtype(df[date_col]):
                        df[date_col] = pd.to_datetime(df[date_col], unit='D', origin='1899-12-30', errors='coerce')
                    else:
                        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                    df = df.dropna(subset=[date_col])
                    num_cols = [c for c in df.select_dtypes(include='number').columns if c != date_col]
                    for col in num_cols[:4]:
                        ax.plot(df[date_col], df[col], label=str(col), linewidth=1.2)
                    if num_cols:
                        ax.legend(fontsize=7, loc='best')
                        if len(df) > 30:
                            ax.xaxis.set_major_locator(plt.MaxNLocator(10))
                        plotted = True
                except Exception as e:
                    print(f"[{sheet}] 日期列解析失败: {e}")

        # ---------- 启发3：类别型（如"涨停行业分布"，第一列文本，后面数值） ----------
        if not plotted:
            try:
                x_col = df.columns[0]
                y_col = None
                for c in df.columns[1:]:
                    if pd.api.types.is_numeric_dtype(df[c]):
                        y_col = c
                        break
                if y_col is not None:
                    plot_df = df.head(15)
                    vals = plot_df[y_col].astype(float)
                    ax.barh(range(len(plot_df)), vals, color='#c23531')
                    ax.set_yticks(range(len(plot_df)))
                    labels = [str(x)[:12] for x in plot_df[x_col]]
                    ax.set_yticklabels(labels, fontsize=8)
                    ax.invert_yaxis()
                    plotted = True
            except Exception as e:
                print(f"[{sheet}] 类别图解析失败: {e}")

        if not plotted:
            ax.text(0.5, 0.5, "数据格式暂不支持自动绘图", ha='center', va='center', transform=ax.transAxes)

        ax.set_title(sheet, fontsize=11)
        plt.tight_layout()
        fig.savefig(png_path, dpi=100, bbox_inches='tight')
        plt.close(fig)
        charts.append((sheet, safe_name))
        print(f"已生成图表: {png_path}")

    return charts, date_str

def git_commit():
    """把生成的图片提交回仓库，以便生成Raw URL"""
    os.system("git config user.name 'github-actions[bot]'")
    os.system("git config user.email '41898282+github-actions[bot]@users.noreply.github.com'")
    os.system(f"git add {OUTPUT_DIR}/")
    os.system("git commit -m 'daily: auto charts' || echo 'No changes to commit'")
    os.system("git push")

def push_message(text_content, charts, date_str):
    """Pushplus推送：文字 + Markdown图片链接"""
    if not TOKEN:
        print("未设置 PUSHPLUS_TOKEN，跳过推送")
        return

    owner_repo = REPO if REPO else "你的用户名/你的仓库名"
    base_url = f"https://raw.githubusercontent.com/{owner_repo}/main/output/{date_str}"

    # 组装Markdown
    md = f"{text_content}\n\n---\n\n"
    for sheet_name, safe_name in charts:
        img_url = f"{base_url}/{safe_name}.png"
        md += f"**{sheet_name}**\n![{sheet_name}]({img_url})\n\n"

    payload = {
        "token": TOKEN,
        "title": f"{date_str} A股每日复盘",
        "content": md,
        "template": "markdown"
    }
    if TOPIC:
        payload["topic"] = TOPIC

    res = requests.post("http://www.pushplus.plus/send", data=payload, timeout=20)
    print("Pushplus响应:", res.json())

# ==================== 主流程 ====================
if __name__ == "__main__":
    excel = get_latest_file(["xlsx", "xls"])
    text  = get_latest_file(["txt", "md", "docx"])

    if not excel:
        raise FileNotFoundError(f"请在 {DATA_DIR}/ 目录下上传当日的Excel文件")

    print(f"读取Excel: {excel}")
    print(f"读取文字: {text}")

    text_content = read_text(text)
    charts, date_str = generate_charts(excel)

    if charts:
        git_commit()
        print("图片已提交，等待GitHub CDN刷新...")
        time.sleep(45)                      # 关键：等待Raw链接生效
        push_message(text_content, charts, date_str)
    else:
        print("未生成任何图表，跳过推送")
