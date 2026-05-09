import os
import glob
import time
import requests
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from datetime import datetime
from docx import Document

# ==================== 配置 ====================
TOKEN = os.environ.get("PUSHPLUS_TOKEN", "")
REPO = os.environ.get("GITHUB_REPOSITORY", "")
TOPIC = os.environ.get("PUSHPLUS_TOPIC", "")
DATA_DIR = "data"
OUTPUT_DIR = "output"

# ==================== 工具函数 ====================
def get_latest_file(exts):
    """抓取 data/ 目录下指定扩展名的最新文件"""
    files = []
    for ext in exts:
        pattern = os.path.join(DATA_DIR, "*." + ext)
        matched = glob.glob(pattern)
        files.extend(matched)
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def read_text(path):
    """读取文字描述，支持 txt/md/docx"""
    if not path or not os.path.exists(path):
        return "今日无文字描述。"
    if path.endswith(".docx"):
        doc = Document(path)
        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return "\n".join(lines)
    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip()

def generate_charts(excel_path):
    """为Excel每个Sheet生成一张图，标题即Sheet名"""
    date_str = datetime.now().strftime("%Y-%m-%d")
    out_dir = os.path.join(OUTPUT_DIR, date_str)
    os.makedirs(out_dir, exist_ok=True)

    xl = pd.ExcelFile(excel_path)
    charts = []

    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet)
            df = df.dropna(how="all").dropna(axis=1, how="all")
            if df.empty:
                continue

            safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in sheet).strip("_")
            if not safe_name:
                safe_name = "chart_" + str(len(charts))
            png_path = os.path.join(out_dir, safe_name + ".png")

            fig, ax = plt.subplots(figsize=(8, 3.5), dpi=100)
            plotted = False

            # ---------- 启发1：行业分布类（涨停行业分布）----------
            if "行业" in sheet or "分布" in sheet:
                try:
                    cols = [c for c in df.columns if df[c].notna().any()]
                    if len(cols) >= 2:
                        x_col, y_col = cols[0], cols[1]
                        plot_df = df[[x_col, y_col]].dropna().head(15)
                        plot_df[y_col] = pd.to_numeric(plot_df[y_col], errors='coerce')
                        plot_df = plot_df.dropna()
                        if len(plot_df) > 0:
                            ax.barh(range(len(plot_df)), plot_df[y_col].values, color='#c23531')
                            ax.set_yticks(range(len(plot_df)))
                            labels = [str(x)[:12] for x in plot_df[x_col].values]
                            ax.set_yticklabels(labels, fontsize=8)
                            ax.invert_yaxis()
                            plotted = True
                except Exception as e:
                    print(f"[{sheet}] 行业分布绘图失败: {e}")

            # ---------- 启发2：宽格式（连板个股、大盘小盘&成长价值）----------
            if not plotted and (len(df) <= 10 and len(df.columns) > 10):
                try:
                    idx_col = df.columns[0]
                    df_t = df.set_index(idx_col).T
                    # 尝试解析索引为日期（Excel序列号或字符串）
                    try:
                        numeric_idx = pd.to_numeric(df_t.index, errors='coerce')
                        df_t.index = pd.to_datetime(numeric_idx, unit='D', origin='1899-12-30', errors='coerce')
                    except Exception:
                        df_t.index = pd.to_datetime(df_t.index, errors='coerce')
                    df_t = df_t.dropna()
                    
                    num_cols = [c for c in df_t.columns if pd.api.types.is_numeric_dtype(df_t[c])]
                    for col in num_cols[:4]:
                        ax.plot(df_t.index, df_t[col].astype(float), label=str(col), linewidth=1.2)
                    if num_cols:
                        ax.legend(fontsize=7, loc='best')
                        if len(df_t) > 30:
                            ax.xaxis.set_major_locator(plt.MaxNLocator(10))
                        plotted = True
                except Exception as e:
                    print(f"[{sheet}] 宽格式绘图失败: {e}")

            # ---------- 启发3：长格式（涨停数量等）----------
            if not plotted:
                date_col = None
                for col in df.columns:
                    col_str = str(col).lower()
                    if any(k in col_str for k in ["date", "日期", "时间", "day"]):
                        date_col = col
                        break

                if date_col is not None:
                    try:
                        # 混合类型日期处理：先尝试Excel序列号，再尝试字符串
                        numeric = pd.to_numeric(df[date_col], errors='coerce')
                        parsed = pd.to_datetime(numeric, unit='D', origin='1899-12-30', errors='coerce')
                        mask = parsed.isna()
                        if mask.any():
                            parsed[mask] = pd.to_datetime(df.loc[mask, date_col], errors='coerce')
                        df[date_col] = parsed
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
                        print(f"[{sheet}] 长格式绘图失败: {e}")

            # ---------- 启发4：兜底——前两列简单画图 ----------
            if not plotted:
                try:
                    cols = [c for c in df.columns if df[c].notna().any()]
                    if len(cols) >= 2:
                        x_col, y_col = cols[0], cols[1]
                        plot_df = df[[x_col, y_col]].dropna().head(20)
                        plot_df[y_col] = pd.to_numeric(plot_df[y_col], errors='coerce')
                        plot_df = plot_df.dropna()
                        if len(plot_df) > 0:
                            ax.bar(range(len(plot_df)), plot_df[y_col].values, color='#2f4554')
                            ax.set_xticks(range(len(plot_df)))
                            ax.set_xticklabels([str(v)[:8] for v in plot_df[x_col].values], rotation=45, ha='right', fontsize=7)
                            plotted = True
                except Exception as e:
                    print(f"[{sheet}] 兜底绘图失败: {e}")

            if not plotted:
                ax.text(0.5, 0.5, "数据格式暂不支持自动绘图", ha='center', va='center', transform=ax.transAxes)

            ax.set_title(sheet, fontsize=11)
            plt.tight_layout()
            fig.savefig(png_path, dpi=100, bbox_inches='tight')
            plt.close(fig)
            charts.append((sheet, safe_name))
            print(f"已生成图表: {png_path}")

        except Exception as e:
            print(f"[{sheet}] Sheet处理异常: {e}")
            plt.close('all')

    return charts, date_str

def git_commit():
    """把生成的图片提交回仓库，以便生成Raw URL"""
    os.system("git config user.name 'github-actions[bot]'")
    os.system("git config user.email '41898282+github-actions[bot]@users.noreply.github.com'")
    os.system("git add " + OUTPUT_DIR + "/")
    os.system("git commit -m 'daily: auto charts' || echo 'No changes to commit'")
    os.system("git push")

def push_message(text_content, charts, date_str):
    """Pushplus推送：文字 + Markdown图片链接"""
    if not TOKEN:
        print("未设置 PUSHPLUS_TOKEN，跳过推送")
        return

    owner_repo = REPO if REPO else "你的用户名/你的仓库名"
    base_url = "https://raw.githubusercontent.com/" + owner_repo + "/main/output/" + date_str

    md_lines = [text_content, "", "---", ""]
    for sheet_name, safe_name in charts:
        img_url = base_url + "/" + safe_name + ".png"
        md_lines.append("**" + sheet_name + "**")
        md_lines.append("![" + sheet_name + "](" + img_url + ")")
        md_lines.append("")

    payload = {
        "token": TOKEN,
        "title": date_str + " A股每日复盘",
        "content": "\n".join(md_lines),
        "template": "markdown"
    }
    if TOPIC:
        payload["topic"] = TOPIC

    res = requests.post("http://www.pushplus.plus/send", data=payload, timeout=20)
    print("Pushplus响应:", res.json())

# ==================== 主流程 ====================
if __name__ == "__main__":
    excel = get_latest_file(["xlsx", "xls"])
    text = get_latest_file(["txt", "md", "docx"])

    if not excel:
        raise FileNotFoundError("请在 data/ 目录下上传当日的Excel文件")

    print("读取Excel:", excel)
    print("读取文字:", text)

    text_content = read_text(text)
    charts, date_str = generate_charts(excel)

    if charts:
        git_commit()
        print("图片已提交，等待GitHub CDN刷新...")
        time.sleep(45)
        push_message(text_content, charts, date_str)
    else:
        print("未生成任何图表，跳过推送")
