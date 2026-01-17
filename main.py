import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import font_manager, rcParams
import numpy as np
import argparse
import re

parser = argparse.ArgumentParser()
parser.add_argument("--column", help="集計する列名")
parser.add_argument("--title", help="グラフタイトル")
parser.add_argument("--input", default="data.xlsx", help="入力Excelファイル")
parser.add_argument("--auto", action="store_true", help="全列を自動で円グラフ化")
parser.add_argument("--graph", choices=["barh", "pie"], default="barh", help="グラフ形式（barh: 横棒, pie: 円グラフ）")
parser.add_argument("--hide-threshold", type=float, default=1.0, help="注釈を非表示にする割合未満")
parser.add_argument("--show-threshold", type=float, default=1.0, help="％表示を行う割合以上")
args = parser.parse_args()

HIDE_THRESHOLD = args.hide_threshold
SHOW_THRESHOLD = args.show_threshold

input_file = args.input

font_path = "C:/Windows/Fonts/meiryo.ttc"
font_prop = font_manager.FontProperties(fname=font_path)
rcParams["font.family"] = font_prop.get_name()

df = pd.read_excel(input_file)


def safe_filename(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", text)

def is_datetime_column(series: pd.Series) -> bool:
    return pd.api.types.is_datetime64_any_dtype(series)

def draw_pie(column_name: str, title_text: str):
    counts = df[column_name].value_counts()
    if counts.empty:
        print(f"スキップ（データなし）: {column_name}")
        return

    total = counts.sum()
    ratios = counts / total * 100

    _, ax = plt.subplots(figsize=(8, 8))

    wedges = ax.pie(
        counts,
        labels=None,
        autopct=lambda p: f"{p:.1f}%" if p >= SHOW_THRESHOLD else "",
        startangle=90,
        counterclock=False,
    )[0]

    ax.axis("equal")
    ax.set_title(title_text, fontsize=14, pad=28)

    used_y_positions = []
    MAX_Y = 1.05
    MIN_GAP = 0.08

    for i, wedge in enumerate(wedges):
        pct = ratios.iloc[i]
        if pct < HIDE_THRESHOLD:
            continue

        angle = (wedge.theta1 + wedge.theta2) / 2
        x = np.cos(np.deg2rad(angle))
        y = np.sin(np.deg2rad(angle))

        y_text = y * 1.15

        while any(abs(y_text - used_y) < MIN_GAP for used_y in used_y_positions):
            y_text += -MIN_GAP if y_text > 0 else MIN_GAP

        if y_text > MAX_Y:
            y_text = MAX_Y

        used_y_positions.append(y_text)

        ax.annotate(
            counts.index[i],
            xy=(x * 0.7, y * 0.7),
            xytext=(x * 1.3, y_text),
            ha="left" if x >= 0 else "right",
            va="center",
            arrowprops=dict(arrowstyle="-", lw=1.2),
        )

    filename = f"{safe_filename(title_text)}.png"
    plt.savefig(filename, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"出力しました: {filename}")

def draw_barh(column_name: str, title_text: str):
    counts = df[column_name].value_counts()
    if counts.empty:
        print(f"スキップ（データなし）: {column_name}")
        return

    total = counts.sum()
    ratios = counts / total * 100

    _, ax = plt.subplots(figsize=(10, max(4, len(counts) * 0.4)))

    y = np.arange(len(counts))

    ax.set_title(title_text, fontsize=14, pad=20)
    ax.set_xlabel("割合 (%)")

    ax.set_yticks(y)
    ax.set_yticklabels(counts.index, fontsize=9)

    ax.barh(y, ratios, color="tab:blue", zorder=2)

    max_ratio = max(ratios)
    ax.set_xlim(0, max_ratio * 1.3)
    LABEL_OFFSET = max_ratio * 0.01

    for i, pct in enumerate(ratios):
        ax.text(
            pct + LABEL_OFFSET,
            y[i],
            f"{pct:.1f}%",
            ha="left",
            va="center",
            fontsize=8,
            backgroundcolor="white",
            zorder=3,
        )

    ax.invert_yaxis()

    filename = f"{safe_filename(title_text)}.png"
    plt.savefig(filename, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"出力しました: {filename}")

def draw_graph(column_name: str, title_text: str):
    if args.graph == "pie":
        draw_pie(column_name, title_text)
    else:
        draw_barh(column_name, title_text)

if args.auto:
    for col in df.columns:
        if is_datetime_column(df[col]):
            continue
        draw_graph(col, col)
else:
    if not args.column:
        raise ValueError("--column を指定するか --auto を使用してください")

    title = args.title if args.title else args.column
    draw_graph(args.column, title)
