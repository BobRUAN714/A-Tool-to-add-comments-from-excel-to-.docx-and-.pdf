import os
from pathlib import Path

# 基础路径设置：获取当前项目根目录
BASE_DIR = Path(__file__).resolve().parent

# 1. 输入文件所在文件夹（Excel 和作业都在这里，也可自行分开）
INPUT_FOLDER = BASE_DIR / "homeworks"

# 2. 评语 Excel 的文件名
EXCEL_FILENAME = "comments.xlsx"
EXCEL_PATH = INPUT_FOLDER / EXCEL_FILENAME

# 3. 处理结果保存位置
OUTPUT_FOLDER = BASE_DIR / "processed"

# --- PDF 字体设置（防止中文乱码） ---
# Windows 系统默认黑体路径
FONT_PATH = 'C:/Windows/Fonts/simhei.ttf'