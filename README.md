# B站数据清洗与报告生成

> 评论清洗 → 弹幕去重 → 数据可视化 → Word/Excel 报告一键生成

## 功能

| 脚本 | 功能 |
|------|------|
| clean_comments.py | 评论清洗：去无效、按点赞排序 |
| export_csv.py | 评论导出 CSV |
| export_danmaku.py | 弹幕去重 + 排序 + 导出 Excel |
| 	o_excel.py | 一键 Excel：读取→清洗→排序→导出 |
| uild_report.py | 生成 Word 评论分析报告 |
| uild_full_report.py | 完整综合分析报告（含 matplotlib 图表） |
| uild_docx.py | DOCX 构建工具函数 |
| uild_xlsx.mjs | Excel 构建工具（Node.js） |

## 依赖安装

`ash
pip install python-docx matplotlib openpyxl
`

## 使用

`ash
# 清洗评论
python clean_comments.py

# 弹幕去重 + 导出
python export_danmaku.py

# 生成分析报告
python build_full_report.py
`

## 数据来源

数据由 [B站通用筛选爬取平台](https://github.com/MilleZhao/bilibili-scraper) 采集，输出至 outputs/ 目录后，本项目读取并处理。

## 项目结构

`
Project003-数据清洗与报告生成/
├── clean_comments.py        # 评论清洗
├── export_csv.py            # CSV 导出
├── export_danmaku.py        # 弹幕去重导出
├── to_excel.py              # 一键 Excel
├── build_report.py          # Word 报告生成
├── build_full_report.py     # 完整分析报告
├── build_docx.py            # DOCX 工具
├── build_xlsx.mjs           # XLSX 工具
├── outputs/                 # 输出目录（gitignored）
└── bilibili-scraper-skill/  # Codex skill 插件
`

## 免责声明

本工具仅供学习研究使用。使用者须遵守 Bilibili 用户协议，不得将数据用于商业用途。

## 许可证

MIT License
