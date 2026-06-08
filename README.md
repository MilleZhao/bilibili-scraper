# B站通用筛选爬取平台

> 关键词搜索 → 分区浏览 → 一键爬取评论与弹幕。零外部依赖，一个 python server.py 即可启动。

[English](#english) | [简体中文](#简体中文)

---

## 简体中文

### 功能

- **通用搜索** — 关键词搜索 + 分区筛选 + 多种排序，跨全站或限定分区
- **一键爬取** — 勾选视频，自动采集视频信息、评论（支持页码）、弹幕（protobuf 解析）
- **Web 界面** — 本地浏览器打开，搜索/筛选/导出 CSV 全在网页完成
- **零依赖** — 只用 Python 标准库，不需要 pip install 任何东西

### 快速开始

`ash
# 1. 克隆仓库
git clone https://github.com/你的用户名/bilibili-scraper.git
cd bilibili-scraper

# 2. 配置 B站 Cookie（选一个方式）
python setup.py              # 交互式引导（推荐）
# 或手动创建 my_cookies.py（参考 my_cookies.example.py）
# 或设置环境变量: SESSDATA, BILI_JCT, DEDEUSERID, BUVID3

# 3. 启动
python server.py
# 打开浏览器访问 http://127.0.0.1:8765
`

### Cookie 获取方法

1. 打开 https://www.bilibili.com 并登录
2. 按 F12 → Application → Cookies → bilibili.com
3. 找到这四个值填入：
   - SESSDATA — 登录态核心（必填）
   - ili_jct — CSRF Token（必填）
   - DedeUserID — 用户 ID（选填）
   - uvid3 — 设备标识（选填）

> ⚠️ **Cookie 是敏感信息，不要提交到 Git 或分享给他人。** 仓库已配置 .gitignore 排除 my_cookies.py。

### 项目结构

`
bilibili-scraper/
├── server.py                 # Web 服务（多线程 HTTP Server）
├── bilibili_search_engine.py # 搜索引擎（Wbi 签名 + 分类检索）
├── bilibili_scraper.py       # 爬虫引擎（评论 + 弹幕 protobuf 解析）
├── index.html                # 前端界面（单页应用）
├── setup.py                  # Cookie 配置向导
├── my_cookies.example.py     # Cookie 配置模板
├── outputs/                  # 爬取结果输出目录
└── 踩坑日志.md               # 开发排错记录
`

### 技术亮点

| 特性 | 实现 |
|------|------|
| B站 Wbi 签名 | 纯 Python 实现，无需调用第三方 |
| 弹幕解析 | 手写 protobuf 解析器，零依赖 |
| 评论翻页 | cursor 分页，支持设定最大页码 |
| 前后端通信 | 轮询事件模型，避免 SSE 阻塞问题 |
| Web 服务器 | http.server + ThreadingMixIn，多线程处理 |
| 前端 | 原生 JavaScript，无框架 |

### 截图

> 📸 运行 python server.py 后打开 http://127.0.0.1:8765 即可看到界面。

*(TODO: 添加截图)*

### 踩坑日志

开发过程中遇到的关键问题及解决方案记录在 [踩坑日志.md](./踩坑日志.md) 中（中文）。

---

## English

### Features

- **Universal Search** — Keyword search + category filter + multiple sort modes across all Bilibili categories
- **One-Click Scraping** — Select videos, auto-collect video info, comments (paginated), and danmaku (protobuf parsing)
- **Web UI** — Full search/filter/export CSV workflow in browser
- **Zero Dependencies** — Python standard library only, no pip install needed

### Quick Start

`ash
git clone https://github.com/yourname/bilibili-scraper.git
cd bilibili-scraper
python setup.py      # Interactive cookie setup
python server.py     # Start server, open http://127.0.0.1:8765
`

### Tech Stack

- **Backend:** Python stdlib http.server + ThreadingMixIn
- **Frontend:** Vanilla JavaScript (no frameworks)
- **Bilibili API:** Wbi signing, cursor-based comment pagination, custom protobuf danmaku parser
- **Architecture:** Polling-based event model (avoids SSE blocking issues with single-threaded server)

### Disclaimer

This project is for **educational and research purposes only**. Users are responsible for complying with Bilibili's Terms of Service. Do not use for commercial purposes or at scales that may impact Bilibili's services.

### License

MIT License — see LICENSE file for details.
