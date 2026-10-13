# 城市中国 · 湖北通城网 数据采集面板合集

面向 **城市中国 · 湖北通城网（hbtc.ccoo.cn）** 招聘后台的两款 Windows 桌面采集面板。
均为 Python + Flask 本地服务 + 浏览器界面，已打包为**单文件 exe，双击即用**。

> 🌐 在线说明站：https://chenzhuanxin.github.io/hbtc-data-panels/
> 📦 下载 exe：https://github.com/chenzhuanxin/hbtc-data-panels/releases

---

## 包含两个子项目

| 子项目 | 目录 | 端口 | 一句话说明 |
|---|---|---|---|
| 城市中国平台数据面板 | [`city-data-panel/`](city-data-panel/) | 8787 | 抓取招聘后台 **7 个类目**全量数据，自动 OCR 识别电话图片，并可补抓**简历全量字段**（91 → 116 列），按规范导出 Excel / ZIP |
| 企业投递简历采集面板 | [`resume-panel/`](resume-panel/) | 8790 | 按企业名称链式采集 **企业 → 岗位 → 投递简历 → 完整简历内容**，导出单表多页 Excel |

每个子项目目录下的 `index.html` 是**详细说明页**，`README.md` 是简要说明，`src/` 是完整源码。

> **关于「完整简历内容」**：简历的完整正文（工作经历、教育经历、语言能力、获得证书）只存在于前台简历详情页
> `http://www.hbtc.ccoo.cn/post/qiuzhi/{id}x.html`，该页**必须带前台站点登录 Cookie**，否则返回登录页。
> 两个面板都把「后台 Cookie」与「前台 Cookie」分开配置，各管一段。

---

## 快速开始

1. 打开 [Release 页面](https://github.com/chenzhuanxin/hbtc-data-panels/releases)，下载需要的 exe：
   - `CityDataPanel-v1.1.0-windows-x64.exe` — 平台数据面板
   - `ResumeCollector-v1.1.0-windows-x64.exe` — 企业投递简历采集面板
2. 把 exe 放到**一个单独的文件夹**里，双击运行。
3. 控制台会自动打开浏览器面板（8787 / 8790），关闭控制台窗口即停止服务。
4. 在面板中填入后台账号密码或 Cookie，点「测试」验证后即可抓取、导出。
   若要拿到**完整简历内容**，再粘贴一份前台站点（`www.hbtc.ccoo.cn`）的 Cookie 并点「测试前台 Cookie」。

> exe 首次运行会在自身所在目录生成 `config.json` 与 `data_cache/`，所有配置与数据都保存在本地，不上传任何服务器。

---

## 从源码运行

```bash
# 平台数据面板
cd city-data-panel/src
pip install flask requests beautifulsoup4 lxml rapidocr_onnxruntime openpyxl onnxruntime
copy config.example.json config.json     # 填入后台 Cookie / 前台 Cookie
python -X utf8 面板.py

# 企业投递简历采集面板
cd resume-panel/src
pip install flask requests beautifulsoup4 lxml openpyxl
copy config.example.json config.json     # 填入后台 Cookie / 前台 Cookie
python -X utf8 面板.py
```

Windows 下也可以直接双击各项目 `src/启动面板.bat`。

---

## 重新打包 exe

```bash
pip install pyinstaller
cd <项目>/src
python 打包exe.py
```

产物在 `src/dist/` 下。exe 体积约 110 MB，是因为把 OCR 的 onnxruntime 与模型文件一并打包了进去。

---

## 目录结构

```
.
├── index.html              # 说明站首页（GitHub Pages）
├── README.md               # 本文件
├── .nojekyll               # 关闭 Jekyll，保证 Pages 正常发布
├── city-data-panel/        # 子项目一
│   ├── index.html          # 详细说明页
│   ├── README.md
│   └── src/                # 源码（面板.py / static / 打包脚本 / 图标）
└── resume-panel/           # 子项目二
    ├── index.html
    ├── README.md
    └── src/
```

---

## 免责说明

本仓库工具仅供**持有合法后台账号**的使用者进行本单位数据的整理与备份。
源码中不含任何账号、Cookie 或凭据；仓库内的 `config.example.json` 仅为空白模板。
使用者需自行遵守目标平台的服务条款与相关法律法规。
