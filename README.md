# 城市中国 · 湖北通城网 数据采集面板合集

面向 **城市中国 · 湖北通城网（hbtc.ccoo.cn）** 招聘后台的 Windows 桌面采集工具集。
均为 Python + Flask 本地服务 + 浏览器界面，已打包为**单文件 exe，双击即用**。

> 🌐 在线说明站：https://chenzhuanxin.github.io/hbtc-data-panels/
> 📦 下载 exe：https://github.com/chenzhuanxin/hbtc-data-panels/releases

---

## 包含三个子项目

| 子项目 | 目录 | 端口 | 一句话说明 |
|---|---|---|---|
| 城市中国平台数据面板 | [`city-data-panel/`](city-data-panel/) | 8787 | 抓取招聘后台 **7 个类目**全量数据，自动 OCR 识别电话图片，并可补抓**简历全量字段**（91 → 116 列），按规范导出 Excel / ZIP |
| 企业投递简历采集面板 | [`resume-panel/`](resume-panel/) | 8790 | 按企业名称链式采集 **企业 → 岗位 → 投递简历 → 完整简历内容**，导出单表多页 Excel |
| 城市中国综合数据面板 | [`combined-panel/`](combined-panel/) | 8795 | 把上面两套工具**装进同一个 exe**：一个门户页 + 两个内置子面板，**一份 config.json 共享配置** |

每个子项目目录下的 `index.html` 是**详细说明页**，`README.md` 是简要说明，`src/` 是完整源码。

> **综合面板 vs. 两个独立面板**：功能完全一致，区别只在部署形态。
> 独立面板适合「只用一个功能」；综合面板适合「两个都要用」——只需开一个 exe、一个端口，两种 Cookie 与保存目录只填一次。

> **关于「完整简历内容」**：简历的完整正文（工作经历、教育经历、语言能力、获得证书）只存在于前台简历详情页
> `http://www.hbtc.ccoo.cn/post/qiuzhi/{id}x.html`，该页**必须带前台站点登录 Cookie**，否则返回登录页。
> 所有面板都把「后台 Cookie」与「前台 Cookie」分开配置，各管一段。

---

## ⚠️ 反爬限速（务必先读）

前台简历详情页前方是 **nginx WAF**。实测：**8 线程连抓 200 份时，大约第 120 份开始返回带滑动验证的 `403`**
（页面文案「疑似异常访问请求，已被系统拦截 / IP频繁访问」），之后**本机 IP 会被临时拉黑**，连正常请求也一起 403。

对比手工脚本（单线程 + 每份 `time.sleep(1)`）长期可用，可知瓶颈是**速率而不是并发数**。因此所有面板内置了限速：

| 参数 | 默认 | 说明 |
|---|---|---|
| `front_rate` | `1.0` | 任意两次前台简历请求的最小间隔（秒），全局令牌桶强制生效 |
| `fetch_workers` | `4` | 抓取线程（1–8）。受 `front_rate` 约束，**调大不会更快**，只会更容易触发拦截 |

- **在途并发上限 2**：同时最多 2 个连接打到站点。
- **命中即熔断**：连续 3 次被 WAF 拦截就自动停止剩余任务，并给出提示，避免 IP 进入更长时间的黑名单。
- **断点续抓**：已抓到的详情会回写缓存，被拦截后隔几分钟重抓即可接着来，不重复不丢失。
- **预期耗时**：按 1 请求/秒算，40 份约 1 分钟、500 份约 8 分钟、15,733 份全量约 4–5 小时。

> 列表类接口（7 个类目、企业搜索、岗位、投递明细）走的是后台 `job.admin.ccoo.cn`，**不受前台限速影响**，仍是并发抓取。

---

## 快速开始

1. 打开 [Release 页面](https://github.com/chenzhuanxin/hbtc-data-panels/releases)，下载需要的 exe：
   - `CityDataPanel-v1.2.0-windows-x64.exe` — 平台数据面板
   - `ResumeCollector-v1.2.0-windows-x64.exe` — 企业投递简历采集面板
   - `CombinedPanel-v1.2.0-windows-x64.exe` — 综合面板（两套工具合一）
2. 把 exe 放到**一个单独的文件夹**里，双击运行。
3. 控制台会自动打开浏览器面板（8787 / 8790 / 8795），关闭控制台窗口即停止服务。
4. 在面板中填入后台账号密码或 Cookie，点「测试」验证后即可抓取、导出。
   若要拿到**完整简历内容**，再粘贴一份前台站点（`www.hbtc.ccoo.cn`）的 Cookie 并点「测试前台 Cookie」。

> exe 首次运行会在自身所在目录生成 `config.json` 与数据缓存目录，所有配置与数据都保存在本地，不上传任何服务器。

---

## 从源码运行

```bash
# 平台数据面板（端口 8787）
cd city-data-panel/src
pip install flask requests beautifulsoup4 lxml rapidocr_onnxruntime openpyxl onnxruntime
copy config.example.json config.json     # 填入后台 Cookie / 前台 Cookie
python -X utf8 面板.py

# 企业投递简历采集面板（端口 8790）
cd resume-panel/src
pip install flask requests beautifulsoup4 lxml openpyxl
copy config.example.json config.json
python -X utf8 面板.py

# 综合面板（端口 8795，门户 + /city/ + /jd/）
cd combined-panel/src
pip install flask requests beautifulsoup4 lxml rapidocr_onnxruntime openpyxl onnxruntime
copy config.example.json config.json
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
├── city-data-panel/        # 子项目一：7 类平台数据
│   ├── index.html          # 详细说明页
│   ├── README.md
│   └── src/
├── resume-panel/           # 子项目二：企业投递简历
│   ├── index.html
│   ├── README.md
│   └── src/
└── combined-panel/         # 子项目三：综合面板
    ├── index.html
    ├── README.md
    └── src/
        ├── 面板.py          # 门户 + DispatcherMiddleware 挂载两个子面板
        ├── panels/city/     # 平台数据子面板（panel.py + static/）
        ├── panels/jd/       # 企业投递子面板（panel.py + static/）
        ├── resume_front.py  # 前台简历抓取模块（含反爬限速）
        └── static/          # 门户页
```

---

## 免责说明

本仓库工具仅供**持有合法后台账号**的使用者进行本单位数据的整理与备份。
源码中不含任何账号、Cookie 或凭据；仓库内的 `config.example.json` 仅为空白模板。
使用者需自行遵守目标平台的服务条款与相关法律法规。
