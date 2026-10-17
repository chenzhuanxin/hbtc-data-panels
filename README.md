<p align="center"><img src="assets/logo.png" width="120" alt="hbtc-data-panels"></p>

# 城市中国 · 湖北通城网 数据采集面板合集

面向 **城市中国 · 湖北通城网（hbtc.ccoo.cn）** 招聘后台的 Windows 桌面采集工具集。
均为 Python + Flask 本地服务 + 浏览器界面，已打包为**单文件 exe，双击即用**。

> 🌐 在线说明站：https://chenzhuanxin.github.io/hbtc-data-panels/
> 📦 下载 exe：https://github.com/chenzhuanxin/hbtc-data-panels/releases

---

## 包含三个子项目

| 子项目 | 目录 | 端口 | 一句话说明 |
|---|---|---|---|
| 城市中国平台数据面板 | [`city-data-panel/`](city-data-panel/) | 8787 | 抓取招聘后台 **7 个类目**全量数据，自动 OCR 识别电话图片，简历/岗位/企业/用户详情从后台编辑页补抓**全量字段**，按规范导出 Excel / ZIP |
| 企业投递简历采集面板 | [`resume-panel/`](resume-panel/) | 8790 | 按企业名称链式采集 **企业 → 岗位 → 投递简历 → 完整简历内容**，导出单表多页 Excel |
| 城市中国综合数据面板 | [`combined-panel/`](combined-panel/) | 8795 | 把上面两套工具**装进同一个 exe**：一个门户页 + 两个内置子面板，**一份 config.json 共享配置** |

每个子项目目录下的 `index.html` 是**详细说明页**，`README.md` 是简要说明，`src/` 是完整源码。

> **综合面板 vs. 两个独立面板**：功能完全一致，区别只在部署形态。
> 独立面板适合「只用一个功能」；综合面板适合「两个都要用」——只需开一个 exe、一个端口，后台 Cookie / 总后台 Cookie 与保存目录只填一次。

---

## v2.0 变化：详情抓取全面切换为后台编辑页

**不再抓前台网站，不再需要前台 Cookie**。所有详情字段都从后台 / 总后台的编辑页直接取：

| 数据 | 详情来源页面 | 所需 Cookie |
|---|---|---|
| 简历全量字段 | `job.admin.ccoo.cn/Post/JlEdit?source=1&jlid={jlid}` | 后台 Cookie |
| 岗位全量字段（全职 / 兼职） | `job.admin.ccoo.cn/Post/JobEdit?Source=1&id={jobid}` | 后台 Cookie |
| 企业全量字段 | `job.admin.ccoo.cn/Enterprise/CompanyAdd?source=1&companyId={companyId}` | 后台 Cookie |
| 用户全量字段 | `www.admin.ccoo.cn/yunying/users/user_set.asp?id={userId}` | **总后台 Cookie** |

带来的变化：

- **不需要前台 Cookie**（旧版必须登录前台 www.hbtc.ccoo.cn 才能抓简历正文）；
- **新增总后台 Cookie**：仅抓「注册用户」全量字段时需要，可留空（此时用户只有列表字段）；
- **速度大幅提升**：后台编辑页无前台 nginx WAF，不再需要令牌桶限速 / 请求间隔，全部全并发抓取；
- **OCR 仍然保留**：列表里的电话图片照常 base64 → RapidOCR 离线识别；
- **用户页面无需 OCR**：`user_set.asp` 是纯文本表格，直接解析（注册信息 / 个人信息 / 论坛信息 / 会员服务）。

---

## 快速开始

1. 打开 [Release 页面](https://github.com/chenzhuanxin/hbtc-data-panels/releases)，下载需要的 exe。
2. 把 exe 放到**一个单独的文件夹**里，双击运行。
3. 控制台会自动打开浏览器面板（8787 / 8790 / 8795），关闭控制台窗口即停止服务。
4. 在面板中填入后台账号密码或 Cookie，点「测试」验证后即可抓取、导出。
   若要拿到**注册用户全量字段**，再粘贴一份总后台（`www.admin.ccoo.cn`）的 Cookie 并点「测试总后台 Cookie」。

> exe 首次运行会在自身所在目录生成 `config.json` 与数据缓存目录，所有配置与数据都保存在本地，不上传任何服务器。

---

## 从源码运行

```bash
# 平台数据面板（端口 8787）
cd city-data-panel/src
pip install flask requests beautifulsoup4 lxml rapidocr_onnxruntime openpyxl onnxruntime
copy config.example.json config.json     # 填入后台 Cookie；可选总后台 Cookie
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

## 抓取机制（后台全量补抓）

| 机制 | 说明 |
|---|---|
| `fetch_workers` | 列表 / 后台详情页并发线程，默认 `8`。后台无 WAF，调大即提速 |
| 详情补抓 | 简历 JlEdit / 岗位 JobEdit / 企业 CompanyAdd / 用户 user_set.asp，抓完列表自动全并发补全 |
| 断点续抓 | 每个类目按锚点列（姓名 / 岗位标题 / 企业名称 / 家乡通ID）判断是否已有详情，「只补缺失」跳过已有、只抓空缺 |
| 失败自动补抓 | 一轮结束后自动再跑 2 轮；`need_login` 单独统计（Cookie 失效时明确提示） |
| 测试按钮 | 「测试连接」验后台 Cookie；「测试总后台 Cookie」用指定用户（如 79557050）试抓验证总后台 Cookie |
| 覆盖统计 | 补抓结束后面板直接报告「详情覆盖 120/200」这类数字 |

### 怎么确认"拿到了全量"

面板在补抓结束后会直接报告：**「详情覆盖 120/200」** 这样的数字。

- 覆盖数 = 总行数 → 该批详情已抓全（剩下为空的字段是**数据本身没填**，不是没抓到）；
- 覆盖数 < 总行数 → 点「**只补缺失**」续抓即可。

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
        ├── admin_detail.py  # 后台详情抓取模块（JlEdit / JobEdit / CompanyAdd / user_set.asp）
        └── static/          # 门户页
```

---

## 免责说明

本仓库工具仅供**持有合法后台账号**的使用者进行本单位数据的整理与备份。
源码中不含任何账号、Cookie 或凭据；仓库内的 `config.example.json` 仅为空白模板。
使用者需自行遵守目标平台的服务条款与相关法律法规。
