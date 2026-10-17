# 城市中国综合数据面板

> 城市中国 · 湖北通城网 招聘频道（hbtc.ccoo.cn）| **两套工具装进同一个 exe**
> 在线详细说明：https://chenzhuanxin.github.io/hbtc-data-panels/combined-panel/

本地网页面板（`http://127.0.0.1:8795`）：一个门户页 + 两个内置子面板，**一份 `config.json` 共享配置**。

---

## 三个入口

| 入口 | 等价于 | 说明 |
|---|---|---|
| `/` | 门户导航页 | 两个工具的入口卡片，显示各自配置状态 |
| `/city/` | [城市中国平台数据面板](../city-data-panel/) | 7 类目全量抓取 + 电话图片 OCR + 简历/岗位/企业/用户全量字段补抓 |
| `/jd/` | [企业投递简历采集面板](../resume-panel/) | 企业 → 岗位 → 投递简历 → 完整简历内容 |

功能与两个独立面板**完全一致**，区别只在部署形态：一个 exe、一个端口、一份配置。

> **为什么做这个**：两个独立面板各自要填一遍后台 Cookie、总后台 Cookie、保存目录，换电脑或 Cookie 过期时得改两处。
> 综合面板里这些**只填一次**，两个子面板实时共享。

---

## 核心功能

- **一个 exe · 一个端口**：门户 `/`、平台数据 `/city/`、企业投递 `/jd/`，同进程挂载（`DispatcherMiddleware`），互不冲突
- **配置共享**：`config.json` 一份，两个子面板的读写都是「读盘 → 合并 → 写盘」，在任一侧保存即刻对另一侧生效
- **数据分存**：`数据_平台面板/` 与 `数据_简历面板/` 各自独立，互不干扰
- **7 类目 + OCR + 四类全量字段补抓**（`/city/`）：订单管理 / 全职岗位 / 兼职岗位 / 企业库 / 会员企业 / 简历库 / 注册用户
- **企业 → 岗位 → 投递 → 完整简历**（`/jd/`）：企业模糊搜索、岗位列表、投递明细（明文电话/邮箱/QQ）、简历深度解析
- **详情全走后台编辑页**：简历 JlEdit / 岗位 JobEdit / 企业 CompanyAdd（job.admin.ccoo.cn）+ 用户 user_set.asp（www.admin.ccoo.cn），**不需要前台 Cookie，无前台反爬，全并发提速**

## 共享配置项

| 配置项 | 默认 | 用途 |
|---|---|---|
| `cookie` | — | 后台 Cookie，供**两个**子面板的列表接口 + 简历/岗位/企业详情使用 |
| `master_cookie` | — | 总后台 Cookie（www.admin.ccoo.cn），仅「注册用户」全量字段需要，可留空 |
| `save_dir` | — | 导出的默认保存目录 |
| `ocr_scale` | 1 | 电话图片 OCR 放大倍数（实测 1 最准） |
| `fetch_workers` | 8 | 列表 / 后台详情并发线程（后台无 WAF，调大即提速） |
| `fetch_resume_detail` | 1 | `/city/`：抓类目时是否顺带补抓后台全量字段 |
| `resume_detail_limit` | 0 | `/city/`：详情条数上限，0 = 不限 |

> 后台 Cookie 与总后台 Cookie 的分工：后台 Cookie（job.admin.ccoo.cn）管列表 + 简历/岗位/企业详情；
> 总后台 Cookie（www.admin.ccoo.cn）只管用户详情页 user_set.asp（纯文本，无需 OCR）。详见
> [平台数据面板文档](../city-data-panel/)。

## 后台详情补抓机制（v2.0）

| 机制 | 说明 |
|---|---|
| 详情来源 | 简历 JlEdit / 岗位 JobEdit / 企业 CompanyAdd / 用户 user_set.asp，按 ID 打开编辑页全量解析 |
| 全并发 | 后台编辑页无前台 nginx WAF，无令牌桶、无请求间隔，`fetch_workers` 默认 8 |
| 失败自动补抓 | 一轮跑完后对失败详情自动再跑 2 轮；`need_login` 单独统计，Cookie 失效时明确提示 |
| 断点续抓 | 按锚点列（姓名 / 岗位标题 / 企业名称 / 家乡通ID）判断详情是否已有，「只补缺失」跳过已有 |
| 覆盖率反馈 | 结束时报告「详情覆盖 200/200」；覆盖数 = 总行数即拿到全量 |

## 运行

**方式一：exe（推荐）**

到 [Release](https://github.com/chenzhuanxin/hbtc-data-panels/releases) 下载
`CombinedPanel-v2.0.0-windows-x64.exe`，放到单独目录双击运行。
首次运行会在 exe 同目录生成 `config.json`、`数据_平台面板/`、`数据_简历面板/`。

**方式二：源码**

```bash
pip install flask requests beautifulsoup4 lxml rapidocr_onnxruntime openpyxl onnxruntime
cd src
copy config.example.json config.json
#   cookie        = 后台 Cookie（两个子面板共用）
#   master_cookie = 总后台 Cookie（注册用户全量字段，可留空）
#   fetch_workers = 并发线程（默认 8）
python -X utf8 面板.py
# 门户 http://127.0.0.1:8795/  （平台数据 /city/   企业投递 /jd/）
```

## 重新打包

```bash
pip install pyinstaller
cd src
python 打包exe.py      # 产物：dist/城市中国综合数据面板.exe
```

> 打包命令里显式带三条 `--add-data`，把门户页与两个子面板的前端一起塞进包内。
> exe 体积约 110 MB，主要是 RapidOCR 的 onnxruntime 与模型文件。

## 目录结构

```
combined-panel/
├── index.html                       ← 详细说明页
├── README.md
└── src/
    ├── 面板.py                       # 入口：门户 + DispatcherMiddleware 挂载两个子面板
    ├── admin_detail.py               # 后台详情抓取模块（JlEdit / JobEdit / CompanyAdd / user_set.asp）
    ├── panels/
    │   ├── city/panel.py             # 平台数据子面板
    │   ├── city/static/index.html
    │   ├── jd/panel.py               # 企业投递子面板
    │   └── jd/static/index.html
    ├── static/index.html             # 门户首页
    ├── 打包exe.py / 启动面板.bat
    ├── 城市中国综合数据面板.spec
    ├── app.ico / icon_preview.png
    └── config.example.json
```

## 常见问题

- **端口 8795 被占用**：自动顺延到 8796/8797/8798，以控制台打印的地址为准。
- **两个子面板的配置会互相覆盖吗**：不会。保存时写回的是「默认值 + 磁盘文件」合并后的完整对象，各自专属键会保留。
- **注册用户没有全量字段**：总后台 Cookie 没配或已过期，配好 `master_cookie` 后点「只补抓用户详情」即可。
- **简历库没有工作/教育经历**：后台 Cookie 没配或已过期，配好 `cookie` 后点「只补抓简历库详情」即可。
- **Cookie 只能用 `;` 切分**：`ccoo=logintime=…&uid=…` 这类 Cookie 的**值本身含 `&`**，按 `&` 拆会截断、认证失效。面板已按正确规则处理，整串粘贴即可。
