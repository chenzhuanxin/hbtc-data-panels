# 城市中国综合数据面板

> 城市中国 · 湖北通城网 招聘频道（hbtc.ccoo.cn）| **两套工具装进同一个 exe**
> 在线详细说明：https://chenzhuanxin.github.io/hbtc-data-panels/combined-panel/

本地网页面板（`http://127.0.0.1:8795`）：一个门户页 + 两个内置子面板，**一份 `config.json` 共享配置**。

---

## 三个入口

| 入口 | 等价于 | 说明 |
|---|---|---|
| `/` | 门户导航页 | 两个工具的入口卡片，显示各自配置状态 |
| `/city/` | [城市中国平台数据面板](../city-data-panel/) | 7 类目全量抓取 + 电话图片 OCR + 简历全量字段补抓 |
| `/jd/` | [企业投递简历采集面板](../resume-panel/) | 企业 → 岗位 → 投递简历 → 完整简历内容 |

功能与两个独立面板**完全一致**，区别只在部署形态：一个 exe、一个端口、一份配置。

> **为什么做这个**：两个独立面板各自要填一遍后台 Cookie、前台 Cookie、保存目录，换电脑或 Cookie 过期时得改两处。
> 综合面板里这些**只填一次**，两个子面板实时共享。

---

## 核心功能

- **一个 exe · 一个端口**：门户 `/`、平台数据 `/city/`、企业投递 `/jd/`，同进程挂载（`DispatcherMiddleware`），互不冲突
- **配置共享**：`config.json` 一份，两个子面板的读写都是「读盘 → 合并 → 写盘」，在任一侧保存即刻对另一侧生效
- **数据分存**：`数据_平台面板/` 与 `数据_简历面板/` 各自独立，互不干扰
- **7 类目 + OCR + 简历全量字段**（`/city/`）：订单管理 / 全职岗位 / 兼职岗位 / 企业库 / 会员企业 / 简历库 / 注册用户；简历库可补抓完整简历（91 → 116 列）
- **企业 → 岗位 → 投递 → 完整简历**（`/jd/`）：企业模糊搜索、岗位列表、投递明细（明文电话/邮箱/QQ）、简历深度解析
- **内置反爬限速**：前台简历抓取默认 1 请求/秒，命中 WAF 自动熔断

## 共享配置项

| 配置项 | 默认 | 用途 |
|---|---|---|
| `cookie` | — | 后台 Cookie，供**两个**子面板的列表类接口使用 |
| `front_cookie` | — | 前台 Cookie，供**两个**子面板抓完整简历内容 |
| `save_dir` | — | 导出的默认保存目录 |
| `ocr_scale` | 1 | 电话图片 OCR 放大倍数（实测 1 最准） |
| `fetch_workers` | 4 | 抓取线程（1–8） |
| `front_rate` | 1.0 | 前台简历请求最小间隔（秒） |
| `fetch_resume_detail` | 1 | `/city/`：抓简历库时是否顺带补抓完整简历 |
| `resume_detail_limit` | 0 | `/city/`：详情条数上限，0 = 不限 |

> 后台 Cookie 与前台 Cookie 的区别，以及「完整简历内容」为什么必须走前台页，见
> [平台数据面板文档](../city-data-panel/#简历库全量简历内容怎么拿v120-重点)。

## ⚠️ 反爬限速（重要）

前台简历详情页 `/post/qiuzhi/{id}x.html` 前方是 nginx WAF。实测：**8 线程连抓 200 份时，大约第 120 份开始返回带滑动验证的 `403`**
（文案「疑似异常访问请求，已被系统拦截 / IP频繁访问」），之后**本机 IP 会被临时拉黑**，连正常请求也一起 403。

对比手工脚本（单线程 + 每份 `time.sleep(1)`）长期可用，可知瓶颈是**速率而非并发数**。因此：

- **全局令牌桶**：任意两次前台请求之间强制间隔 `front_rate` 秒，默认 `1.0`（≈1 请求/秒）
- **在途并发上限 2**：信号量限制同时打开的连接数
- **命中即熔断**：连续 3 次被 WAF 拦截就停止剩余任务并提示
- **断点续抓**：已抓到的详情回写缓存，被拦截后隔几分钟点「只补抓简历库详情」即可接着抓
- **拦截识别**：响应体出现「网络安全验证 / 疑似异常访问请求 / IP频繁访问」即判定拦截

| 规模 | 按 1 请求/秒估算 |
|---|---|
| 40 份简历详情 | 约 1 分钟 |
| 500 份简历详情 | 约 8 分钟 |
| 15,733 份（简历库全量） | 约 4–5 小时 |

> 列表类接口（7 个类目、企业搜索、岗位、投递明细）走后台 `job.admin.ccoo.cn`，**不受此限速影响**，仍是并发抓取。

## 运行

**方式一：exe（推荐）**

到 [Release](https://github.com/chenzhuanxin/hbtc-data-panels/releases) 下载
`CombinedPanel-v1.2.0-windows-x64.exe`，放到单独目录双击运行。
首次运行会在 exe 同目录生成 `config.json`、`数据_平台面板/`、`数据_简历面板/`。

**方式二：源码**

```bash
pip install flask requests beautifulsoup4 lxml rapidocr_onnxruntime openpyxl onnxruntime
cd src
copy config.example.json config.json
#   cookie / front_cookie = 后台 / 前台 Cookie（两个子面板共用）
#   fetch_workers         = 抓取线程（1-8，默认 4）
#   front_rate            = 前台简历请求最小间隔（秒），默认 1.0
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
    ├── panels/
    │   ├── city/panel.py             # 平台数据子面板
    │   ├── city/static/index.html
    │   ├── jd/panel.py               # 企业投递子面板
    │   └── jd/static/index.html
    ├── resume_front.py               # 前台简历抓取模块（含反爬限速）
    ├── static/index.html             # 门户首页
    ├── 打包exe.py / 启动面板.bat
    ├── 城市中国综合数据面板.spec
    ├── app.ico / icon_preview.png
    └── config.example.json
```

## 常见问题

- **端口 8795 被占用**：自动顺延到 8796/8797/8798，以控制台打印的地址为准。
- **两个子面板的配置会互相覆盖吗**：不会。保存时写回的是「默认值 + 磁盘文件」合并后的完整对象，各自专属键会保留。
- **提示「被站点安全策略拦截」**：前台 IP 被临时拉黑，等几分钟、把「前台请求间隔」调大后续抓即可，已抓到的部分不丢。
- **简历库没有工作/教育经历**：前台 Cookie 没配或已过期。
- **前台 Cookie 只能用 `;` 切分**：`ccoo=logintime=…&uid=…` 这类 Cookie 的**值本身含 `&`**，按 `&` 拆会截断、认证失效。面板已按正确规则处理，整串粘贴即可。
