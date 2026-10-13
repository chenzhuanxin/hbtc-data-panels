# 城市中国平台数据面板

> 城市中国 · 湖北通城网 招聘频道（hbtc.ccoo.cn）后台 **7 类数据一键采集与导出**
> 在线详细说明：https://chenzhuanxin.github.io/hbtc-data-panels/city-data-panel/

本地网页面板（`http://127.0.0.1:8787`），把招聘后台 7 个菜单下的数据一次抓全、清洗好、按规范格式导出 Excel。

---

## 核心功能

- **7 类目全量抓取**：订单管理、全职岗位、兼职岗位、企业库、会员企业、简历库、注册用户
- **电话图片 OCR**：简历库 `tel`、注册用户 `mobile` 是 base64 内嵌图片（前缀写 png 实为 JPEG），用 RapidOCR **离线**识别成文本
- **简历全量字段补抓**（v1.1.0）：列表接口只给摘要，面板再逐份打开前台简历详情页，把**工作经历 / 教育经历 / 语言能力 / 获得证书**等完整内容补进同一张表（简历库 91 列 → **116 列**）
- **勾选式导出**：单类目 = 一个 `.xlsx`，多类目 = 自动打包 `.zip`；文件名自动带类目与时间戳
- **字段完整 + 智能隐藏**：保留接口全部字段（两行表头：英文 / 中文），对照参考 Excel 把次要列设为隐藏列
- **脏数据清洗**：自动剔除 Excel 不支持的 XML 控制字符，超长文本截断到 32767 字符
- **并发抓取**：分页并发（列表默认 6 线程）+ 详情并发（默认 8 线程，可配置），`count` 偏小时自动顺序补拉

## 7 个类目与实测数据量

| 类目 | 后台接口 | 实测行数 |
|---|---|---|
| 订单管理 | `/Order/AllIncomeReportDataNew` | 4,203 |
| 全职岗位 | `/Post/GetYZJobList` | 719 |
| 兼职岗位 | `/Post/GetJZ_YZJobList` | 23 |
| 企业库 | `/Enterprise/GetCompanyListData` | 4,735 |
| 会员企业 | `/Enterprise/GetEnterpriseData` | 198 |
| 简历库 | `/Post/GetPostJlList` | 15,733 |
| 注册用户 | `/Customer/GetUserInfoList` | 73,046 |
| **合计** | | **98,657** |

> 「账户订单管理」与「订单管理」同数据源，后台已停用，面板中标记为不可用并自动跳过。

## 简历库「全量简历内容」怎么拿（v1.1.0 重点）

列表接口 `/Post/GetPostJlList` 返回的是**列表级字段**——姓名性别年龄、学历、期望职位、电话图片等，
**不含简历正文**。完整简历只存在于**前台简历详情页**：

```
http://www.hbtc.ccoo.cn/post/qiuzhi/{简历ID}x.html
```

该页面是服务端渲染的完整简历，但**必须带前台站点登录 Cookie**：不带 Cookie（或只带后台 Cookie）
会返回约 22 KB 的登录页，里面没有任何简历内容——这正是「拿到了列表却没有简历全量数据」的原因。

因此面板把两种 Cookie 分开配置：

| 配置项 | 用途 |
|---|---|
| `cookie` | 后台 Cookie——7 个类目的接口 |
| `front_cookie` | 前台 Cookie——逐份打开简历详情页，取完整简历 |

**前台 Cookie 取法**：浏览器登录 `http://www.hbtc.ccoo.cn` → F12 → Network → 任意请求 →
复制 Request Headers 里的整条 `Cookie` → 粘到面板「前台 Cookie」框 → 点「测试前台 Cookie」。

**补抓后多出的 25 列**：

| 新增字段 | 列数 |
|---|---|
| 工作经历 1 – 10 | 10 |
| 教育经历 1 – 5 | 5 |
| 语言能力 1 – 5 | 5 |
| 获得证书 1 – 5 | 5 |

> 姓名 / 性别 / 年龄 / 学历 / 工作经验 / 现居地 / 求职状态 / 期望职位 / 职位类别 / 期望月薪 /
> 自我描述标签 / 自我描述详情 这 12 个前台字段在列表里本来就有，补抓时只覆盖校准，不新增列。
> 导出默认隐藏 14 列次要列（工作经历 7-10、教育经历 3-5、语言能力 2-5、获得证书 3-5）。

**两个已知坑**：

- **Cookie 只能用 `;` 切分**。`ccoo=logintime=…&uid=…`、`ccoocitycrm=pass&crm_username=…`
  这些 Cookie 的**值本身含 `&`**，按 `&` 拆会把它们截断，认证随即失效、页面变回登录页。
- **本机代理**。若代理把 `127.0.0.1` 也接管了，请求会被劫持成 502。
  面板给每个线程单独建 `requests.Session` 并关闭环境代理（`trust_env = False`）规避。

**实测性能**：默认 8 线程（`fetch_workers`，1–16 可调），单份平均约 1.4 s；40 份 2 秒完成、
30 份在 16 线程下 1.4 秒（约 22 份/秒）；单份失败自动重试 3 次，整批汇总「成功 / 失败 / 需登录」。

> 只想补详情不必重跑列表：点简历库行的「只补抓简历库详情」即可。建议先用「详情条数上限」小批量试跑。

## 几个关键实现

- **OCR 放大倍数默认 1**：经过真值对比（39/40），**原始尺寸识别最准，放大反而更容易失败**——与常见「放大 10 倍」的经验相反。识别结果缓存到 `data_cache/ocr_cache.json`。
- **时间统一格式化**：`/Date(毫秒)/` 与毫秒数值统一转为 `yyyy-mm-dd hh:mm:ss`（东八区）；.NET 最小值 `-62135596800000` 转空值。
- **简历详情页按小节解析**：工作经历 / 教育经历 / 语言能力 / 获得证书四个区块**共用同一个 class**，直接按 class 抓会错位（实测出现过「工作经历6 = 心理健康指导」）。解析时先按 `<p class="j-tit-module">工作经历</p>` 定位区块，再在区块内匹配条目；时间段节点自带括号，需避免二次包裹成 `（（2019.6-2020.8））`；正文 `<br>` 转换行并反转义实体。
- **头像占位图识别**：占位图文件名存在 `offer-list-autor.png` / `offer-list-autor1.png` 两种拼写，统一按 `offer-list-autor` 前缀 + `nophoto` / `noavatar` / `touxiang_moren` 等关键字过滤。
- **导出勾选状态独立存储**：勾选状态放在独立 `Set` 里而不是读 DOM，后台轮询刷新时不会把用户已选类目重置；新抓取完成的类目自动补勾。

## 运行

**方式一：exe（推荐）**

到 [Release](https://github.com/chenzhuanxin/hbtc-data-panels/releases) 下载
`CityDataPanel-v1.1.0-windows-x64.exe`，放到单独目录双击运行。

**方式二：源码**

```bash
pip install flask requests beautifulsoup4 lxml rapidocr_onnxruntime openpyxl onnxruntime
cd src
copy config.example.json config.json
#   cookie              = 后台 Cookie（7 个类目接口）
#   front_cookie        = 前台 Cookie（简历全量内容，必填才有工作/教育/语言/证书）
#   fetch_resume_detail = 1 抓简历详情 / 0 只抓列表
#   resume_detail_limit = 0 不限条数，填 N 则只补抓前 N 份
#   fetch_workers       = 详情并发线程数（1-16，默认 8）
python -X utf8 面板.py
```

## 重新打包

```bash
pip install pyinstaller
cd src
python 打包exe.py      # 产物：dist/城市中国平台数据面板.exe
```

## 目录结构

```
city-data-panel/
├── index.html                    ← 详细说明页
├── README.md
└── src/
    ├── 面板.py                    # 主程序（Flask 服务 + 抓取 + 导出）
    ├── resume_front.py            # 前台简历详情抓取与解析（两个项目共用同一份）
    ├── static/index.html          # 面板前端
    ├── 打包exe.py                 # 一键打包
    ├── 启动面板.bat               # 源码启动
    ├── 城市中国平台数据面板.spec
    ├── app.ico / icon_preview.png
    └── config.example.json
```

## 常见问题

- **端口 8787 被占用**：面板自动顺延 8788/8789…。若页面数据对不上，多半是旧进程占着 8787、新进程跑在 8788，用 `netstat -ano | findstr :8787` 找 PID 结束后重启。
- **提示登录失效**：Cookie 有时效，重新登录或粘贴最新 Cookie 后点「测试」。
- **简历库抓完了没有工作/教育经历**：前台 Cookie 没配或已过期。列表接口本身不含简历正文，配好 `front_cookie` 后点「只补抓简历库详情」即可，不必重抓列表。
- **前台 Cookie 测试返回登录页**：① Cookie 过期 → 重新登录复制；② 整串很长，可能漏了后半段；③ 粘贴被换行截断（面板会自动去首尾空白并按分号重组，仍建议一次性整串粘贴）。
- **OCR 不准**：保持 `ocr_scale = 1`。
