# 城市中国平台数据面板

> 城市中国 · 湖北通城网 招聘频道（hbtc.ccoo.cn）后台 **7 类数据一键采集与导出**
> 在线详细说明：https://chenzhuanxin.github.io/hbtc-data-panels/city-data-panel/

本地网页面板（`http://127.0.0.1:8787`），把招聘后台 7 个菜单下的数据一次抓全、清洗好、按规范格式导出 Excel。

---

## 核心功能

- **7 类目全量抓取**：订单管理、全职岗位、兼职岗位、企业库、会员企业、简历库、注册用户
- **电话图片 OCR**：简历库 `tel`、注册用户 `mobile` 是 base64 内嵌图片（前缀写 png 实为 JPEG），用 RapidOCR **离线**识别成文本
- **详情全量字段补抓（v2.0）**：列表接口只给摘要，面板按 ID 打开**后台编辑页**补全：
  - 简历 → `job.admin.ccoo.cn/Post/JlEdit?source=1&jlid={jlid}`（工作经历 1-10 / 教育经历 1-5 / 语言能力 1-5 / 获得证书 1-5）
  - 岗位（全职+兼职）→ `/Post/JobEdit?Source=1&id={jobid}`（职位类别/学历/人数/经验/描述/公司信息等 25 列）
  - 企业 → `/Enterprise/CompanyAdd?source=1&companyId={companyId}`（联系人/电话/简介/行业/规模/认证/到期时间等 24 列）
  - 用户 → `www.admin.ccoo.cn/yunying/users/user_set.asp?id={userId}`（用户名/昵称/注册IP/账号状态/论坛信息/会员服务等 33 列，**纯文本无需 OCR**）
- **勾选式导出**：单类目 = 一个 `.xlsx`，多类目 = 自动打包 `.zip`；文件名自动带类目与时间戳
- **字段完整 + 智能隐藏**：保留接口全部字段（两行表头：英文 / 中文），对照参考 Excel 把次要列设为隐藏列
- **脏数据清洗**：自动剔除 Excel 不支持的 XML 控制字符，超长文本截断到 32767 字符
- **全并发抓取**：列表分页并发 + 后台详情全并发（默认 8 线程），后台编辑页无前台 WAF，不再需要限速 / 请求间隔

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

## 三种 Cookie 的分工（v2.0）

| 配置项 | 用途 |
|---|---|
| `username` / `password` | 后台账号密码（job.admin.ccoo.cn），可用「验证码登录」自动获取 Cookie |
| `cookie` | **后台 Cookie**——7 个类目的列表接口 + 简历/岗位/企业详情编辑页 |
| `master_cookie` | **总后台 Cookie**（www.admin.ccoo.cn）——仅「注册用户」全量字段需要；可留空（此时用户只有列表字段） |

**已不需要前台 Cookie**（旧版必须登录前台 www.hbtc.ccoo.cn 才能抓简历正文）。
**总后台 Cookie 取法**：浏览器登录 `http://www.admin.ccoo.cn` → F12 → Network → 任意请求 →
复制 Request Headers 里的整条 `Cookie` → 粘到面板「总后台 Cookie」框 → 点「测试总后台 Cookie」。

**用户全量字段（33 列，来自 user_set.asp）**：

| 分组 | 字段 |
|---|---|
| 注册信息 | 用户名、昵称、注册时间、注册IP、最后登录时间、最后登录IP |
| 个人信息 | 姓名、性别、出生日期（年龄）、手机、QQ、邮箱、头像、个人简介、详细地址 |
| 账号状态 | 账号状态、用户组、经验值、积分、论坛头衔 |
| 会员服务 | 房屋中介 / 名企招聘 / 论坛登录 / 车行 / 二手商家 / 新楼盘管理 / 分类信息发布 / 微信登录 等状态 |
| 用户链接 | 论坛主页 / 后台编辑页链接（便于人工复核） |

**两个已知坑**：

- **Cookie 只能用 `;` 切分**。`ccoo=logintime=…&uid=…`、`ccoocitycrm=pass&crm_username=…`
  这些 Cookie 的**值本身含 `&`**，按 `&` 拆会把它们截断，认证随即失效、页面变回登录页。
- **本机代理**。若代理把 `127.0.0.1` 也接管了，请求会被劫持成 502。
  面板给每个线程单独建 `requests.Session` 并关闭环境代理（`trust_env = False`）规避。

## 后台详情补抓机制

| 机制 | 说明 |
|---|---|
| `fetch_workers` | 列表 / 后台详情并发线程，默认 `8`。后台无 WAF，调大即提速 |
| 失败自动补抓 | 一轮跑完后，对失败的详情自动再跑 2 轮，直到抓全 |
| `need_login` 统计 | Cookie 失效导致的失败单独统计并明确提示（而不是混在普通失败里） |
| 断点续抓 | 按锚点列判断是否已有详情（简历库=`姓名`、岗位=`岗位标题`、企业=`企业名称`、用户=`家乡通ID`），「只补缺失」跳过已有、只抓空缺 |
| 覆盖率反馈 | 补抓结束后面板直接报告「详情覆盖 200/200」这类数字 |

**怎么确认「拿到了全量」**：面板报告覆盖率，如
「简历详情 成功 200 / 失败 0，**详情覆盖 200/200**」。覆盖数 = 总行数即说明每一份都取到了后台详情；
若提示 `need_login`，说明 Cookie 失效，重新粘贴后点「只补缺失详情」续抓即可。

## 几个关键实现

- **OCR 放大倍数默认 1**：经过真值对比（39/40），**原始尺寸识别最准**，放大反而更容易失败。识别结果缓存到 `data_cache/ocr_cache.json`。
- **时间统一格式化**：`/Date(毫秒)/` 与毫秒数值统一转为 `yyyy-mm-dd hh:mm:ss`（东八区）；.NET 最小值 `-62135596800000` 转空值。
- **用户页按表格解析**：user_set.asp 是 `td.align=right` 标签 → 下一 `td` 值的表格结构；注册IP 值形如 `117.154.28.195[查看IP所在地]`，自动去掉 `[...]` 后缀；会员年龄 `31岁（生日：1995/1/1）` 拆出年龄+生日。
- **导出勾选状态独立存储**：勾选状态放在独立 `Set` 里而不是读 DOM，后台轮询刷新时不会把用户已选类目重置；新抓取完成的类目自动补勾。

## 运行

**方式一：exe（推荐）**

到 [Release](https://github.com/chenzhuanxin/hbtc-data-panels/releases) 下载
`CityDataPanel-v2.0.0-windows-x64.exe`，放到单独目录双击运行。

**方式二：源码**

```bash
pip install flask requests beautifulsoup4 lxml rapidocr_onnxruntime openpyxl onnxruntime
cd src
copy config.example.json config.json
#   cookie              = 后台 Cookie（7 个类目接口 + 简历/岗位/企业详情）
#   master_cookie       = 总后台 Cookie（注册用户全量字段，可留空）
#   fetch_resume_detail = 1 抓详情 / 0 只抓列表
#   resume_detail_limit = 0 不限条数，填 N 则只补抓前 N 份
#   fetch_workers       = 并发线程数（默认 8，后台无 WAF，可放心调大）
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
    ├── admin_detail.py            # 后台详情抓取模块（JlEdit / JobEdit / CompanyAdd / user_set.asp，三个项目共用同一份）
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
- **简历库抓完了没有工作/教育经历**：后台 Cookie 没配或已过期。详情来自后台编辑页 JlEdit，配好 `cookie` 后点「只补抓简历库详情」即可，不必重抓列表。
- **注册用户没有全量字段**：总后台 Cookie 没配或已过期。配好 `master_cookie` 后点「只补抓用户详情」即可。
- **测试总后台 Cookie 报错**：① Cookie 过期 → 重新登录复制；② 整串很长，可能漏了后半段；③ 粘贴被换行截断（面板会自动去首尾空白并按分号重组，仍建议一次性整串粘贴）。
- **OCR 不准**：保持 `ocr_scale = 1`。
