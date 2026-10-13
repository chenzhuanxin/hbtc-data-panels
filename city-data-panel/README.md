# 城市中国平台数据面板

> 城市中国 · 湖北通城网 招聘频道（hbtc.ccoo.cn）后台 **7 类数据一键采集与导出**
> 在线详细说明：https://chenzhuanxin.github.io/hbtc-data-panels/city-data-panel/

本地网页面板（`http://127.0.0.1:8787`），把招聘后台 7 个菜单下的数据一次抓全、清洗好、按规范格式导出 Excel。

---

## 核心功能

- **7 类目全量抓取**：订单管理、全职岗位、兼职岗位、企业库、会员企业、简历库、注册用户
- **电话图片 OCR**：简历库 `tel`、注册用户 `mobile` 是 base64 内嵌图片（前缀写 png 实为 JPEG），用 RapidOCR **离线**识别成文本
- **勾选式导出**：单类目 = 一个 `.xlsx`，多类目 = 自动打包 `.zip`；文件名自动带类目与时间戳
- **字段完整 + 智能隐藏**：保留接口全部字段（两行表头：英文 / 中文），对照参考 Excel 把次要列设为隐藏列
- **脏数据清洗**：自动剔除 Excel 不支持的 XML 控制字符，超长文本截断到 32767 字符
- **并发抓取**：分页并发（默认 6 线程，可配置），`count` 偏小时自动顺序补拉

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

## 几个关键实现

- **OCR 放大倍数默认 1**：经过真值对比（39/40），**原始尺寸识别最准，放大反而更容易失败**——与常见「放大 10 倍」的经验相反。识别结果缓存到 `data_cache/ocr_cache.json`。
- **时间统一格式化**：`/Date(毫秒)/` 与毫秒数值统一转为 `yyyy-mm-dd hh:mm:ss`（东八区）；.NET 最小值 `-62135596800000` 转空值。
- **导出勾选状态独立存储**：勾选状态放在独立 `Set` 里而不是读 DOM，后台轮询刷新时不会把用户已选类目重置；新抓取完成的类目自动补勾。

## 运行

**方式一：exe（推荐）**

到 [Release](https://github.com/chenzhuanxin/hbtc-data-panels/releases) 下载
`CityDataPanel-v1.0.0-windows-x64.exe`，放到单独目录双击运行。

**方式二：源码**

```bash
pip install flask requests rapidocr_onnxruntime openpyxl onnxruntime
cd src
copy config.example.json config.json     # 填入账号/密码/Cookie
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
    ├── 面板.py                    # 主程序
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
- **OCR 不准**：保持 `ocr_scale = 1`。
