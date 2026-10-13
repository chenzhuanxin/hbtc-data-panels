# 企业投递简历采集面板

> 从一家企业出发，把「发布了哪些岗位 · 谁投了简历 · 简历里写了什么」一次拿全
> 在线详细说明：https://chenzhuanxin.github.io/hbtc-data-panels/resume-panel/

本地网页面板（`http://127.0.0.1:8790`），按企业名称链式采集并导出单表多页 Excel。

**采集链路：** 企业名称 → 企业账号 → 岗位列表 → 投递简历 → 简历详情 → Excel

---

## 核心功能

- **企业模糊搜索**：输入「宝顺」「物流有限公司」等关键词即可命中，无需完整全称
- **岗位列表可单独查看/导出**：先确认目标企业发布了哪些岗位、每个岗位收到多少投递
- **投递明细**：投递人姓名、投递时间、电话 / 邮箱 / QQ（电话为**明文**，无需 OCR）
- **简历深度解析**：基本信息、期望职位与薪资、自我描述、工作经历、教育经历、语言能力、证书、头像照片
- **导出可勾选**：仅岗位列表 / 仅投递简历 / 两者都要，文件名自动区分
- **并发抓取**：岗位投递与简历详情均走线程池（默认 6 线程）

## 接口链路（实测可用）

| 环节 | 接口 | 说明 |
|---|---|---|
| ① 企业模糊搜索 | `GET /Enterprise/GetCompanyListData?keyword=关键词&keyType=1&openSource=0` | **keyType=1 才是模糊搜索**；keyType=0 会返回全量企业 |
| ② 岗位列表 | `GET /Post/GetYZJobList?key={企业username}` | 兼职岗位走 `/Post/GetJZ_YZJobList` |
| ③ 岗位投递简历 | `GET /Post/GetSendJlList?jobid={岗位id}&istj=-1&infotype=1` | 分页返回投递人、时间、电话、邮箱、QQ、姓名、简历ID |
| ④ 简历详情 | `GET /Post/JlEdit?source=1&jlid={简历id}` | 服务端渲染**完整简历表单**，后台 Cookie 即可拿到全部字段 |

> 前台简历页 `www.hbtc.ccoo.cn/post/qiuzhi/{id}x.html` 需要**前台登录**，匿名/后台 Cookie 只会返回登录页，
> 因此改走后台 `/Post/JlEdit`。导出时仍生成前台链接方便人工复核。

## 导出格式

**「投递简历」工作表（47 列）**

- 归属信息：序号、企业名称、企业用户名、岗位ID、岗位名称、岗位链接、简历ID、简历链接、投递时间
- 基本信息：姓名、性别、年龄、出生年、学历、专业、工作经验、求职状态
- 求职意向：期望职位、期望月薪
- 联系方式：电话、邮箱、籍贯、居住地、毕业学校
- 简历正文：自我描标签、自我描述详情、头像、照片
- 经历展开：工作经历 1–8、教育经历 1–5、语言能力 1–3、获得证书 1–3

**「岗位列表」工作表（7 列）**：序号、岗位类型、岗位ID、岗位名称、岗位链接、投递简历数、发布时间

其中 14 个次要列默认隐藏（出生年、毕业学校、头像、照片、专业、工作经历 7–8、教育经历 3–5、语言能力 2–3、获得证书 2–3），
在 Excel 中整体取消隐藏即可查看。表内冻结首行、启用筛选、岗位/简历链接为超链接。

## 实测性能

| 测试对象 | 规模 | 结果 |
|---|---|---|
| 湖北宝顺物流有限公司 | 6 个岗位 / 523 条投递 / 502 份简历详情 | 约 **53 秒**完成，导出 524 KB |
| 导出组合 | 仅岗位列表 / 仅投递简历 / 两者都要 | 三档均正常，文件名自动区分 |

## 运行

**方式一：exe（推荐）**

到 [Release](https://github.com/chenzhuanxin/hbtc-data-panels/releases) 下载
`ResumeCollector-v1.0.0-windows-x64.exe`，放到单独目录双击运行。

**方式二：源码**

```bash
pip install flask requests beautifulsoup4 lxml openpyxl
cd src
copy config.example.json config.json     # 必须！否则 Cookie 为空会返回登录页
python -X utf8 面板.py
```

## 重新打包

```bash
pip install pyinstaller
cd src
python 打包exe.py      # 产物：dist/企业投递简历采集面板.exe
```

## 目录结构

```
resume-panel/
├── index.html                    ← 详细说明页
├── README.md
└── src/
    ├── 面板.py                    # 主程序
    ├── static/index.html          # 面板前端
    ├── 打包exe.py                 # 一键打包
    ├── 启动面板.bat               # 源码启动
    ├── 企业投递简历采集面板.spec
    ├── app.ico / icon_preview.png
    └── config.example.json
```

## 常见问题

- **企业搜不到**：换成企业全称里的关键片段（如「宝顺」）。
- **投递数为 0**：该岗位确实没有收到投递，属正常情况。
- **改了代码仍是旧数据**：端口被旧进程占用，新进程静默退出。用 `netstat -ano | findstr :8790` 找 PID 结束后重启。
- **首次启动报 JSONDecodeError**：`config.json` 不存在或 Cookie 为空，复制 `config.example.json` 并填写后重试。
