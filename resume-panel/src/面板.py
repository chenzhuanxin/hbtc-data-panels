#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
城市中国·湖北通城网 —— 企业岗位投递简历 采集面板
================================================================
流程：
  ① 输入企业名称（模糊匹配）  → ② 得到企业用户名(sid...)
  ③ 拉取该企业发布的职位列表 → ④ 生成职位网页链接
  ⑤ 拉取每个职位收到的投递简历（自动翻页）
  ⑥ 逐份抓取简历内容（后台简历详情页解析）+ 生成简历链接
  ⑦ 全部导出到一个 Excel（含「投递简历」「岗位列表」两个工作表）

运行：python 面板.py      （自动打开  http://127.0.0.1:8790）
"""

import os
import re
import io
import sys
import json
import time
import base64
import hashlib
import threading
import traceback
import webbrowser
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

import resume_front as RF

sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)

# ----------------------------------------------------------------------
# 路径 / 常量
# ----------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _is_frozen():
    return bool(getattr(sys, 'frozen', False))


def _res_dir():
    """打包内置资源目录（static 等只读资源）。"""
    if _is_frozen():
        return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(sys.executable)))
    return BASE_DIR


def _data_dir():
    """可写数据目录（config.json / data_cache / 导出）。exe 模式取 exe 所在目录。"""
    if _is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return BASE_DIR


RES_DIR = _res_dir()
DATA_DIR = _data_dir()
CACHE_DIR = os.path.join(DATA_DIR, 'data_cache')
CONFIG_PATH = os.path.join(DATA_DIR, 'config.json')
STATE_PATH = os.path.join(CACHE_DIR, 'state.json')
OCR_CACHE_PATH = os.path.join(CACHE_DIR, 'ocr_cache.json')
os.makedirs(CACHE_DIR, exist_ok=True)

PORT = 8790
BASE_URL = 'http://job.admin.ccoo.cn'          # 后台
FRONT_URL = 'http://www.hbtc.ccoo.cn'          # 前台网站（生成对外链接用）
LOGIN_URL = BASE_URL + '/Home/login'
VALIDATE_CODE_URL = 'https://vcodeapi.bccoo.cn/home/getvalidcode'

TZ = timezone(timedelta(hours=8))
DT_MIN_MS = -62135596800000                    # .NET DateTime.MinValue

API_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'X-Requested-With': 'XMLHttpRequest',
}

# 职位列表接口参数（key = 企业用户名）
JOB_PARAMS_YZ = {'key': '', 'kindId': '', 'IsCmd': -1, 'fromtime': '2-9000',
                 'fromtype': -1, 'IsExit': -1, 'istjusername': '', 'paixu': 0,
                 'selCcoochk': -1, 'classid': -1, 'hasSend': 0, 'ccoochk': 1,
                 'openSource': 0, 'username': '', 'zphjoblistType': 1, 'zphid': 0}
JOB_PARAMS_JZ = {'key': '', 'kindId': -1, 'IsCmd': -1, 'IsExit': -1,
                 'istjusername': '', 'paixu': 0, 'selCcoochk': -1, 'ccoochk': 1}

SEND_JL_PARAMS = {'key': '', 'istj': -1, 'jobid': '', 'infotype': 1}


def job_link(jobid):
    return f'{FRONT_URL}/post/zhaopin/{jobid}x.html'


def resume_link(jlid):
    return f'{FRONT_URL}/post/qiuzhi/{jlid}x.html'


# ----------------------------------------------------------------------
# 导出列定义
#
# 简历内容改为抓取【前台简历详情页】www.hbtc.ccoo.cn/post/qiuzhi/{jlid}x.html
# （后台 /Post/JlEdit 字段不全，只能拿部分内容）。
# 因此列集合与 resume_front.DETAIL_COLS 对齐，保证详情可直接按列名合并。
# ----------------------------------------------------------------------
WORK_COLS = RF.WORK_COLS     # 工作经历 1-10
EDU_COLS = RF.EDU_COLS       # 教育经历 1-5
LANG_COLS = RF.LANG_COLS     # 语言能力 1-5
CERT_COLS = RF.CERT_COLS     # 获得证书 1-5

BASE_COLS = ['序号', '企业名称', '企业用户名', '岗位ID', '岗位名称', '岗位链接',
             '简历ID', '简历链接', '投递时间',
             '姓名', '性别', '年龄', '学历', '工作经验', '现居地',
             '求职状态', '期望职位', '职位类别', '期望月薪',
             '电话', '邮箱', 'QQ',
             '自我描述标签', '自我描述详情', '头像', '照片']
COLS = (BASE_COLS
        + [f'工作经历{i}' for i in range(1, WORK_COLS + 1)]
        + [f'教育经历{i}' for i in range(1, EDU_COLS + 1)]
        + [f'语言能力{i}' for i in range(1, LANG_COLS + 1)]
        + [f'获得证书{i}' for i in range(1, CERT_COLS + 1)])

# 次要字段：保留但默认隐藏（在 Excel 中可手动取消隐藏）
HIDDEN_COLS = {'头像', '照片',
               '工作经历7', '工作经历8', '工作经历9', '工作经历10',
               '教育经历3', '教育经历4', '教育经历5',
               '语言能力2', '语言能力3', '语言能力4', '语言能力5',
               '获得证书3', '获得证书4', '获得证书5'}

JOB_SHEET_COLS = ['序号', '岗位类型', '岗位ID', '岗位名称', '岗位链接', '投递简历数', '发布时间']
JOB_SHEET_HIDDEN = set()


# ----------------------------------------------------------------------
# 配置
# ----------------------------------------------------------------------
CONFIG_LOCK = threading.RLock()
DEFAULT_CONFIG = {'username': '', 'password': '', 'cookie': '',
                  'front_cookie': '',
                  'save_dir': '', 'ocr_scale': 1, 'fetch_workers': 4,
                  'front_rate': 1.0}               # 前台简历请求最小间隔（秒），1.0=1请求/秒


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            cfg.update(json.load(f))
    except Exception:
        pass
    return cfg


def save_config(cfg):
    with CONFIG_LOCK:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)


CONFIG = load_config()

# ----------------------------------------------------------------------
# 会话（每线程独立）
# ----------------------------------------------------------------------
_TLS = threading.local()


def _cookie_dict(ck):
    d = {}
    for item in (ck or '').split(';'):
        item = item.strip()
        if '=' in item:
            k, v = item.split('=', 1)
            d[k] = v
    return d


def get_session(cookie_str=None):
    s = getattr(_TLS, 'session', None)
    if s is None:
        s = requests.Session()
        s.trust_env = False          # 绕过系统代理，直连
        s.headers.update(API_HEADERS)
        _TLS.session = s
    ck = cookie_str if cookie_str is not None else CONFIG.get('cookie', '')
    if ck:
        s.cookies.update(_cookie_dict(ck))
    return s


def api_get(path, params, timeout=30, referer=None, cookie_str=None):
    s = get_session(cookie_str)
    h = {}
    if referer:
        h['Referer'] = referer
    r = s.get(BASE_URL + path, params=params, headers=h or None, timeout=timeout)
    return r


def api_json(path, params, timeout=30, referer=None):
    r = api_get(path, params, timeout=timeout, referer=referer)
    return r.json()


def test_cookie(cookie_str=None):
    try:
        j = api_json('/Order/AllIncomeReportDataNew',
                     {'page': 1, 'limit': 1}, timeout=20)
        return j.get('code') == 0, f"连接成功，订单总数 {j.get('count', '?')}"
    except ValueError:
        return False, 'Cookie 无效（后台返回登录页而非数据），请重新登录或粘贴新 Cookie'
    except Exception as e:
        return False, f'验证失败: {e}'


# ----------------------------------------------------------------------
# 文本 / 时间 / 非法字符
# ----------------------------------------------------------------------
RE_DOTNET_DATE = re.compile(r'^/Date\((-?\d+)\)/$')
FW_DIGITS = str.maketrans('０１２３４５６７８９', '0123456789')

ILLEGAL_XML_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
EXCEL_MAX_LEN = 32767
CLEAN_STAT = {'cells': 0, 'chars': 0}


def clean_text(s):
    if not s:
        return s
    if ILLEGAL_XML_RE.search(s):
        n = len(ILLEGAL_XML_RE.findall(s))
        s = ILLEGAL_XML_RE.sub('', s)
        CLEAN_STAT['cells'] += 1
        CLEAN_STAT['chars'] += n
    if len(s) > EXCEL_MAX_LEN:
        s = s[:EXCEL_MAX_LEN]
        CLEAN_STAT['cells'] += 1
    return s


def clean_cell(v):
    if v is None:
        return ''
    if isinstance(v, bool):
        return 1 if v else 0
    if isinstance(v, (int, float)):
        return v
    return clean_text(str(v))


def ms_to_str(ms):
    if ms is None:
        return ''
    try:
        ms = int(ms)
    except (TypeError, ValueError):
        return ''
    if ms <= DT_MIN_MS + 1000 or ms <= 0 or ms > 4102444800000:
        return ''
    return datetime.fromtimestamp(ms / 1000, TZ).strftime('%Y-%m-%d %H:%M:%S')


def to_datetime_str(v):
    """兼容 /Date(...)/ 与毫秒时间戳"""
    if v is None:
        return ''
    s = str(v).strip()
    if not s:
        return ''
    m = RE_DOTNET_DATE.match(s)
    if m:
        return ms_to_str(int(m.group(1)))
    if re.fullmatch(r'\d{12,14}', s):
        return ms_to_str(int(s))
    return clean_text(s)


# ----------------------------------------------------------------------
# OCR（base64 电话图片）
# ----------------------------------------------------------------------
_OCR_ENGINE = None
_OCR_LOCK = threading.Lock()
OCR_CACHE = {}
_OCR_CACHE_LOCK = threading.Lock()

try:
    with open(OCR_CACHE_PATH, 'r', encoding='utf-8') as f:
        OCR_CACHE = json.load(f)
except Exception:
    OCR_CACHE = {}


def _persist_ocr_cache():
    try:
        with open(OCR_CACHE_PATH, 'w', encoding='utf-8') as f:
            json.dump(OCR_CACHE, f, ensure_ascii=False)
    except Exception:
        pass


def get_ocr_engine():
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        with _OCR_LOCK:
            if _OCR_ENGINE is None:
                print('[OCR] 正在加载 RapidOCR 引擎…')
                from rapidocr_onnxruntime import RapidOCR
                _OCR_ENGINE = RapidOCR()
                print('[OCR] RapidOCR 就绪')
    return _OCR_ENGINE


def ocr_base64_image(b64_str, scale=1):
    with _OCR_CACHE_LOCK:
        key = hashlib.md5(f'{scale}|{b64_str}'.encode('utf-8')).hexdigest()
        if key in OCR_CACHE:
            return OCR_CACHE[key]
    text = ''
    try:
        raw = base64.b64decode(re.sub(r'^data:image/\w+;base64,', '', b64_str.strip()))
        from PIL import Image
        import numpy as np
        img = Image.open(io.BytesIO(raw))
        if img.mode != 'RGB':
            img = img.convert('RGB')
        if scale and scale > 1:
            img = img.resize((img.width * scale, img.height * scale),
                             Image.Resampling.LANCZOS)
        res, _ = get_ocr_engine()(np.array(img))
        if res:
            text = ''.join(x[1] for x in res).translate(FW_DIGITS).replace(' ', '')
        del raw
    except Exception as e:
        print('[OCR] 识别失败:', e)
    with _OCR_CACHE_LOCK:
        OCR_CACHE[key] = text
    return text


def norm_tel(v, ocr_fn=None):
    """电话字段：base64 图片走 OCR；明文去掉多余空格横线"""
    if v is None:
        return ''
    s = str(v)
    if s.startswith('data:image'):
        return clean_text(ocr_fn(s)) if ocr_fn else ''
    s = s.strip()
    return clean_text(s)


# ----------------------------------------------------------------------
# 业务：企业 → 职位 → 投递简历 → 简历内容
# ----------------------------------------------------------------------
def search_company(keyword, limit=60):
    """企业名称模糊匹配（keyType=1）"""
    j = api_json('/Enterprise/GetCompanyListData',
                 {'page': 1, 'limit': limit, 'keyword': keyword,
                  'keyType': 1, 'openSource': 0},
                 referer=BASE_URL + '/Enterprise/CompanyList')
    out = []
    for x in (j.get('data') or []):
        out.append({
            'id': x.get('id'),
            'CompName': x.get('CompName') or '',
            'username': x.get('username') or '',
            'AreaStr': x.get('AreaStr') or '',
            'ComptradeStr': x.get('ComptradeStr') or '',
            'CompSizeStr': x.get('CompSizeStr') or '',
            'jobCount': x.get('jobCount'),
            'linkman': x.get('oLinkMan') or '',
            'linktel': norm_tel(x.get('oLinkTel')),
            'SysTime': to_datetime_str(x.get('SysTime')),
            'VipType': x.get('VipType'),
        })
    return j.get('count', len(out)), out


def company_jobs(username, include_parttime=True):
    """该企业发布的职位（全职 + 可选兼职）"""
    out = []
    # 全职
    p = dict(JOB_PARAMS_YZ)
    p['key'] = username
    p['page'] = 1
    p['limit'] = 100
    j = api_json('/Post/GetYZJobList', p, referer=BASE_URL + '/Post/JobListPage')
    for x in (j.get('data') or []):
        out.append({
            'id': x.get('id'),
            'title': x.get('title') or '',
            'type': '全职',
            'countjl': x.get('countjl') or 0,
            'compname': x.get('compname') or '',
            'username': x.get('username') or '',
            'link': job_link(x.get('id')),
            'systime': to_datetime_str(x.get('systime') or x.get('edittime')),
            'checkstate': x.get('checkstate'),
            'isExit': x.get('IsExit'),
        })
    if include_parttime:
        try:
            p2 = dict(JOB_PARAMS_JZ)
            p2['key'] = username
            p2['page'] = 1
            p2['limit'] = 100
            j2 = api_json('/Post/GetJZ_YZJobList', p2,
                          referer=BASE_URL + '/Post/JZ_JobListPage')
            for x in (j2.get('data') or []):
                out.append({
                    'id': x.get('id'),
                    'title': x.get('title') or '',
                    'type': '兼职',
                    'countjl': x.get('countjl') or 0,
                    'compname': x.get('compname') or '',
                    'username': x.get('username') or '',
                    'link': job_link(x.get('id')),
                    'systime': to_datetime_str(x.get('systime') or x.get('edittime')),
                    'checkstate': x.get('checkstate'),
                    'isExit': x.get('IsExit'),
                })
        except Exception as e:
            print('[兼职职位] 获取失败（忽略）:', e)
    return out


def fetch_send_jl(jobid, max_pages=0):
    """某职位收到的投递简历（自动翻页）"""
    rows = []
    page = 1
    total = None
    while True:
        p = dict(SEND_JL_PARAMS)
        p['jobid'] = jobid
        p['page'] = page
        p['limit'] = 100
        j = api_json('/Post/GetSendJlList', p,
                     referer=BASE_URL + '/Post/JobJlList')
        if total is None:
            total = int(j.get('count') or 0)
        data = j.get('data') or []
        if not data:
            break
        rows.extend(data)
        if len(rows) >= total or len(data) < 100:
            break
        page += 1
        if max_pages and page > max_pages:
            break
        time.sleep(0.15)
    return rows


# ---------------------- 简历详情（后台编辑页解析） ----------------------
def _inp(soup_or_tag, name):
    """按 name 取表单控件的值（优先 input/textarea，其次 select）"""
    for tag in ('input', 'textarea', 'select'):
        el = soup_or_tag.find(tag, attrs={'name': name})
        if el is None:
            continue
        if tag == 'textarea':
            return (el.get_text() or '').strip()
        if tag == 'select':
            return _sel(el)
        return (el.get('value') or '').strip()
    return ''


def _sel(tag):
    """取 <select> 中被选中项的文本"""
    if tag is None:
        return ''
    for o in tag.find_all('option'):
        if o.has_attr('selected'):
            v = (o.get('value') or '').strip()
            if v == '':
                return ''
            t = o.get_text(strip=True)
            return t or v
    return ''


    return ''


def _checked_title(tag, name):
    vals = []
    for el in tag.find_all(attrs={'name': name}):
        if el.has_attr('checked'):
            vals.append(el.get('title') or el.get('value') or '')
    return vals


def _period(sy, sm, ey, em):
    a = f'{sy}.{sm}' if sm else (sy or '')
    if ey == '至今' or ey == '0':
        b = '至今'
    else:
        b = f'{ey}.{em}' if em else (ey or '')
    if not a and not b:
        return ''
    return f'{a}-{b}' if (a and b) else (a or b)


def _fmt_work(li):
    com = _inp(li, 'js-comName')
    pos = _inp(li, 'js-workName')
    sy = _sel(li.select_one('select.js-stime-y'))
    sm = _sel(li.select_one('select.js-stime-m'))
    ey = _sel(li.select_one('select.js-etime-y'))
    em = _sel(li.select_one('select.js-etime-m'))
    ta = li.select_one('textarea.js-work-info')
    content = ta.get_text(strip=True) if ta else ''
    if not (com or pos or content):
        return ''
    per = _period(sy, sm, ey, em)
    parts = []
    if com:
        parts.append(f'{com}（{per}）' if per else com)
    if pos:
        parts.append(f'职位名称：{pos}')
    if content:
        parts.append(f'工作内容：{content}')
    return '  '.join(parts)


def _fmt_edu(li):
    sch = _inp(li, 'js-schoolName')
    zy = _inp(li, 'js-zyName')
    sy = _sel(li.select_one('select.js-edu-stime-y'))
    sm = _sel(li.select_one('select.js-edu-stime-m'))
    ey = _sel(li.select_one('select.js-edu-etime-y'))
    em = _sel(li.select_one('select.js-edu-etime-m'))
    if not (sch or zy):
        return ''
    per = _period(sy, sm, ey, em)
    parts = []
    if sch:
        parts.append(f'{sch}（{per}）' if per else sch)
    if zy:
        parts.append(f'专业：{zy}')
    return '  '.join(parts)


RE_HTML_TAG = re.compile(r'<[^>]+>')


def strip_html(s):
    """去掉富文本残留的 HTML 标签（自我描述常带 <br>、<p> 等）"""
    if not s:
        return ''
    s = re.sub(r'(?i)<\s*br\s*/?\s*>', '\n', s)
    s = re.sub(r'(?i)</\s*(p|div|li|tr)\s*>', '\n', s)
    s = RE_HTML_TAG.sub('', s)
    s = (s.replace('&nbsp;', ' ').replace('&amp;', '&')
          .replace('&lt;', '<').replace('&gt;', '>')
          .replace('&quot;', '"').replace('&#39;', "'"))
    s = re.sub(r'[ \t\u3000]+', ' ', s)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()


def parse_resume_html(html, jlid=''):
    """解析 /Post/JlEdit 页面，得到结构化简历内容"""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    d = {
        '简历ID': jlid,
        '简历链接': resume_link(jlid) if jlid else '',
        '姓名': _inp(soup, 'truename'),
        '性别': _inp(soup, 'Sex'),
        '出生年': _inp(soup, 'bornyear'),
        '学历': _inp(soup, 'edu'),
        '专业': _inp(soup, 'edukind'),
        '工作经验': _inp(soup, 'record'),
        '期望职位': _inp(soup, 'title'),
        '期望月薪': _inp(soup, 'pay'),
        '电话': _inp(soup, 'tel'),
        '邮箱': _inp(soup, 'email'),
        '籍贯': _inp(soup, 'jiguan'),
        '居住地': _inp(soup, 'juzhudi'),
        '毕业学校': _inp(soup, 'eduaddr'),
        '头像': _inp(soup, 'headPic'),
        '照片': _inp(soup, 'UpPic'),
        '求职状态': '',
        '自我描标签': '',
        '自我描述详情': '',
    }
    st = _checked_title(soup, 'jlstate')
    d['求职状态'] = st[0] if st else ''
    tags = [t for t in _checked_title(soup, 'liangdian') if t]
    d['自我描标签'] = ', '.join(tags)
    el = soup.find('input', id='text_merecord')
    if el is not None:
        d['自我描述详情'] = strip_html((el.get('value') or '').strip())

    # 工作经历
    works = []
    ul = soup.find('ul', class_='js-work-list')
    if ul is not None:
        for li in ul.find_all('li'):
            if 'item' not in (li.get('class') or []):
                continue
            t = _fmt_work(li)
            if t:
                works.append(t)
    for i in range(WORK_COLS):
        d[f'工作经历{i + 1}'] = works[i] if i < len(works) else ''

    # 教育经历
    edus = []
    ul = soup.find('ul', class_='js-edu-list')
    if ul is not None:
        for li in ul.find_all('li'):
            if 'item' not in (li.get('class') or []):
                continue
            t = _fmt_edu(li)
            if t:
                edus.append(t)
    for i in range(EDU_COLS):
        d[f'教育经历{i + 1}'] = edus[i] if i < len(edus) else ''

    # 语言能力（同名字段按出现顺序配对）
    langs, says, reads = [], [], []
    for el in soup.find_all(attrs={'name': 'LanguageType'}):
        langs.append(_sel(el))
    for el in soup.find_all(attrs={'name': 'SayAbility'}):
        if el.name == 'select':
            says.append(_sel(el))
        else:
            says.append((el.get('value') or '').strip())
    for el in soup.find_all(attrs={'name': 'ReadAbility'}):
        if el.name == 'select':
            reads.append(_sel(el))
        else:
            reads.append((el.get('value') or '').strip())
    lang_items = []
    for i in range(max(len(langs), len(says), len(reads))):
        name = langs[i] if i < len(langs) else ''
        sa = says[i] if i < len(says) else ''
        rd = reads[i] if i < len(reads) else ''
        if not (name or sa or rd):
            continue
        parts = [name] if name else []
        if sa:
            parts.append(f'听说：{sa}')
        if rd:
            parts.append(f'读写：{rd}')
        lang_items.append('  '.join(parts))
    for i in range(LANG_COLS):
        d[f'语言能力{i + 1}'] = lang_items[i] if i < len(lang_items) else ''

    # 获得证书
    cers, zs = [], []
    for el in soup.find_all(attrs={'name': 'CertificateTitle'}):
        cers.append((el.get('value') or '').strip())
    for el in soup.find_all(attrs={'name': 'zstime'}):
        zs.append((el.get('value') or '').strip())
    cer_items = []
    for i in range(max(len(cers), len(zs))):
        n = cers[i] if i < len(cers) else ''
        t = zs[i] if i < len(zs) else ''
        if not n:          # 证书名称为空则不收录（页面默认带了个日期占位）
            continue
        cer_items.append(f'{n}（{t}）' if t else n)
    for i in range(CERT_COLS):
        d[f'获得证书{i + 1}'] = cer_items[i] if i < len(cer_items) else ''

    # 年龄
    by = d.get('出生年') or ''
    m = re.search(r'(19|20)\d{2}', by)
    if m:
        age = datetime.now(TZ).year - int(m.group(0))
        if 10 <= age <= 90:
            d['年龄'] = f'{age}岁'
    d.setdefault('年龄', '')

    return d


def fetch_resume_detail(jlid):
    """
    抓取并解析单份简历。

    主路径：**前台简历详情页** www.hbtc.ccoo.cn/post/qiuzhi/{jlid}x.html
            —— 字段最全（含工作经历1-10 / 教育经历 / 语言能力 / 证书 / 职位类别 / 现居地），
               必须带前台登录 Cookie。
    兜底：后台 /Post/JlEdit（字段较少，仅在前台 Cookie 未配置或失效时使用）
    """
    fc = CONFIG.get('front_cookie') or ''
    if RF.normalize_cookie(fc):
        d = RF.fetch_one(jlid, fc)
        if d and not d.get('_error'):
            return d
    # ---- 兜底：后台编辑页 ----
    try:
        r = api_get('/Post/JlEdit', {'source': 1, 'jlid': jlid},
                    timeout=30, referer=BASE_URL + '/Post/JobJlList')
        r.encoding = 'utf-8'
        if r.status_code != 200 or len(r.text) < 2000:
            return {'简历ID': jlid, '简历链接': resume_link(jlid)}
        d = parse_resume_html(r.text, jlid)
        # 旧列名映射到新列名
        if d.get('自我描标签') and not d.get('自我描述标签'):
            d['自我描述标签'] = d['自我描标签']
        if d.get('居住地') and not d.get('现居地'):
            d['现居地'] = d['居住地']
        return d
    except Exception as e:
        print(f'[简历详情] {jlid} 失败: {e}')
        return {'简历ID': jlid, '简历链接': resume_link(jlid)}


# ----------------------------------------------------------------------
# 全局状态
# ----------------------------------------------------------------------
STATE_LOCK = threading.RLock()
STATE = {'company': None, 'jobs': [], 'apps': [], 'built_at': 0}
PROG = {'phase': 'idle', 'msg': '', 'job_done': 0, 'job_total': 0,
        'apps': 0, 'detail_done': 0, 'detail_total': 0, 'rows': 0}
PROG_LOCK = threading.Lock()
FETCHING = {'on': False}


def _prog(**kw):
    with PROG_LOCK:
        PROG.update(kw)


def save_state():
    with STATE_LOCK:
        tmp = STATE_PATH + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(STATE, f, ensure_ascii=False)
        os.replace(tmp, STATE_PATH)


def load_state():
    global STATE
    if os.path.exists(STATE_PATH):
        try:
            with open(STATE_PATH, 'r', encoding='utf-8') as f:
                STATE = json.load(f)
            print(f"[缓存] 企业={((STATE.get('company') or {}).get('CompName'))!r} "
                  f"职位={len(STATE.get('jobs') or [])} 投递={len(STATE.get('apps') or [])}")
        except Exception as e:
            print('[缓存] 读取失败:', e)


# ----------------------------------------------------------------------
# 抓取流水线
# ----------------------------------------------------------------------
def build_rows(apps, jobs_by_id, company):
    """把投递记录 + 简历详情拼成导出行"""
    rows = []
    for i, a in enumerate(apps, 1):
        jobid = a.get('jobid')
        jb = jobs_by_id.get(str(jobid)) or {}
        det = a.get('detail') or {}
        row = {c: '' for c in COLS}
        row['序号'] = i
        row['企业名称'] = company.get('CompName') or jb.get('compname') or ''
        row['企业用户名'] = company.get('username') or ''
        row['岗位ID'] = jobid or ''
        row['岗位名称'] = a.get('jobtitle') or jb.get('title') or ''
        row['岗位链接'] = job_link(jobid) if jobid else ''
        row['简历ID'] = a.get('jlid') or ''
        row['简历链接'] = a.get('jllink') or (resume_link(a.get('jlid')) if a.get('jlid') else '')
        row['投递时间'] = a.get('addtime_str') or ''
        # 简历详情
        for k in COLS:
            if k in det and det.get(k) not in (None, ''):
                row[k] = det[k]
        # 简历详情缺失时用投递接口的字段兜底
        if not row['姓名']:
            row['姓名'] = a.get('truename') or ''
        if not row['电话']:
            row['电话'] = a.get('tel_str') or ''
        if not row['期望职位'] and a.get('title'):
            row['期望职位'] = a['title']
        if not row['现居地'] and a.get('telarea'):
            row['现居地'] = a['telarea']
        if not row['邮箱']:
            row['邮箱'] = a.get('email') or ''
        if not row['QQ']:
            row['QQ'] = a.get('qq') or ''
        rows.append(row)
    return rows


def run_fetch(company, jobs, fetch_detail=True, use_ocr=True, max_pages=0):
    """抓取指定职位的投递简历 + 简历详情"""
    try:
        FETCHING['on'] = True
        scale = int(CONFIG.get('ocr_scale') or 1)
        workers = max(1, min(16, int(CONFIG.get('fetch_workers') or 8)))
        ocr_fn = (lambda s: ocr_base64_image(s, scale)) if use_ocr else None
        job_total = len(jobs)
        _prog(phase='running', msg=f'开始抓取 {job_total} 个职位的投递简历…',
              job_done=0, job_total=job_total, apps=0,
              detail_done=0, detail_total=0, rows=0)

        apps = []
        lock = threading.Lock()
        done = [0]

        def one(job):
            jid = job.get('id')
            try:
                raw = fetch_send_jl(jid, max_pages=max_pages)
            except Exception as e:
                print(f'[投递] 职位 {jid} 失败: {e}')
                raw = []
            got = []
            for x in raw:
                got.append({
                    'jobid': jid,
                    'jobtitle': job.get('title') or '',
                    'jobtype': job.get('type') or '',
                    'id': x.get('id'),
                    'jlid': x.get('jlid'),
                    'jllink': resume_link(x.get('jlid')) if x.get('jlid') else '',
                    'truename': x.get('truename') or '',
                    'username': x.get('username') or '',
                    'title': x.get('title') or '',
                    'tel_str': norm_tel(x.get('tel'), ocr_fn),
                    'email': x.get('email') or '',
                    'qq': x.get('qq') or '',
                    'telarea': x.get('telarea') or '',
                    'addtime_str': to_datetime_str(x.get('addtime')),
                    'detail': None,
                })
            with lock:
                apps.extend(got)
                done[0] += 1
                _prog(job_done=done[0], apps=len(apps),
                      msg=f'已抓取 {done[0]}/{job_total} 个职位，共 {len(apps)} 条投递记录')
            return got

        with ThreadPoolExecutor(max_workers=workers) as ex:
            list(ex.map(one, jobs))

        # 简历详情（前台简历详情页，字段最全）
        if fetch_detail:
            jlids = []
            seen = set()
            for a in apps:
                jl = a.get('jlid')
                if jl and str(jl) not in seen:
                    seen.add(str(jl))
                    jlids.append(jl)
            _prog(detail_total=len(jlids), msg=f'正在抓取 {len(jlids)} 份简历内容…')
            det_map = {}
            need_login = 0
            waf_block = 0
            fck = CONFIG.get('front_cookie') or ''

            if RF.normalize_cookie(fck) and jlids:
                f_workers = max(1, min(8, int(CONFIG.get('fetch_workers') or 4)))
                try:
                    f_rate = float(CONFIG.get('front_rate') or 1.0)
                except (TypeError, ValueError):
                    f_rate = 1.0

                def onp(done, total, ok, fail):
                    if done % 5 == 0 or done == total:
                        _prog(detail_done=done,
                              msg=f'简历详情 {done}/{total}（成功 {ok} / 失败 {fail}）…')
                got, stat = RF.fetch_many(jlids, fck, workers=f_workers, on_progress=onp,
                                          min_interval=f_rate, max_inflight=2,
                                          waf_stop_after=3, retry_rounds=2,
                                          retry_delay=8.0)
                det_map = {str(k): v for k, v in got.items()}
                need_login = int(stat.get('need_login') or 0)
                waf_block = int(stat.get('waf_block') or 0)
                print(f'[简历详情] 前台抓取 成功 {stat["ok"]} 失败 {stat["fail"]} '
                      f'需登录 {need_login} 被拦截 {waf_block}')

            # 未取到的（未配置前台 Cookie / Cookie 失效 / 个别失败）用后台页兜底
            miss = [j for j in jlids if str(j) not in det_map]
            if miss:
                _prog(msg=f'有 {len(miss)} 份未取到前台详情，改用后台兜底抓取…')
                dn = [0]

                def one_detail(jl):
                    d = fetch_resume_detail(jl)
                    with lock:
                        det_map[str(jl)] = d
                        dn[0] += 1
                        if dn[0] % 5 == 0 or dn[0] == len(miss):
                            _prog(detail_done=len(det_map),
                                  msg=f'兜底抓取 {dn[0]}/{len(miss)} …')

                with ThreadPoolExecutor(max_workers=workers) as ex:
                    list(ex.map(one_detail, miss))

            for a in apps:
                a['detail'] = det_map.get(str(a.get('jlid')))

            if need_login and need_login == len(jlids):
                _prog(msg=f'⚠ {need_login} 份简历全部返回登录页：前台 Cookie 已失效，'
                          f'请在「连接设置」中更新前台 Cookie 后重抓')
            elif waf_block:
                _prog(msg=f'⚠ 有 {waf_block} 份被站点安全策略拦截，已自动熔断停止；'
                          f'其余部分已改用后台页兜底。IP 解封需要时间，请等十几分钟后再重抓齐全。')

        _persist_ocr_cache()

        jobs_by_id = {str(j.get('id')): j for j in jobs}
        rows = build_rows(apps, jobs_by_id, company)
        with STATE_LOCK:
            STATE['company'] = company
            STATE['jobs'] = jobs
            STATE['apps'] = apps
            STATE['built_at'] = time.time()
        save_state()
        _prog(phase='done', rows=len(rows),
              msg=f'完成：{len(jobs)} 个职位，{len(apps)} 条投递记录，{len(rows)} 行数据')
        print(f'[抓取] 完成：{len(apps)} 条投递 / {len(rows)} 行')
    except Exception as e:
        traceback.print_exc()
        _prog(phase='error', msg=f'{e}')
    finally:
        FETCHING['on'] = False


def start_fetch(company, jobs, fetch_detail=True, use_ocr=True, max_pages=0):
    with PROG_LOCK:
        if FETCHING['on']:
            return False, '已有抓取任务进行中'
    threading.Thread(target=run_fetch,
                     args=(company, jobs, fetch_detail, use_ocr, max_pages),
                     daemon=True).start()
    return True, '已开始'


# ----------------------------------------------------------------------
# Excel 导出
# ----------------------------------------------------------------------
EXPORT = {'phase': 'idle', 'msg': '', 'file': '', 'written': 0, 'total': 0}
EXPORT_LOCK = threading.Lock()

HEADER_FILL = '4F7CFF'
HEADER_FONT = 'FFFFFF'


def _style_header(ws, cols, hidden):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    thin = Side(style='thin', color='D9E1F2')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=1, column=i, value=clean_cell(c))
        cell.font = Font(bold=True, color=HEADER_FONT, size=11)
        cell.fill = PatternFill('solid', fgColor=HEADER_FILL)
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=True)
        cell.border = border
        letter = get_column_letter(i)
        w = max(len(str(c)) * 2 + 4, 10)
        if c in ('岗位链接', '简历链接'):
            w = 46
        elif c in ('自我描述详情', '工作内容'):
            w = 40
        elif c.startswith(('工作经历', '教育经历')):
            w = 38
        elif c in ('自我描标签', '期望职位'):
            w = 24
        cd = ws.column_dimensions[letter]
        cd.width = min(w, 60)
        if c in hidden:
            cd.hidden = True
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = 'A2'


def export_excel(rows, jobs, company, path, sheets=('apps', 'jobs')):
    """导出 Excel。sheets 指定要写的工作表：'apps'=投递简历，'jobs'=岗位列表。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    sheets = [s for s in ('apps', 'jobs') if s in set(sheets or ())]
    if not sheets:
        sheets = ['apps']
    wb = Workbook()
    wb.remove(wb.active)
    n_apps = n_jobs = 0

    # --- 工作表：投递简历 ---
    if 'apps' in sheets:
        ws = wb.create_sheet('投递简历')
        _style_header(ws, COLS, HIDDEN_COLS)
        for ri, row in enumerate(rows, 2):
            for ci, c in enumerate(COLS, 1):
                v = clean_cell(row.get(c, ''))
                cell = ws.cell(row=ri, column=ci, value=v)
                if c in ('简历链接', '岗位链接'):
                    cell.hyperlink = v
                    cell.font = Font(color='0563C1', underline='single')
                elif c in ('自我描述详情', '工作经历1', '工作经历2', '工作经历3',
                           '工作经历4', '工作经历5', '工作经历6'):
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            n_apps = ri - 1
            if ri % 200 == 0:
                with EXPORT_LOCK:
                    EXPORT['written'] = ri - 1
        ws.auto_filter.ref = f'A1:{get_column_letter(len(COLS))}1'

    # --- 工作表：岗位列表 ---
    if 'jobs' in sheets:
        ws2 = wb.create_sheet('岗位列表')
        _style_header(ws2, JOB_SHEET_COLS, JOB_SHEET_HIDDEN)
        for ri, j in enumerate(jobs, 2):
            vals = [ri - 1, j.get('type') or '', j.get('id') or '', j.get('title') or '',
                    j.get('link') or '', j.get('countjl') or 0, j.get('systime') or '']
            for ci, v in enumerate(vals, 1):
                cell = ws2.cell(row=ri, column=ci, value=clean_cell(v))
                if ci == 5:
                    cell.hyperlink = v
                    cell.font = Font(color='0563C1', underline='single')
            n_jobs = ri - 1
            if ri % 200 == 0:
                with EXPORT_LOCK:
                    EXPORT['written'] = ri - 1
        ws2.auto_filter.ref = f'A1:{get_column_letter(len(JOB_SHEET_COLS))}1'

    if not wb.sheetnames:
        wb.create_sheet('空')
    wb.save(path)
    return {'apps': n_apps, 'jobs': n_jobs}


def auto_filename(company, kind='投递简历'):
    name = (company.get('CompName') or '企业').strip()
    name = re.sub(r'[\\/:*?"<>|]', '_', name)[:24]
    stamp = datetime.now().strftime('%Y%m%d_%H%M')
    return f'{name}_{kind}_{stamp}'


def do_export(save_dir, only_jobids=None, sheets=None):
    global EXPORT
    try:
        with EXPORT_LOCK:
            EXPORT = {'phase': 'running', 'msg': '准备导出…', 'file': '',
                      'written': 0, 'total': 0, 'files': []}
        CLEAN_STAT['cells'] = 0
        CLEAN_STAT['chars'] = 0
        if not save_dir or not os.path.isdir(save_dir):
            raise RuntimeError(f'保存目录无效: {save_dir}')
        want = [s for s in ('apps', 'jobs') if s in set(sheets or ('apps', 'jobs'))]
        if not want:
            want = ['apps', 'jobs']
        with STATE_LOCK:
            company = STATE.get('company') or {}
            jobs = list(STATE.get('jobs') or [])
            apps = list(STATE.get('apps') or [])

        if only_jobids:
            keep = {str(x) for x in only_jobids}
            jobs = [j for j in jobs if str(j.get('id')) in keep]
            apps = [a for a in apps if str(a.get('jobid')) in keep]

        if 'apps' in want and not apps:
            raise RuntimeError('还没有投递简历数据，请先抓取，或取消勾选「投递简历」')
        if 'jobs' in want and not jobs:
            raise RuntimeError('还没有岗位数据，请先获取岗位列表')

        jobs_by_id = {str(j.get('id')): j for j in jobs}
        rows = build_rows(apps, jobs_by_id, company) if 'apps' in want else []
        total = (len(rows) if 'apps' in want else 0) + \
                (len(jobs) if 'jobs' in want else 0)
        with EXPORT_LOCK:
            EXPORT['total'] = total
            EXPORT['msg'] = f'正在写入 {total} 行…'

        if want == ['jobs']:
            kind = '岗位列表'
        elif want == ['apps']:
            kind = '投递简历'
        else:
            kind = '投递简历+岗位列表'
        path = os.path.join(save_dir, auto_filename(company, kind) + '.xlsx')
        res = export_excel(rows, jobs, company, path, want)
        parts = []
        if 'apps' in want:
            parts.append(f'投递简历 {res["apps"]} 行')
        if 'jobs' in want:
            parts.append(f'岗位列表 {res["jobs"]} 个')
        msg = '导出完成：' + '，'.join(parts)
        if CLEAN_STAT['cells']:
            msg += f"｜已清洗 {CLEAN_STAT['cells']} 处非法字符"
        print(f'[导出] {msg} -> {path}')
        with EXPORT_LOCK:
            EXPORT.update(phase='done', msg=msg, file=path, files=[path],
                          written=total)
    except Exception as e:
        traceback.print_exc()
        with EXPORT_LOCK:
            EXPORT = {'phase': 'error', 'msg': str(e), 'file': '',
                      'written': 0, 'total': 0, 'files': []}


def start_export(save_dir, only_jobids=None, sheets=None):
    with EXPORT_LOCK:
        if EXPORT.get('phase') == 'running':
            return False, '已有导出任务进行中'
    threading.Thread(target=do_export, args=(save_dir, only_jobids, sheets),
                     daemon=True).start()
    return True, '已开始'


# ----------------------------------------------------------------------
# Flask
# ----------------------------------------------------------------------
from flask import Flask, request, jsonify, send_from_directory, Response

app = Flask(__name__, static_folder=os.path.join(RES_DIR, 'static'),
            static_url_path='/static')
_LOGIN_SESSIONS = {}
_PICK_LOCK = threading.Lock()


@app.after_request
def add_cors(resp):
    resp.headers['Access-Control-Allow-Origin'] = '*'
    resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    return resp


@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')


@app.route('/api/config', methods=['GET'])
def api_config_get():
    return jsonify(dict(CONFIG))


@app.route('/api/config', methods=['POST'])
def api_config_set():
    global CONFIG
    d = request.get_json(force=True, silent=True) or {}
    with CONFIG_LOCK:
        for k in ('username', 'password', 'cookie', 'front_cookie', 'save_dir'):
            if k in d:
                CONFIG[k] = str(d[k] or '')
        if 'fetch_workers' in d:
            try:
                CONFIG['fetch_workers'] = max(1, min(8, int(d['fetch_workers'])))
            except (TypeError, ValueError):
                pass
        if 'front_rate' in d:
            try:
                CONFIG['front_rate'] = min(5.0, max(0.0, float(d['front_rate'])))
            except (TypeError, ValueError):
                pass
        if 'ocr_scale' in d:
            try:
                CONFIG['ocr_scale'] = max(1, min(20, int(d['ocr_scale'])))
            except (TypeError, ValueError):
                pass
        save_config(CONFIG)
        CONFIG = load_config()
    return jsonify({'ok': True})


@app.route('/api/config/export', methods=['GET'])
def api_config_export():
    data = json.dumps(CONFIG, ensure_ascii=False, indent=2)
    return Response(data, mimetype='application/json',
                    headers={'Content-Disposition':
                             'attachment; filename=qiye_toudi_config.json'})


@app.route('/api/config/import', methods=['POST'])
def api_config_import():
    global CONFIG
    d = request.get_json(force=True, silent=True)
    if not isinstance(d, dict):
        return jsonify({'ok': False, 'msg': '不是有效的 JSON 配置文件'})
    with CONFIG_LOCK:
        for k in ('username', 'password', 'cookie', 'front_cookie', 'save_dir'):
            if k in d:
                CONFIG[k] = str(d[k] or '')
        if 'fetch_workers' in d:
            try:
                CONFIG['fetch_workers'] = max(1, min(8, int(d['fetch_workers'])))
            except (TypeError, ValueError):
                pass
        if 'front_rate' in d:
            try:
                CONFIG['front_rate'] = min(5.0, max(0.0, float(d['front_rate'])))
            except (TypeError, ValueError):
                pass
        if 'ocr_scale' in d:
            try:
                CONFIG['ocr_scale'] = max(1, min(20, int(d['ocr_scale'])))
            except (TypeError, ValueError):
                pass
        save_config(CONFIG)
        CONFIG = load_config()
    return jsonify({'ok': True, 'msg': '导入成功'})


@app.route('/api/captcha', methods=['POST'])
def api_captcha():
    try:
        s = requests.Session()
        s.trust_env = False
        s.headers.update(API_HEADERS)
        s.get(BASE_URL + '/', timeout=15)
        r = s.get(VALIDATE_CODE_URL, timeout=15)
        from PIL import Image
        img = Image.open(io.BytesIO(r.content))
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        b64 = base64.b64encode(buf.getvalue()).decode()
        _LOGIN_SESSIONS[threading.get_ident()] = s
        return jsonify({'ok': True, 'image': 'data:image/png;base64,' + b64})
    except Exception as e:
        return jsonify({'ok': False, 'msg': str(e)})


@app.route('/api/login', methods=['POST'])
def api_login():
    global CONFIG
    d = request.get_json(force=True, silent=True) or {}
    username = str(d.get('username') or '')
    password = str(d.get('password') or '')
    code = str(d.get('code') or '')
    s = _LOGIN_SESSIONS.get(threading.get_ident())
    if s is None:
        return jsonify({'ok': False, 'msg': '请先获取验证码'})
    try:
        r = s.post(LOGIN_URL, data={'username': username, 'password': password,
                                    'code': code}, timeout=15)
        j = r.json()
        if j.get('code', -1) >= 0:
            ck = '; '.join(f'{c.name}={c.value}' for c in s.cookies)
            with CONFIG_LOCK:
                CONFIG['username'] = username
                CONFIG['password'] = password
                CONFIG['cookie'] = ck
                save_config(CONFIG)
                CONFIG = load_config()
            _LOGIN_SESSIONS.pop(threading.get_ident(), None)
            ok, msg = test_cookie(ck)
            return jsonify({'ok': ok, 'msg': '登录成功，Cookie 已保存' if ok else msg,
                            'cookie': ck})
        return jsonify({'ok': False, 'msg': f'登录失败: {j.get("msg", j)}'})
    except Exception as e:
        return jsonify({'ok': False, 'msg': f'登录异常: {e}'})


@app.route('/api/test_cookie', methods=['POST'])
def api_test_cookie():
    d = request.get_json(force=True, silent=True) or {}
    ck = d.get('cookie')
    if ck is None:
        ck = CONFIG.get('cookie', '')
    ok, msg = test_cookie(str(ck))
    return jsonify({'ok': ok, 'msg': msg})


@app.route('/api/front_test', methods=['POST'])
def api_front_test():
    """测试「前台 Cookie」是否可用于抓取简历详情页"""
    global CONFIG
    d = request.get_json(force=True, silent=True) or {}
    ck = d.get('front_cookie')
    if ck is None:
        ck = CONFIG.get('front_cookie', '')
    ck = str(ck)
    ok, msg, sample = RF.test_front_cookie(ck)
    if ok and ck:
        with CONFIG_LOCK:
            CONFIG['front_cookie'] = ck
            save_config(CONFIG)
            CONFIG = load_config()
    return jsonify({'ok': ok, 'msg': msg, 'sample': sample})


@app.route('/api/search_company', methods=['POST'])
def api_search_company():
    d = request.get_json(force=True, silent=True) or {}
    kw = str(d.get('keyword') or '').strip()
    if not kw:
        return jsonify({'ok': False, 'msg': '请输入企业名称'})
    ok, msg = test_cookie()
    if not ok:
        return jsonify({'ok': False, 'msg': 'Cookie 无效，请先在「连接设置」中配置'})
    try:
        cnt, lst = search_company(kw)
        return jsonify({'ok': True, 'count': cnt, 'list': lst})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'ok': False, 'msg': f'搜索失败: {e}'})


@app.route('/api/company_jobs', methods=['POST'])
def api_company_jobs():
    d = request.get_json(force=True, silent=True) or {}
    username = str(d.get('username') or '').strip()
    if not username:
        return jsonify({'ok': False, 'msg': '缺少企业用户名'})
    try:
        jobs = company_jobs(username, include_parttime=bool(d.get('include_parttime', True)))
        with STATE_LOCK:
            STATE['company'] = d.get('company') or {'username': username}
            STATE['jobs'] = jobs
            STATE['apps'] = []
        save_state()
        return jsonify({'ok': True, 'jobs': jobs})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'ok': False, 'msg': f'获取职位失败: {e}'})


@app.route('/api/fetch', methods=['POST'])
def api_fetch():
    d = request.get_json(force=True, silent=True) or {}
    jobs = d.get('jobs') or []
    company = d.get('company') or {}
    if not jobs:
        return jsonify({'ok': False, 'msg': '请至少选择一个职位'})
    ok, msg = test_cookie()
    if not ok:
        return jsonify({'ok': False, 'msg': 'Cookie 无效，请重新配置'})
    ok, msg = start_fetch(company, jobs,
                          fetch_detail=bool(d.get('fetch_detail', True)),
                          use_ocr=bool(d.get('use_ocr', True)),
                          max_pages=int(d.get('max_pages') or 0))
    return jsonify({'ok': ok, 'msg': msg})


@app.route('/api/progress', methods=['GET'])
def api_progress():
    with PROG_LOCK:
        return jsonify(dict(PROG))


@app.route('/api/state', methods=['GET'])
def api_state():
    with STATE_LOCK:
        company = STATE.get('company')
        jobs = STATE.get('jobs') or []
        apps = STATE.get('apps') or []
        built = STATE.get('built_at') or 0
    jobs_by_id = {str(j.get('id')): j for j in jobs}
    rows = build_rows(apps, jobs_by_id, company or {}) if apps else []
    return jsonify({
        'company': company,
        'jobs': jobs,
        'apps': len(apps),
        'rows': len(rows),
        'built_at': (datetime.fromtimestamp(built, TZ).strftime('%m-%d %H:%M:%S')
                     if built else ''),
    })


@app.route('/api/data', methods=['GET'])
def api_data():
    page = max(1, int(request.args.get('page', 1)))
    size = max(1, min(500, int(request.args.get('size', 50))))
    kw = (request.args.get('kw') or '').strip()
    with STATE_LOCK:
        company = STATE.get('company') or {}
        jobs = list(STATE.get('jobs') or [])
        apps = list(STATE.get('apps') or [])
    jobs_by_id = {str(j.get('id')): j for j in jobs}
    rows = build_rows(apps, jobs_by_id, company) if apps else []
    if kw:
        rows = [r for r in rows
                if any(kw in str(v) for v in r.values())]
    total = len(rows)
    start = (page - 1) * size
    return jsonify({'ok': True, 'total': total, 'cols': COLS,
                    'hidden': sorted(HIDDEN_COLS),
                    'rows': rows[start:start + size]})


@app.route('/api/export', methods=['POST'])
def api_export():
    d = request.get_json(force=True, silent=True) or {}
    ok, msg = start_export(d.get('dir') or CONFIG.get('save_dir', ''),
                           d.get('jobids'), d.get('sheets'))
    return jsonify({'ok': ok, 'msg': msg})


@app.route('/api/export_progress', methods=['GET'])
def api_export_progress():
    with EXPORT_LOCK:
        return jsonify(dict(EXPORT))


@app.route('/api/pick_dir', methods=['POST'])
def api_pick_dir():
    """弹出系统目录选择框"""
    d = request.get_json(force=True, silent=True) or {}
    init = d.get('init') or CONFIG.get('save_dir') or os.path.expanduser('~')
    if not _PICK_LOCK.acquire(blocking=False):
        return jsonify({'ok': False, 'msg': '目录选择框已打开'})
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        p = filedialog.askdirectory(title='选择 Excel 保存目录',
                                    initialdir=init if os.path.isdir(init) else None)
        root.destroy()
        if p:
            p = os.path.normpath(p)
            with CONFIG_LOCK:
                CONFIG['save_dir'] = p
                save_config(CONFIG)
            return jsonify({'ok': True, 'dir': p})
        return jsonify({'ok': False, 'msg': '已取消'})
    except Exception as e:
        return jsonify({'ok': False, 'msg': f'打开选择框失败: {e}'})
    finally:
        _PICK_LOCK.release()


@app.route('/api/open_path', methods=['POST'])
def api_open_path():
    d = request.get_json(force=True, silent=True) or {}
    p = d.get('path') or ''
    try:
        if os.path.isdir(p):
            os.startfile(p)  # noqa
        elif os.path.isfile(p):
            os.startfile(os.path.dirname(p))  # noqa
        else:
            return jsonify({'ok': False, 'msg': '路径不存在'})
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'msg': str(e)})


@app.route('/api/clear', methods=['POST'])
def api_clear():
    with STATE_LOCK:
        STATE['company'] = None
        STATE['jobs'] = []
        STATE['apps'] = []
        STATE['built_at'] = 0
    save_state()
    _prog(phase='idle', msg='', job_done=0, job_total=0, apps=0,
          detail_done=0, detail_total=0, rows=0)
    return jsonify({'ok': True})


# ----------------------------------------------------------------------
def main():
    load_state()
    port = PORT
    for p in (PORT, PORT + 1, PORT + 2, PORT + 3):
        try:
            import socket
            sk = socket.socket()
            sk.bind(('127.0.0.1', p))
            sk.close()
            port = p
            break
        except OSError:
            continue
    url = f'http://127.0.0.1:{port}'
    print('=' * 62)
    print('  城市中国 · 企业岗位投递简历 采集面板')
    print(f'  面板地址: {url}')
    print(f'  配置/缓存目录: {DATA_DIR}')
    print('  按 Ctrl+C 退出')
    print('=' * 62)
    try:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    except Exception:
        pass
    app.run(host='127.0.0.1', port=port, debug=False, threaded=True)


if __name__ == '__main__':
    main()
