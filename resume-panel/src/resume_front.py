# -*- coding: utf-8 -*-
"""
前台简历详情抓取模块（城市中国·湖北通城网）
================================================
来源方法：用户提供的「获取最新简历信息详情/获取简历信息.py」

核心事实（已实测）：
  * 简历完整内容在【前台】页面  http://www.hbtc.ccoo.cn/post/qiuzhi/{jlid}x.html
  * 该页面**必须带前台登录 Cookie** 才能看到内容（无 Cookie / 只带后台 Cookie 会返回登录页，约 22KB）
  * 后台编辑页 /Post/JlEdit 拿到的字段不全，不能作为主数据源

本模块只做一件事：给一个 jlid，返回该简历在前台页面上的**全量字段**。

可被两个面板直接 import：
    from resume_front import (DETAIL_COLS, resume_url, fetch_one, fetch_many,
                              test_front_cookie, normalize_cookie)
"""
import re
import threading
import time

import requests
from bs4 import BeautifulSoup

FRONT_BASE = 'http://www.hbtc.ccoo.cn'

FRONT_HEADERS = {
    'accept': ('text/html,application/xhtml+xml,application/xml;q=0.9,'
               'image/avif,image/webp,image/apng,*/*;q=0.8'),
    'accept-encoding': 'gzip, deflate',
    'accept-language': 'zh-CN,zh;q=0.9',
    'cache-control': 'max-age=0',
    'connection': 'keep-alive',
    'upgrade-insecure-requests': '1',
    'user-agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36'),
}

# ----------------------------------------------------------------------
# 导出列定义（两个面板共用）
# ----------------------------------------------------------------------
WORK_COLS = 10     # 工作经历 1-10
EDU_COLS = 5       # 教育经历 1-5
LANG_COLS = 5      # 语言能力 1-5
CERT_COLS = 5      # 获得证书 1-5

BASE_DETAIL_COLS = [
    '姓名', '性别', '年龄', '学历', '工作经验', '现居地',
    '求职状态', '期望职位', '职位类别', '期望月薪',
    '自我描述标签', '自我描述详情', '头像', '照片',
]
DETAIL_COLS = (BASE_DETAIL_COLS
               + [f'工作经历{i}' for i in range(1, WORK_COLS + 1)]
               + [f'教育经历{i}' for i in range(1, EDU_COLS + 1)]
               + [f'语言能力{i}' for i in range(1, LANG_COLS + 1)]
               + [f'获得证书{i}' for i in range(1, CERT_COLS + 1)])

# ----------------------------------------------------------------------
# 岗位前台详情字段（参照简历：从 /post/zhaopin/{id}x.html 补抓全量字段）
# 前台岗位详情页相对后台列表接口补充/文本化的完整字段。
# ----------------------------------------------------------------------
JOB_DETAIL_COLS = ['岗位标题', '薪资', '职位类别', '学历要求', '招聘人数',
                   '工作经验', '公司地址', '更新时间', '剩余有效期', '信息关注度',
                   '职位描述', '公司介绍', '公司名称', '公司行业', '公司规模',
                   '公司性质']
JOB_DETAIL_HIDDEN = {'公司介绍', '剩余有效期', '信息关注度', '公司规模'}

# 默认头像占位图（识别后不作为有效头像）
_AVATAR_PLACEHOLDERS = ('offer-list-autor', 'nophoto', 'default', 'noavatar', 'noimg',
                        '/images/blank', 'touxiang_moren')

_LOGIN_PAGE_MARK = 'search_nologin'


def resume_url(jlid):
    return f'{FRONT_BASE}/post/qiuzhi/{jlid}x.html'


def job_url(jobid):
    return f'{FRONT_BASE}/post/zhaopin/{jobid}x.html'


def normalize_cookie(s):
    """
    把 Cookie 字符串或 dict 归一化成 dict。

    ⚠️ 只能按 `;` 分隔！部分 Cookie 的**值本身含 `&`**
    （如 `ccoo=logintime=..&uid=12345678`、后台 `ccoocitycrm=pass&crm_username=..`），
    按 `&` 切会把一条 Cookie 拆成多条，导致服务端判定未登录。
    """
    if not s:
        return {}
    if isinstance(s, dict):
        return {k: str(v) for k, v in s.items() if k}
    out = {}
    for part in re.split(r';', str(s)):
        part = part.strip()
        if not part or '=' not in part:
            continue
        k, v = part.split('=', 1)
        k, v = k.strip(), v.strip()
        if k:
            out[k] = v
    return out


def _safe_text(el):
    return el.get_text(strip=True) if el else ''


def clean_rich(el):
    """取元素文本，保留 <br> 为换行，去掉其余标签与多余空白"""
    if el is None:
        return ''
    s = el.decode_contents() if hasattr(el, 'decode_contents') else str(el)
    s = re.sub(r'(?i)<\s*br\s*/?\s*>', '\n', s)
    s = re.sub(r'(?i)</\s*(p|div|li|tr)\s*>', '\n', s)
    s = re.sub(r'<[^>]+>', '', s)
    s = (s.replace('&nbsp;', ' ').replace('&amp;', '&')
          .replace('&lt;', '<').replace('&gt;', '>')
          .replace('&quot;', '"').replace('&#39;', "'"))
    s = re.sub(r'[ \t\u3000]+', ' ', s)
    s = re.sub(r'\n\s*\n+', '\n', s)
    return s.strip()


def _strip_paren(s):
    s = (s or '').strip()
    while len(s) >= 2 and s[0] in '（(' and s[-1] in '）)':
        s = s[1:-1].strip()
    return s


def _section_by_title(soup, title):
    """按模块标题定位区块（标题 <p class="j-tit-module"> 是区块 div 的直接子元素）"""
    for p in soup.find_all('p', class_='j-tit-module'):
        if title in _safe_text(p):
            par = p.parent
            while par is not None and getattr(par, 'name', None):
                if par.get('class'):
                    return par
                par = par.parent
    return None


def _label_value(li, keyword):
    """在 li 内按 <span class="fl s1">关键词：</span><span>值</span> 取值"""
    for lab in li.find_all('span', class_='fl s1'):
        if keyword in _safe_text(lab):
            v = lab.find_next_sibling('span')
            return clean_rich(v) if v is not None else ''
    return ''


def extract_work(soup):
    """
    提取工作经历。

    ⚠️ 关键：必须**限定在「工作经历」区块内**再匹配 j-work-module，
    因为「语言能力」「获得证书」区块用的是同一个 class，
    直接对整页 HTML 正则匹配会把证书/语言项误当成工作经历。
    """
    section = _section_by_title(soup, '工作经历')
    if section is None:
        for div in soup.find_all('div', class_='j-work'):
            if not (set(div.get('class') or []) & {'j-qz', 'j-language', 'j-proof'}):
                section = div
                break
    if section is None:
        return []

    results = []
    for mod in section.find_all('div', class_='j-work-module'):
        if 'step-circle' not in (mod.get('class') or []):
            continue
        b = mod.find('b')
        company = _safe_text(b)
        t = mod.find('i', class_='time')
        period = _strip_paren(_safe_text(t))
        position = _label_value(mod, '职位名称')
        content = _label_value(mod, '工作内容')
        if not (company or position or content):
            continue
        parts = [company]
        if period:
            parts.append(f'（{period}）')
        if position:
            parts.append(f'职位名称：{position}')
        if content:
            parts.append(f'工作内容：{content}')
        results.append(' '.join([p for p in parts if p]))
    return results


def extract_work_by_regex(html):
    """兼容保留：对整页 HTML 正则提取（会串入证书/语言块，仅作兜底）"""
    pattern = (r'<div class="j-work-module clearfix step-circle">(.*?)</div>\s*'
               r'(?=<div class="j-work-module|</div>\s*</div>|$)')
    blocks = re.findall(pattern, html, re.DOTALL)
    results = []
    for block in blocks:
        m = re.search(r'<b>(.*?)</b>', block)
        company = m.group(1).strip() if m else ''
        m = re.search(r'<i class="time">(.*?)</i>', block)
        period = _strip_paren(m.group(1)) if m else ''
        m = re.search(r'职位名称：</span>\s*<span>(.*?)</span>', block)
        position = m.group(1).strip() if m else ''
        m = re.search(r'工作内容：</span>\s*<span>(.*?)</span>', block, re.DOTALL)
        content = m.group(1).strip() if m else ''
        parts = [company,
                 f'（{period}）' if period else '',
                 f'职位名称：{position}' if position else '',
                 f'工作内容：{content}' if content else '']
        result = ' '.join([p for p in parts if p])
        if result:
            results.append(result)
    return results


def _abs_url(u):
    """//img.xxx/... -> https://img.xxx/..."""
    if not u:
        return ''
    u = u.strip()
    if u.startswith('//'):
        return 'https:' + u
    if u.startswith('/'):
        return FRONT_BASE + u
    return u


def _extract_span_value(soup, label_text):
    """找 <span class="fl s1">标签：</span> 的相邻 <span> 值"""
    label = soup.find('span', class_='fl s1', string=re.compile(label_text))
    if label:
        sib = label.find_next_sibling('span')
        if sib is not None:
            return clean_rich(sib)
    return ''


def _parse_education(edu_ul):
    school = period = ''
    first_li = edu_ul.find('li', class_='j-pos')
    if first_li:
        b = first_li.find('b')
        if b:
            school = _safe_text(b)
        t = first_li.find('i', class_='time')
        if t:
            period = _strip_paren(_safe_text(t))
    major = ''
    for li in edu_ul.find_all('li', class_='j-pos'):
        v = _label_value(li, '所读专业')
        if v:
            major = v
    parts = []
    if school:
        parts.append(school)
    if period:
        parts.append(f'（{period}）')
    if major:
        parts.append(f'专业：{major}')
    return ' '.join(parts)


def _parse_language(div):
    lang = _safe_text(div.find('b'))
    listen = _label_value(div, '听说能力')
    read = _label_value(div, '读写能力')
    parts = []
    if lang:
        parts.append(lang)
    if listen:
        parts.append(f'听说：{listen}')
    if read:
        parts.append(f'读写：{read}')
    return ' '.join(parts)


def _parse_certificate(div):
    name = _safe_text(div.find('b'))
    get_time = _label_value(div, '领证时间')
    parts = []
    if name:
        parts.append(name)
    if get_time:
        parts.append(f'获得时间：{get_time}')
    return ' '.join(parts)


def is_login_page(html):
    """判断是否返回了「需登录」页（无简历内容）"""
    if not html:
        return True
    if is_waf_block(html):
        return False
    if '<div class="j-name">' in html or 'class="j-name"' in html:
        return False
    if 'j-work' in html and 'j-name' in html:
        return False
    return _LOGIN_PAGE_MARK in html or 'j-tx' not in html


# ----------------------------------------------------------------------
# 反爬：WAF 拦截识别 + 全局限速 + 冷却退避
#
# 实测结论：站点前置 nginx WAF 会对前台简历页的高频访问直接返回 403
# 滑动验证页（"疑似异常访问请求…IP频繁访问"）。8 线程连续抓 200 份时，
# 大约第 120 份开始被拦截，之后整段 IP 短时黑名单，全部请求 403。
# 因此必须：① 全局限速（任意两次请求之间强制间隔）② 命中即冷却。
# ----------------------------------------------------------------------
_WAF_MARKS = ('网络安全验证', '疑似异常访问请求', 'IP频繁访问',
              '访问操作提示', '系统拦截')

# 全局限速：任意两次前台请求之间至少间隔这么多秒。
# 默认 2.0 ≈ 0.5 请求/秒 —— 前台 WAF 对高频访问极敏感（IP 频繁访问即 403 拉黑），
# 岗位与简历详情都从前台抓取，因此统一采用保守速率，降低被拦截概率。
MIN_INTERVAL = 2.0
# 同时的在途请求数上限（并发连接数）。1 = 完全串行，最安全。
MAX_INFLIGHT = 1
_RATE_LOCK = threading.Lock()
_inflight = threading.Semaphore(2)
_rate_state = {'next_at': 0.0}
_waf_until = [0.0]
# 命中的 WAF 冷却时长基准（秒），按重试次数递增
WAF_COOLDOWN = 60.0


def is_waf_block(html, status=None):
    """判断响应是否被站点 WAF 拦截（滑动验证页）"""
    if status == 403:
        return True
    if not html:
        return False
    return any(m in html for m in _WAF_MARKS)


def set_rate(min_interval=None, max_inflight=None):
    """调整全局限速。interval 越大越安全；max_inflight 越小越安全。"""
    global MIN_INTERVAL, MAX_INFLIGHT, _inflight
    if min_interval is not None:
        try:
            MIN_INTERVAL = max(0.0, float(min_interval))
        except (TypeError, ValueError):
            pass
    if max_inflight is not None:
        try:
            n = max(1, int(max_inflight))
            if n != MAX_INFLIGHT:
                MAX_INFLIGHT = n
                _inflight = threading.Semaphore(n)
        except (TypeError, ValueError):
            pass
    return MIN_INTERVAL, MAX_INFLIGHT


def _throttle():
    """阻塞到本线程可以发起下一次请求的时刻（全局限速 + WAF 冷却）。

    注意：判断与占位在锁内完成，但**睡眠在锁外**，
    否则多个线程会串行地各睡一遍，批量抓取时越等越久。
    """
    while True:
        with _RATE_LOCK:
            now = time.time()
            wait = max(_rate_state['next_at'], _waf_until[0]) - now
            if wait <= 0:
                _rate_state['next_at'] = now + MIN_INTERVAL
                return
        time.sleep(min(wait, 1.0))


def _trip_waf(cooldown=None):
    """命中 WAF：设置全局冷却截止时间，让所有线程一起停下来"""
    cd = WAF_COOLDOWN if cooldown is None else cooldown
    with _RATE_LOCK:
        _waf_until[0] = max(_waf_until[0], time.time() + cd)


def waf_cooldown_remain():
    """距离 WAF 冷却结束还剩多少秒（0 表示未被冷却）"""
    return max(0.0, _waf_until[0] - time.time())


def parse_front_resume(html, jlid=''):
    """解析前台简历详情页，返回全量字段 dict（失败返回 None）"""
    soup = BeautifulSoup(html, 'html.parser')
    if not soup.find('div', class_='j-name') and not soup.find('div', class_='j-work'):
        return None

    d = {c: '' for c in DETAIL_COLS}
    d['简历ID'] = jlid

    # ---------- 姓名（优先带 h1 的 PC 区块） ----------
    name = ''
    for nd in soup.find_all('div', class_='j-name'):
        h1 = nd.find('h1')
        if h1:
            name = _safe_text(h1)
            break
    if not name:
        for nd in soup.find_all('div', class_='j-name'):
            sp = nd.find('span', class_='jlname')
            if sp:
                name = _safe_text(sp)
                break
    if not name:
        nd = soup.find('div', class_='j-name')
        if nd:
            for i in nd.find_all('i'):
                i.extract()
            name = _safe_text(nd)
    d['姓名'] = name

    # ---------- 头像（默认占位图不算有效头像） ----------
    avatar = ''
    for img in soup.find_all('img', class_='j-tx'):
        src = _abs_url(img.get('src', ''))
        if src and not any(p in src.lower() for p in _AVATAR_PLACEHOLDERS):
            avatar = src
            break
    d['头像'] = avatar

    # ---------- 年龄 / 性别 / 学历 / 工作经验 / 现居地 ----------
    pos_div = soup.find('div', class_='j-pos')
    if pos_div:
        spans = [_safe_text(s) for s in pos_div.find_all('span')]
        d['年龄'] = spans[0] if len(spans) > 0 else ''
        d['性别'] = spans[1] if len(spans) > 1 else ''
        d['学历'] = spans[2] if len(spans) > 2 else ''
        d['工作经验'] = re.sub(r'\s+', ' ', spans[3]).strip() if len(spans) > 3 else ''
        d['现居地'] = spans[4] if len(spans) > 4 else ''

    # ---------- 求职意向 ----------
    d['求职状态'] = _extract_span_value(soup, '求职状态')
    d['期望职位'] = _extract_span_value(soup, '期望职位')
    d['职位类别'] = _extract_span_value(soup, '职位类别')
    d['期望月薪'] = _extract_span_value(soup, '期望月薪')

    # ---------- 自我描述 ----------
    desc_ul = soup.find('ul', class_='j-hot-list clearfix')
    if desc_ul:
        d['自我描述标签'] = ', '.join(_safe_text(li) for li in desc_ul.find_all('li'))
    desc_div = soup.find('div', class_='j-text')
    d['自我描述详情'] = clean_rich(desc_div) if desc_div is not None else ''

    # ---------- 工作经历（限定在工作经历区块内提取） ----------
    works = extract_work(soup)
    for i in range(WORK_COLS):
        d[f'工作经历{i + 1}'] = works[i] if i < len(works) else ''

    # ---------- 教育经历 ----------
    edus = []
    edu_section = _section_by_title(soup, '教育经历') or soup.find('div', class_='j-edu')
    if edu_section:
        for ul in edu_section.find_all('ul', class_='j-edu-list'):
            t = _parse_education(ul)
            if t:
                edus.append(t)
    for i in range(EDU_COLS):
        d[f'教育经历{i + 1}'] = edus[i] if i < len(edus) else ''

    # ---------- 语言能力 ----------
    langs = []
    sec = _section_by_title(soup, '语言能力') or soup.find('div', class_='j-work j-language')
    if sec:
        for div in sec.find_all('div', class_='j-work-module'):
            if 'step-circle' not in (div.get('class') or []):
                continue
            t = _parse_language(div)
            if t:
                langs.append(t)
    for i in range(LANG_COLS):
        d[f'语言能力{i + 1}'] = langs[i] if i < len(langs) else ''

    # ---------- 获得证书 ----------
    certs = []
    sec = _section_by_title(soup, '获得证书') or soup.find('div', class_='j-work j-proof')
    if sec:
        for div in sec.find_all('div', class_='j-work-module'):
            if 'step-circle' not in (div.get('class') or []):
                continue
            t = _parse_certificate(div)
            if t:
                certs.append(t)
    for i in range(CERT_COLS):
        d[f'获得证书{i + 1}'] = certs[i] if i < len(certs) else ''

    # ---------- 照片 ----------
    photo = ''
    pic_ul = soup.find('ul', class_=lambda c: bool(c) and 'material-box' in c
                       and 'evaluation-pic-list' in c)
    if pic_ul:
        li = pic_ul.find('li', class_='li1')
        if li:
            photo = _abs_url(li.get('data-url', '') or li.get('data-src', ''))
    d['照片'] = photo

    return d


# ----------------------------------------------------------------------
# 岗位前台详情解析（/post/zhaopin/{id}x.html）
# ----------------------------------------------------------------------
def _job_basic(ib, keyword):
    """在岗位基本信息 infoBox 内按 <span class=name>标签</span><span class=info>值</span> 取值"""
    for li in ib.find_all('li', class_=lambda c: c and 'item' in c):
        for lab in li.find_all('span', class_='name'):
            if keyword in _safe_text(lab):
                v = lab.find_next_sibling('span')
                return clean_rich(v) if v is not None else ''
    return ''


def parse_front_job(html, jobid=''):
    """解析前台岗位详情页，返回岗位全量字段 dict（页面无效返回 None）"""
    soup = BeautifulSoup(html, 'html.parser')
    if not soup.find(class_='zMain-titBox'):
        return None
    d = {c: '' for c in JOB_DETAIL_COLS}
    d['岗位ID'] = jobid

    # ---------- 标题 ----------
    tt = soup.find(class_='zMain-titBox')
    d['岗位标题'] = _safe_text(tt)

    # ---------- 薪资（主岗位 p.price 或 span.tab.price） ----------
    sal = soup.find('p', class_='price')
    if sal is None:
        sal = soup.find('span', class_=lambda c: c and 'price' in c and 'tab' in c)
    d['薪资'] = _safe_text(sal)

    # ---------- 基本信息（主岗位 infoBox：含 招聘职位/学历要求） ----------
    main_ib = None
    for ib in soup.find_all(class_=lambda c: c and 'infoBox' in c):
        if '招聘职位' in _safe_text(ib) or '学历要求' in _safe_text(ib):
            main_ib = ib
            break
    if main_ib is not None:
        d['职位类别'] = _job_basic(main_ib, '招聘职位')
        d['学历要求'] = _job_basic(main_ib, '学历要求')
        d['招聘人数'] = _job_basic(main_ib, '招聘人数')
        d['工作经验'] = _job_basic(main_ib, '工作经验')
        d['公司地址'] = _job_basic(main_ib, '公司地址')

    # ---------- 更新时间 / 剩余有效期 / 信息关注度 ----------
    top = soup.find('div', class_='zMainBox')
    if top is not None:
        t = _safe_text(top)
        m = re.search(r'更新时间[：:]\s*([0-9]{2,4}[-/年].*?)(?=\s|剩余|信息|举报)', t)
        if m:
            d['更新时间'] = m.group(1).strip()
        m = re.search(r'剩余有效期[：:]\s*([\d\s.]+?)\s*天', t)
        if m:
            d['剩余有效期'] = m.group(1).strip() + '天'
        m = re.search(r'信息关注度[：:]\s*([\d,]+)\s*人次', t)
        if m:
            d['信息关注度'] = m.group(1).strip() + '人次'

    # ---------- 职位描述 / 公司介绍（zCon 区块，按标题定位） ----------
    for z in soup.find_all('div', class_='zCon'):
        t = _safe_text(z)
        if t.startswith('职位描述'):
            d['职位描述'] = t[len('职位描述'):].strip()
        elif t.startswith('公司介绍'):
            d['公司介绍'] = t[len('公司介绍'):].strip()

    # ---------- 公司信息（zComBox：<p class=name> 名称 + ul.items 行业/规模/性质） ----------
    for cb in soup.find_all('div', class_=lambda c: c and 'zComBox' in c):
        name_el = cb.find('p', class_='name')
        if name_el is None:
            continue
        d['公司名称'] = _safe_text(name_el)
        for item in cb.find_all('li', class_='item'):
            span = item.find('span')
            if span is None:
                continue
            st = _safe_text(span)
            for key, col in (('公司行业', '公司行业'), ('公司规模', '公司规模'),
                             ('公司性质', '公司性质')):
                if st.startswith(key):
                    val = st[len(key):].lstrip('：:').strip()
                    if val:
                        d[col] = val
        break

    return d


def is_job_login_page(html):
    """岗位详情页有效判断：无岗位标题区即视为登录页/无内容"""
    if not html:
        return True
    if is_waf_block(html):
        return False
    return 'zMain-titBox' not in html and 'zMainBox' not in html


# ----------------------------------------------------------------------
# 抓取
# ----------------------------------------------------------------------
_TLS = threading.local()


def _session(cookie_str=None):
    s = getattr(_TLS, 'session', None)
    if s is None:
        s = requests.Session()
        s.trust_env = False          # 绕过本机代理，避免 127.0.0.1 被劫持
        _TLS.session = s
    if cookie_str:
        s.cookies = requests.cookies.cookiejar_from_dict(normalize_cookie(cookie_str))
    return s


def fetch_one(jlid, cookie_str, timeout=25, retries=3, session=None, kind='resume'):
    """抓取并解析单个对象。

    kind='resume' 抓简历详情页 /post/qiuzhi/{id}x.html；
    kind='job'    抓岗位详情页 /post/zhaopin/{id}x.html。
    成功返回 dict，失败返回 {'_error': ...}。
    """
    url = job_url(jlid) if kind == 'job' else resume_url(jlid)
    idkey = '岗位ID' if kind == 'job' else '简历ID'
    last = None
    for attempt in range(max(1, retries)):
        try:
            _throttle()                      # 全局限速 + WAF 冷却
            s = session or _session(cookie_str)
            with _inflight:                  # 限制在途并发连接数
                r = s.get(url, headers=FRONT_HEADERS, timeout=timeout)
            r.encoding = 'utf-8'
            if is_waf_block(r.text, r.status_code):
                # 被 WAF 拦截：立刻返回，不做任何等待或重试。
                return {'_error': 'waf_block', idkey: jlid}
            if r.status_code != 200:
                last = f'HTTP {r.status_code}'
                time.sleep(0.5 * (attempt + 1))
                continue
            if kind == 'job':
                if is_job_login_page(r.text):
                    return {'_error': 'need_login', idkey: jlid}
                d = parse_front_job(r.text, jlid)
            else:
                if is_login_page(r.text):
                    return {'_error': 'need_login', idkey: jlid}
                d = parse_front_resume(r.text, jlid)
            if d is None:
                last = '页面无内容'
                time.sleep(0.5 * (attempt + 1))
                continue
            if kind == 'job':
                d['岗位链接'] = url
            else:
                d['简历链接'] = url
            return d
        except Exception as e:
            last = f'{type(e).__name__}: {e}'
            time.sleep(0.5 * (attempt + 1))
    return {'_error': str(last or 'unknown'), idkey: jlid}


def fetch_one_job(jobid, cookie_str, **kw):
    """抓取并解析单个岗位详情页。"""
    return fetch_one(jobid, cookie_str, kind='job', **kw)


def _fetch_round(jlids, cookie_str, workers, timeout, retries, should_stop=None,
                 on_progress=None, offset=0, grand_total=None, ok0=0, fail0=0,
                 stats=None, out=None, lock=None, kind='resume'):
    """跑一轮并发抓取；成功的写入 out，失败的收集后返回，供下一轮重试。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed

    stats = stats if stats is not None else {'need_login': 0, 'waf_block': 0}
    out = out if out is not None else {}
    lock = lock or threading.Lock()
    grand = grand_total or len(jlids)

    def work(jid):
        one = requests.Session()
        one.trust_env = False
        one.cookies = requests.cookies.cookiejar_from_dict(normalize_cookie(cookie_str))
        return jid, fetch_one(jid, cookie_str, timeout=timeout,
                              retries=retries, session=one, kind=kind)

    failed = []
    stopped = [False]
    done = [0]

    with ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
        futs = [ex.submit(work, j) for j in jlids]
        for fu in as_completed(futs):
            if stopped[0] or (should_stop and should_stop()):
                for f in futs:
                    f.cancel()
                stopped[0] = True
                break
            try:
                jid, res = fu.result()
            except Exception as e:
                jid, res = None, {'_error': f'{type(e).__name__}: {e}'}
            with lock:
                done[0] += 1
                err = (res or {}).get('_error')
                if res and not err:
                    out[jid] = res
                else:
                    failed.append(jid)
                    if err == 'need_login':
                        stats['need_login'] += 1
                    elif err == 'waf_block':
                        stats['waf_block'] += 1
                if on_progress:
                    on_progress(offset + done[0], grand, len(out),
                                fail0 + len(failed))
    return out, failed, stats, stopped[0]


def fetch_many(jlids, cookie_str, workers=4, on_progress=None, should_stop=None,
               timeout=25, retries=3, min_interval=None, max_inflight=None,
               waf_stop_after=3, retry_rounds=2, retry_delay=8.0, kind='resume'):
    """
    并发抓取多份简历/岗位（全量）。返回 (out, stats)

      kind='resume' 抓简历详情页；kind='job' 抓岗位详情页。
      out    : {id: detail_dict}    只含成功的
      stats  : {'done','ok','fail','need_login','waf_block','total','stopped',
                'rounds'}

    on_progress(done, total, ok, fail) 每完成一份回调一次
    should_stop() 返回 True 时提前结束

    可靠抓取策略（保证"全量"）：
      1) 全局限速：无论 workers 多大，实际速率被 MIN_INTERVAL 卡住
         （默认 0.5 请求/秒），在途连接由 MAX_INFLIGHT 限制（默认 1，串行）。
      2) 失败自动补抓：第一轮结束后，对失败的 id 再跑 `retry_rounds` 轮，
         轮间等待 `retry_delay` 秒（给站点缓冲）。
      3) WAF 熔断：一旦命中拦截页达到 waf_stop_after 次，立即停止本批，
         避免把 IP 打进更长时间的黑名单；停止后未抓到的 id 由调用方
         在稍后「只补缺失」时续抓。
    """
    if min_interval is not None or max_inflight is not None:
        set_rate(min_interval, max_inflight)

    total = len(jlids)
    if not total:
        return {}, {'done': 0, 'ok': 0, 'fail': 0, 'need_login': 0,
                    'waf_block': 0, 'total': 0, 'stopped': False, 'rounds': 0}

    out, stats_all = {}, {'need_login': 0, 'waf_block': 0}
    lock = threading.Lock()
    pending = list(jlids)
    rounds = 0
    stopped = False
    ok_total = 0

    for rnd in range(max(1, retry_rounds + 1)):
        if not pending:
            break
        if rnd > 0:
            if should_stop and should_stop():
                stopped = True
                break
            # 轮间缓冲：只有在上一轮"有失败"时才等待
            wait = retry_delay * rnd
            if on_progress:
                on_progress(ok_total, total, ok_total, len(pending))
            t_end = time.time() + wait
            while time.time() < t_end:
                if should_stop and should_stop():
                    break
                time.sleep(min(0.5, t_end - time.time()))

        rounds = rnd + 1
        before_waf = stats_all['waf_block']
        offset = ok_total
        _fetch_round(pending, cookie_str, workers, timeout, retries,
                     should_stop=should_stop, on_progress=on_progress,
                     offset=offset, grand_total=total, ok0=before_waf,
                     stats=stats_all, out=out, lock=lock, kind=kind)

        ok_total = len(out)
        # 待补 = 仍未成功的
        pending = [j for j in pending if j not in out]

        # WAF 熔断：本轮出现新的拦截页，立即停止，不再补抓
        if stats_all['waf_block'] > before_waf and \
                stats_all['waf_block'] >= waf_stop_after:
            stopped = True
            break
        if not pending:
            break

    fail_n = total - len(out)
    if on_progress:
        on_progress(total, total, len(out), fail_n)
    return out, {'done': total, 'ok': len(out), 'fail': fail_n,
                 'need_login': stats_all['need_login'],
                 'waf_block': stats_all['waf_block'],
                 'total': total, 'stopped': stopped, 'rounds': rounds}


def fetch_many_job(jobids, cookie_str, **kw):
    """并发抓取多份岗位详情页（全量），复用 fetch_many 的限速/熔断/补抓机制。"""
    return fetch_many(jobids, cookie_str, kind='job', **kw)


def test_front_cookie(cookie_str, probe_jlid='1906236'):
    """用一份已知简历试抓，判断前台 Cookie 是否有效。返回 (ok, msg, sample_name)"""
    if not normalize_cookie(cookie_str):
        return False, '未填写前台 Cookie', ''
    res = fetch_one(probe_jlid, cookie_str, timeout=20, retries=2)
    if not res:
        return False, '请求失败（网络或超时），请重试', ''
    err = res.get('_error')
    if err == 'need_login':
        return False, '前台 Cookie 已失效，请重新登录 www.hbtc.ccoo.cn 后复制最新 Cookie', ''
    if err == 'waf_block':
        return False, ('已被站点安全策略拦截（IP 频繁访问），请等十几分钟后再试、'
                       '并把「抓取线程」调小'), ''
    if err:
        return False, f'抓取失败：{err}', ''
    return True, f'有效（示例简历：{res.get("姓名") or "—"}）', res.get('姓名', '')


if __name__ == '__main__':
    import sys
    ck = sys.argv[1] if len(sys.argv) > 1 else ''
    ids = sys.argv[2:] or ['1906236', '3018176', '2262588']
    print(test_front_cookie(ck))
    for i in ids:
        d = fetch_one(i, ck, retries=2)
        if not d or d.get('_error'):
            print(i, '-> 失败', d)
            continue
        print(f'\n=== {i} {d["姓名"]} {d["性别"]} {d["年龄"]} {d["学历"]} {d["工作经验"]} {d["现居地"]}')
        print(f'    求职状态={d["求职状态"]} 期望职位={d["期望职位"]} '
              f'职位类别={d["职位类别"]} 期望月薪={d["期望月薪"]}')
        print(f'    标签={d["自我描述标签"]!r}  详情={d["自我描述详情"][:80]!r}')
        for k in DETAIL_COLS:
            if k.startswith(('工作经历', '教育经历', '语言能力', '获得证书')) and d[k]:
                print(f'    {k}: {d[k][:110]}')
