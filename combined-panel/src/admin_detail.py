# -*- coding: utf-8 -*-
"""
后台全量字段抓取模块（城市中国·湖北通城网）
================================================
v2.0 重构：详情数据源由「前台站点」切换到「后台/总后台编辑页」。

数据源（均无需前台 Cookie）：
  * 简历全量：后台编辑页  http://job.admin.ccoo.cn/Post/JlEdit?source=1&jlid={jlid}
  * 岗位全量：后台编辑页  http://job.admin.ccoo.cn/Post/JobEdit?Source=1&id={jobid}
  * 企业全量：后台编辑页  http://job.admin.ccoo.cn/Enterprise/CompanyAdd?source=1&companyId={companyId}
  * 用户全量：总后台会员管理页 http://www.admin.ccoo.cn/yunying/users/user_set.asp?id={userId}
              （必须带「总后台 Cookie」，即 www.admin.ccoo.cn 的登录 Cookie；
               页面为纯文本展示，**不需要 OCR**）

与旧版（前台简历/岗位详情页）相比：
  * 不再需要前台站点登录 Cookie（前台 WAF 高频访问会 403 拉黑，已一并移除）
  * 后台页面没有前台 nginx WAF 限速，详情可全并发抓取，速度大幅提升
  * OCR 保持不变：列表接口的 base64 电话图片仍由各面板的 RapidOCR 识别保存
  * 失败自动补抓 / 只补缺失（断点续抓）机制保留

可被两个面板直接 import：
    from admin_detail import (DETAIL_COLS, ADMIN_JOB_COLS, ADMIN_COMPANY_COLS,
                              ADMIN_USER_COLS, normalize_cookie, fetch_one,
                              fetch_many, test_master_cookie, resume_url, job_url)
"""
import re
import threading
import time

import requests
from bs4 import BeautifulSoup

JOB_ADMIN_BASE = 'http://job.admin.ccoo.cn'
MASTER_ADMIN_BASE = 'http://www.admin.ccoo.cn'
FRONT_BASE = 'http://www.hbtc.ccoo.cn'

ADMIN_HEADERS = {
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
# 导出列定义（三个面板共用）
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
# 岗位全量字段（后台 JobEdit 编辑页）
# ----------------------------------------------------------------------
ADMIN_JOB_COLS = ['岗位ID', '岗位标题', '职位类别', '薪资', '学历要求', '招聘人数',
                  '工作经验', '工作区域', '公司名称', '公司地址', '公司行业',
                  '公司规模', '公司性质', '联系人', '联系电话', '邮箱', '福利',
                  '职位描述', '公司简介', '发布时间', '更新时间', '审核状态',
                  '是否下线', '点击量', '岗位链接']
ADMIN_JOB_HIDDEN = {'公司简介', '工作区域', '公司行业', '公司规模', '公司性质',
                    '更新时间', '点击量'}

# ----------------------------------------------------------------------
# 企业全量字段（后台 CompanyAdd 编辑页）
# ----------------------------------------------------------------------
ADMIN_COMPANY_COLS = ['企业ID', '企业名称', '企业用户名', '联系人', '联系电话',
                      '公司简介', '公司行业', '公司类型', '公司规模', '公司地址',
                      '区域', '注册时间', '编辑时间', '到期时间', '是否认证',
                      '是否审核', '审核时间', '营业天数', '发布职位数', 'VIP类型',
                      '是否过期', '是否暂停', '来源类型', '是否删除']
ADMIN_COMPANY_HIDDEN = {'公司简介', '编辑时间', '审核时间', '营业天数', '来源类型',
                        '是否过期', '是否暂停', '是否删除'}

# ----------------------------------------------------------------------
# 用户全量字段（总后台 /yunying/users/user_set.asp 会员管理页）
# ----------------------------------------------------------------------
ADMIN_USER_COLS = ['用户ID', '用户名', '昵称', '真实姓名', '会员姓名', '会员性别',
                   '会员年龄', '生日', '婚姻状况', '职业', '所在城市', '联系电话',
                   '固定电话', '联系QQ', '电子邮件', '邮编', '联系地址', '证件类型',
                   '证件号', '头像', '注册时间', '注册IP', '注册站点', '登录次数',
                   '最后登录时间', '最后登录IP', '账号状态', '论坛头衔', '论坛成长值',
                   '简介', '个性签名', '会员服务', '用户链接']
ADMIN_USER_HIDDEN = {'会员姓名', '会员年龄', '婚姻状况', '职业', '固定电话', '联系QQ',
                     '邮编', '联系地址', '证件类型', '证件号', '最后登录IP',
                     '论坛成长值', '简介', '个性签名', '会员服务'}

# ----------------------------------------------------------------------
# URL
# ----------------------------------------------------------------------
def resume_url(jlid):
    return f'{FRONT_BASE}/post/qiuzhi/{jlid}x.html'


def job_url(jobid):
    return f'{FRONT_BASE}/post/zhaopin/{jobid}x.html'


def company_admin_url(company_id):
    return f'{JOB_ADMIN_BASE}/Enterprise/CompanyAdd?source=1&companyId={company_id}'


def user_admin_url(user_id):
    return f'{MASTER_ADMIN_BASE}/yunying/users/user_set.asp?id={user_id}'


def user_front_url(hometown_id):
    if not hometown_id:
        return ''
    return (f'{FRONT_BASE}/forum/user/index.asp?username={hometown_id}')


KIND_URL = {
    'resume': lambda oid: f'{JOB_ADMIN_BASE}/Post/JlEdit?source=1&jlid={oid}',
    'job': lambda oid: f'{JOB_ADMIN_BASE}/Post/JobEdit?Source=1&id={oid}',
    'company': lambda oid: company_admin_url(oid),
    'user': lambda oid: user_admin_url(oid),
}
KIND_COL = {'resume': '简历ID', 'job': '岗位ID', 'company': '企业ID', 'user': '用户ID'}


# ----------------------------------------------------------------------
# 通用文本工具
# ----------------------------------------------------------------------
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


def _strip_paren(s):
    s = (s or '').strip()
    while len(s) >= 2 and s[0] in '（(' and s[-1] in '）)':
        s = s[1:-1].strip()
    return s


def _abs_url(u):
    """//img.xxx/... -> https://img.xxx/...；/path -> http://www.hbtc.ccoo.cn/path"""
    if not u:
        return ''
    u = u.strip()
    if u.startswith('data:'):
        return u
    if u.startswith('//'):
        return 'https:' + u
    if u.startswith('/'):
        return FRONT_BASE + u
    return u


_RE_DOTNET = re.compile(r'^/Date\((-?\d+)\)/$')


def _fmt_dt(v):
    """把 /Date(ms)/ 或纯毫秒时间戳格式化成 yyyy-mm-dd hh:mm:ss；其它原样返回"""
    if not v:
        return ''
    s = str(v).strip()
    if not s:
        return ''
    m = _RE_DOTNET.match(s)
    if m:
        v = m.group(1)
    if re.fullmatch(r'\d{12,14}', s):
        v = s
    else:
        return s
    try:
        from datetime import datetime, timezone, timedelta
        tz = timezone(timedelta(hours=8))
        return datetime.fromtimestamp(int(v) / 1000, tz).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return s


# ----------------------------------------------------------------------
# 表单控件取值（后台 layui 编辑页通用）
# ----------------------------------------------------------------------
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


def _checked_title(tag, name):
    """取同 name 下所有勾选控件的 title/value（求职状态单选、亮点复选）"""
    vals = []
    for el in tag.find_all(attrs={'name': name}):
        if el.has_attr('checked'):
            vals.append(el.get('title') or el.get('value') or '')
    return vals


def _form_map(soup):
    """
    收集页面所有表单控件的 name -> 值。
    复选/单选：值为勾选项的 title/value 列表；输入类：值为 value/选中文本。
    """
    m = {}
    for el in soup.find_all(['input', 'textarea', 'select']):
        name = el.get('name') or el.get('id')
        if not name:
            continue
        if el.name == 'textarea':
            v = (el.get_text() or '').strip()
            m[name] = v
        elif el.name == 'select':
            m[name] = _sel(el)
        else:
            t = (el.get('type') or 'text').lower()
            if t in ('checkbox', 'radio'):
                if el.has_attr('checked'):
                    v = el.get('title') or el.get('value') or ''
                    m.setdefault(name, []).append(str(v).strip())
                continue
            v = (el.get('value') or '').strip()
            if v or name not in m:
                m[name] = v
    return m


def _fv(m, *names):
    """从表单 map 按多个候选 name 取第一个非空值（支持列表合并）"""
    for n in names:
        v = m.get(n)
        if isinstance(v, list):
            v = ','.join(str(x) for x in v if str(x).strip())
        if v:
            return str(v).strip()
    return ''


def _table_pairs(soup):
    """
    通用「标签：值」表格解析（后台编辑页的只读摘要区 / 总后台详情页）。
    识别 <td align="right">标签：</td> 后随的 <td>值</td>。
    返回 {标签: 值}（标签去掉尾部冒号）。
    """
    pairs = {}
    for td in soup.find_all('td'):
        if (td.get('align') or '').strip() != 'right':
            continue
        txt = _safe_text(td)
        m = re.match(r'^(.*?)[：:]\s*$', txt)
        if not m:
            continue
        label = m.group(1).strip()
        if not label:
            continue
        sib = td.find_next_sibling('td')
        value = clean_rich(sib) if sib is not None else ''
        if label not in pairs or pairs[label] == '':
            pairs[label] = value
    return pairs


# ----------------------------------------------------------------------
# 登录页 / 无效页判断
# ----------------------------------------------------------------------
_ADMIN_LOGIN_MARKS = ('LAY-user-login-username', '请输入账号', 'login-form')


def is_admin_login(html):
    """job.admin.ccoo.cn 详情页是否返回了登录页"""
    if not html:
        return True
    return any(m in html for m in _ADMIN_LOGIN_MARKS)


_MASTER_LOGIN_MARKS = ('verify.asp', '运营管理平台', '立即登录')


def is_master_login(html):
    """www.admin.ccoo.cn 总后台页面是否返回了登录页"""
    if not html:
        return True
    return any(m in html for m in _MASTER_LOGIN_MARKS)


def is_login_page(html, kind='resume'):
    return is_master_login(html) if kind == 'user' else is_admin_login(html)


# ----------------------------------------------------------------------
# 简历解析（后台 /Post/JlEdit 编辑页）
# ----------------------------------------------------------------------
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


def parse_admin_resume(html, jlid=''):
    """解析后台 /Post/JlEdit 简历编辑页，返回全量字段 dict（页面无效返回 None）"""
    soup = BeautifulSoup(html, 'html.parser')
    # 有效页判据：存在简历表单（姓名/姓名拼音等 input 控件）
    if not (soup.find('input', attrs={'name': 'truename'})
            or soup.find('input', attrs={'name': 'title'})
            or soup.find('ul', class_='js-work-list')):
        return None

    d = {c: '' for c in DETAIL_COLS}
    d['简历ID'] = jlid

    m = _form_map(soup)
    d['姓名'] = _fv(m, 'truename')
    d['性别'] = _fv(m, 'Sex')
    d['学历'] = _fv(m, 'edu')
    d['工作经验'] = _fv(m, 'record')
    d['现居地'] = _fv(m, 'juzhudi', 'lifeaddr')
    d['期望职位'] = _fv(m, 'title')
    d['期望月薪'] = _fv(m, 'pay')
    d['职位类别'] = _fv(m, 'kindid', 'ClassId', 'kind', 'positionType')
    d['自我描述标签'] = ', '.join([t for t in _checked_title(soup, 'liangdian') if t])

    el = soup.find('input', id='text_merecord')
    if el is not None:
        d['自我描述详情'] = strip_html((el.get('value') or '').strip())

    # 求职状态（radio/select 均可）
    st = _checked_title(soup, 'jlstate')
    if st:
        d['求职状态'] = st[0]
    else:
        d['求职状态'] = _fv(m, 'jlstate')

    # 年龄：出生年份 -> 当前年龄
    by = _fv(m, 'bornyear', 'bornYear')
    mby = re.search(r'(19|20)\d{2}', by)
    if mby:
        from datetime import datetime
        age = datetime.now().year - int(mby.group(0))
        if 10 <= age <= 90:
            d['年龄'] = f'{age}岁'

    # 头像 / 照片（URL 取原样；base64 图片过大不进 Excel，标 [图片]）
    for src_name, col in (('headPic', '头像'), ('UpPic', '照片'), ('UserImg', '头像')):
        v = _fv(m, src_name)
        if v:
            if v.startswith('data:image'):
                d[col] = '[图片]' if not d.get(col) else d[col]
            else:
                d[col] = _abs_url(v)

    # 工作经历（限定 js-work-list 区块）
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
        says.append(_sel(el) if el.name == 'select' else (el.get('value') or '').strip())
    for el in soup.find_all(attrs={'name': 'ReadAbility'}):
        reads.append(_sel(el) if el.name == 'select' else (el.get('value') or '').strip())
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
        if not n:
            continue
        cer_items.append(f'{n}（{t}）' if t else n)
    for i in range(CERT_COLS):
        d[f'获得证书{i + 1}'] = cer_items[i] if i < len(cer_items) else ''

    return d


# ----------------------------------------------------------------------
# 岗位解析（后台 /Post/JobEdit 编辑页）
# ----------------------------------------------------------------------
_JOB_NAME_MAP = {
    '岗位标题': ('title', 'JobTitle'),
    '职位类别': ('ClassName', 'kindName', 'kind', 'kindid', 'classid'),
    '学历要求': ('education', 'eduname'),
    '招聘人数': ('number', 'num'),
    '工作经验': ('record', 'expname', 'exp'),
    '工作区域': ('workarea', 'workArea'),
    '公司名称': ('compname', 'CompName'),
    '公司地址': ('Compaddr', 'compaddr'),
    '公司行业': ('ComptradeStr', 'Comptrade', 'tradeName'),
    '公司规模': ('CompSizeStr', 'compsize', 'compSize'),
    '公司性质': ('compkind', 'Compkind', 'CompkindStr', 'comptype'),
    '联系人': ('linkman', 'oLinkMan'),
    '联系电话': ('tel', 'oLinkTel', 'bdTel'),
    '邮箱': ('email',),
    '福利': ('fuli',),
    '职位描述': ('info', 'jobinfo', 'content'),
    '公司简介': ('compinfo',),
    '发布时间': ('systime', 'AddTimeStr', 'addtime'),
    '更新时间': ('edittime', 'UpTimeStr', 'updateTime'),
    '审核状态': ('checkstate', 'ccoochk', 'unchk'),
    '是否下线': ('IsExit', 'isExit', 'isdel'),
    '点击量': ('hit', 'hitnum'),
}

# 前台岗位详情页字段与后台编辑页字段的兼容映射（已并入 _JOB_NAME_MAP 的候选名）
_JOB_NAME_MAP['薪资'] = ('salary', 'pay', 'xinzhi')
_JOB_NAME_MAP['岗位链接'] = ('_link_',)


def _job_salary(m):
    sal = _fv(m, 'salary', 'pay', 'xinzhi')
    if sal:
        return sal
    lo = _fv(m, 'minsalary', 'minSalary')
    hi = _fv(m, 'maxsalary', 'maxSalary')
    if lo or hi:
        return f'{lo}-{hi}'
    return ''


def parse_admin_job(html, jobid=''):
    """解析后台 /Post/JobEdit 岗位编辑页，返回全量字段 dict（页面无效返回 None）"""
    soup = BeautifulSoup(html, 'html.parser')
    m = _form_map(soup)
    # 有效页判据：存在岗位标题 input 或表单区
    if not (_fv(m, 'title', 'JobTitle') or soup.find('form')
            or soup.find('textarea', attrs={'name': 'info'})):
        return None

    d = {c: '' for c in ADMIN_JOB_COLS}
    d['岗位ID'] = jobid
    for col in ADMIN_JOB_COLS:
        if col in ('岗位ID', '岗位链接', '薪资'):
            continue
        d[col] = _fv(m, *_JOB_NAME_MAP.get(col, ()))
    d['薪资'] = _job_salary(m)
    d['发布时间'] = _fmt_dt(d['发布时间'])
    d['更新时间'] = _fmt_dt(d['更新时间'])

    # 勾选类字段（是否下线 / 审核状态等可能用 radio）
    for col, names in (('是否下线', ('IsExit', 'isExit')),
                       ('审核状态', ('checkstate', 'ccoochk'))):
        for n in names:
            v = m.get(n)
            if isinstance(v, list):
                d[col] = ','.join(str(x) for x in v)
                break
            if v:
                d[col] = str(v)
                break

    # 岗位链接（前台展示链接）
    d['岗位链接'] = job_url(jobid) if jobid else ''

    # 表格摘要兜底：有些字段以「标签：值」只读展示
    tp = _table_pairs(soup)
    if tp:
        for col, keys in (('职位描述', ('职位描述', '岗位描述', '招聘描述')),
                          ('公司简介', ('公司简介',)),
                          ('公司行业', ('公司行业',)),
                          ('公司规模', ('公司规模',)),
                          ('公司性质', ('公司性质',)),
                          ('学历要求', ('学历要求',)),
                          ('工作经验', ('工作经验',)),
                          ('招聘人数', ('招聘人数',))):
            if not d[col]:
                for k in keys:
                    if tp.get(k):
                        d[col] = tp[k]
                        break
    return d


# ----------------------------------------------------------------------
# 企业解析（后台 /Enterprise/CompanyAdd 编辑页）
# ----------------------------------------------------------------------
_COMPANY_NAME_MAP = {
    '企业名称': ('CompName', 'compname'),
    '企业用户名': ('username',),
    '联系人': ('oLinkMan', 'linkman'),
    '联系电话': ('oLinkTel', 'tel'),
    '公司简介': ('compinfo',),
    '公司行业': ('ComptradeStr', 'Comptrade', 'tradeName'),
    '公司类型': ('CompkindStr', 'Compkind', 'compkind'),
    '公司规模': ('CompSizeStr', 'CompSize', 'compsize'),
    '公司地址': ('Compaddr', 'compaddr'),
    '区域': ('AreaStr', 'Areaid', 'areaid'),
    '注册时间': ('SysTime', 'systime', 'CreateTime'),
    '编辑时间': ('EditTime', 'edittime'),
    '到期时间': ('expTime', 'OEndTIme', 'OEndTime'),
    '是否认证': ('olrz', 'isVerify'),
    '是否审核': ('isVerify', 'checkstate', 'IsVerify'),
    '审核时间': ('verifyTime',),
    '营业天数': ('opendays',),
    '发布职位数': ('jobCount',),
    'VIP类型': ('VipType', 'vipType'),
    '是否过期': ('HasExpire', 'hasexpire'),
    '是否暂停': ('IsPause', 'ispause'),
    '来源类型': ('fromtype',),
    '是否删除': ('isdel',),
}


def parse_admin_company(html, company_id=''):
    """解析后台 /Enterprise/CompanyAdd 企业编辑页，返回全量字段 dict"""
    soup = BeautifulSoup(html, 'html.parser')
    m = _form_map(soup)
    if not (_fv(m, 'CompName', 'compname') or soup.find('form')):
        return None
    d = {c: '' for c in ADMIN_COMPANY_COLS}
    d['企业ID'] = company_id
    for col in ADMIN_COMPANY_COLS:
        if col == '企业ID':
            continue
        d[col] = _fv(m, *_COMPANY_NAME_MAP.get(col, ()))
    for col in ('注册时间', '编辑时间', '到期时间', '审核时间'):
        d[col] = _fmt_dt(d[col])
    # 勾选类字段
    for col, names in (('是否认证', ('olrz', 'isVerify', 'is_verify')),
                       ('是否审核', ('isVerify', 'checkstate')),
                       ('是否过期', ('HasExpire', 'hasexpire')),
                       ('是否暂停', ('IsPause', 'ispause')),
                       ('是否删除', ('isdel',))):
        for n in names:
            v = m.get(n)
            if isinstance(v, list):
                d[col] = ','.join(str(x) for x in v)
                break
            if v:
                d[col] = str(v)
                break
    # 表格摘要兜底
    tp = _table_pairs(soup)
    if tp:
        for col, keys in (('公司简介', ('公司简介',)),
                          ('公司行业', ('公司行业',)),
                          ('公司规模', ('公司规模',)),
                          ('公司地址', ('公司地址',)),
                          ('联系人', ('联系人',)),
                          ('联系电话', ('联系电话',))):
            if not d[col]:
                for k in keys:
                    if tp.get(k):
                        d[col] = tp[k]
                        break
    return d


# ----------------------------------------------------------------------
# 用户解析（总后台 /yunying/users/user_set.asp 会员管理页，已实测页面结构）
# ----------------------------------------------------------------------
_USER_LABEL_MAP = {
    '用户名': '用户名',
    '昵称': '昵称',
    '真实姓名': '真实姓名',
    '会员姓名': '会员姓名',
    '会员性别': '会员性别',
    '婚姻状况': '婚姻状况',
    '职业': '职业',
    '所在城市': '所在城市',
    '联系电话': '联系电话',
    '固定电话': '固定电话',
    '联系QQ': '联系QQ',
    '电子邮件': '电子邮件',
    '邮编': '邮编',
    '联系地址': '联系地址',
    '证件类型': '证件类型',
    '证件号': '证件号',
    '注册时间': '注册时间',
    '注册IP': '注册IP',
    '注册站点': '注册站点',
    '登录次数': '登录次数',
    '最后登录时间': '最后登录时间',
    '最后登录IP': '最后登录IP',
    '账号状态': '账号状态',
    '论坛头衔': '论坛头衔',
    '论坛成长值': '论坛成长值',
    '简介': '简介',
    '个性签名': '个性签名',
}

# 会员服务开关（页面「会员服务」表格中的开通/绑定状态）
_SERVICE_LABELS = ('房屋中介', '名企招聘', '论坛登录', '车行',
                   '二手商家', '新楼盘管理', '分类信息发布', '微信登录')


def parse_admin_user(html, user_id=''):
    """解析总后台 user_set.asp 会员管理页，返回全量字段 dict（纯文本，无需 OCR）"""
    soup = BeautifulSoup(html, 'html.parser')
    # 有效页判据：存在「注册信息 / 个人信息」区块
    if not ('注册信息' in html or '个人信息' in html):
        return None

    d = {c: '' for c in ADMIN_USER_COLS}
    d['用户ID'] = user_id

    tp = _table_pairs(soup)
    for label, col in _USER_LABEL_MAP.items():
        if tp.get(label) is not None:
            d[col] = tp[label]

    # 注册IP 值形如 117.154.28.195[查看IP所在地]，去掉链接文字
    rip = d.get('注册IP') or ''
    if '[' in rip:
        d['注册IP'] = rip.split('[', 1)[0].strip()

    # 会员年龄：31岁（生日：1995/1/1）
    age_v = tp.get('会员年龄') or ''
    ma = re.match(r'(\d+)', age_v)
    if ma:
        d['会员年龄'] = ma.group(1)
    mb = re.search(r'生日[：:]\s*([\d/\-]+)', age_v)
    if mb:
        d['生日'] = mb.group(1)

    # 头像（会员头像行右侧 img）
    for td in soup.find_all('td'):
        if (td.get('align') or '').strip() == 'right' and _safe_text(td) == '会员头像：':
            nxt = td.find_next_sibling('td')
            img = nxt.find('img') if nxt is not None else None
            if img is not None:
                d['头像'] = _abs_url(img.get('src') or '')
            break

    # 会员服务：合并各开关状态，如「房屋中介:未开通; 论坛登录:正常状态; …」
    svc = []
    for lab in _SERVICE_LABELS:
        v = tp.get(lab) or ''
        v = v.replace('×', '').replace('√', '').replace('&nbsp;', '').strip()
        if v:
            svc.append(f'{lab}:{v}')
    d['会员服务'] = '; '.join(svc)

    # 用户链接：由用户名（ttid…）拼前台论坛主页
    d['用户链接'] = user_front_url(d.get('用户名'))

    return d


# ----------------------------------------------------------------------
# 抓取（后台无 WAF，不做全局限速；保留失败重试与自动补抓）
# ----------------------------------------------------------------------
_TLS = threading.local()
_KIND_PARSE = {
    'resume': parse_admin_resume,
    'job': parse_admin_job,
    'company': parse_admin_company,
    'user': parse_admin_user,
}


def _session(cookie_str=None):
    s = getattr(_TLS, 'session', None)
    if s is None:
        s = requests.Session()
        s.trust_env = False          # 绕过本机代理，避免 127.0.0.1 被劫持
        _TLS.session = s
    if cookie_str:
        s.cookies = requests.cookies.cookiejar_from_dict(normalize_cookie(cookie_str))
    return s


def fetch_one(objid, cookie_str, timeout=25, retries=3, session=None, kind='resume'):
    """抓取并解析单个对象（后台编辑页 / 总后台用户页）。

    kind: 'resume' 简历 JlEdit / 'job' 岗位 JobEdit /
          'company' 企业 CompanyAdd / 'user' 用户 user_set.asp（需总后台 Cookie）
    成功返回 dict，失败返回 {'_error': ...}。
    """
    url = KIND_URL[kind](objid)
    idkey = KIND_COL[kind]
    last = None
    for attempt in range(max(1, retries)):
        try:
            s = session or _session(cookie_str)
            r = s.get(url, headers=ADMIN_HEADERS, timeout=timeout)
            r.encoding = 'gb18030' if kind == 'user' else 'utf-8'
            if r.status_code != 200:
                last = f'HTTP {r.status_code}'
                time.sleep(0.4 * (attempt + 1))
                continue
            html = r.text
            if is_login_page(html, kind):
                return {'_error': 'need_login', idkey: objid}
            d = _KIND_PARSE[kind](html, objid)
            if d is None:
                last = '页面无内容'
                time.sleep(0.4 * (attempt + 1))
                continue
            return d
        except Exception as e:
            last = f'{type(e).__name__}: {e}'
            time.sleep(0.4 * (attempt + 1))
    return {'_error': str(last or 'unknown'), idkey: objid}


def fetch_many(ids, cookie_str, workers=8, on_progress=None, should_stop=None,
               timeout=25, retries=3, retry_rounds=2, retry_delay=2.0, kind='resume'):
    """
    并发抓取多个对象的全量字段（后台页面，不设前台限速）。

    kind: 'resume' / 'job' / 'company' / 'user'
    out  : {id: detail_dict}    只含成功的
    stats: {'done','ok','fail','need_login','total','stopped','rounds'}

    on_progress(done, total, ok, fail) 每完成一个回调一次
    should_stop() 返回 True 时提前结束

    可靠抓取策略：
      1) 全并发（workers 默认 8），后台页面无前台 WAF，速度快；
      2) 失败自动补抓：第一轮结束后，对失败的 id 再跑 `retry_rounds` 轮，
         轮间等待 `retry_delay` 秒；
      3) 无需登录的请求直接返回 need_login 统计，由调用方提示 Cookie 问题。
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    total = len(ids)
    if not total:
        return {}, {'done': 0, 'ok': 0, 'fail': 0, 'need_login': 0,
                    'total': 0, 'stopped': False, 'rounds': 0}

    out = {}
    need_login = 0
    lock = threading.Lock()
    pending = list(ids)
    rounds = 0
    stopped = False
    ok_total = 0

    def _round(items):
        nonlocal need_login
        failed = []

        def work(oid):
            one = requests.Session()
            one.trust_env = False
            one.cookies = requests.cookies.cookiejar_from_dict(normalize_cookie(cookie_str))
            return oid, fetch_one(oid, cookie_str, timeout=timeout,
                                  retries=retries, session=one, kind=kind)

        done_cnt = [0]
        stop_flag = [False]
        with ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
            futs = [ex.submit(work, i) for i in items]
            for fu in as_completed(futs):
                if stop_flag[0] or (should_stop and should_stop()):
                    for f in futs:
                        f.cancel()
                    stop_flag[0] = True
                    break
                try:
                    oid, res = fu.result()
                except Exception as e:
                    oid, res = None, {'_error': f'{type(e).__name__}: {e}'}
                with lock:
                    done_cnt[0] += 1
                    err = (res or {}).get('_error')
                    if res and not err:
                        out[oid] = res
                    else:
                        failed.append(oid)
                        if err == 'need_login':
                            need_login += 1
                    if on_progress:
                        on_progress(ok_total + done_cnt[0], total, len(out),
                                    total - len(out))
        return failed, stop_flag[0]

    for rnd in range(max(1, retry_rounds + 1)):
        if not pending:
            break
        if rnd > 0:
            if should_stop and should_stop():
                stopped = True
                break
            t_end = time.time() + retry_delay * rnd
            while time.time() < t_end:
                if should_stop and should_stop():
                    break
                time.sleep(min(0.5, t_end - time.time()))
        rounds = rnd + 1
        ok_total = len(out)
        failed, stop_flag = _round(pending)
        ok_total = len(out)
        pending = [i for i in pending if i not in out]
        if stop_flag:
            stopped = True
            break
        if not pending:
            break

    fail_n = total - len(out)
    if on_progress:
        on_progress(total, total, len(out), fail_n)
    return out, {'done': total, 'ok': len(out), 'fail': fail_n,
                 'need_login': need_login, 'total': total,
                 'stopped': stopped, 'rounds': rounds}


def test_master_cookie(cookie_str, probe_uid='79557050'):
    """用一份已知用户试抓总后台 user_set.asp，判断「总后台 Cookie」是否有效。

    返回 (ok, msg, sample_nick)。
    """
    if not normalize_cookie(cookie_str):
        return False, '未填写总后台 Cookie', ''
    res = fetch_one(probe_uid, cookie_str, timeout=20, retries=2, kind='user')
    if not res:
        return False, '请求失败（网络或超时），请重试', ''
    err = res.get('_error')
    if err == 'need_login':
        return False, '总后台 Cookie 已失效，请重新登录 www.admin.ccoo.cn 后复制最新 Cookie', ''
    if err:
        return False, f'抓取失败：{err}', ''
    nick = res.get('昵称') or res.get('用户名') or ''
    return True, f'有效（示例用户：{nick or "—"}）', nick


if __name__ == '__main__':
    import sys
    kind = sys.argv[1] if len(sys.argv) > 1 else 'resume'
    ck = sys.argv[2] if len(sys.argv) > 2 else ''
    ids = sys.argv[3:] or (['79557050'] if kind == 'user' else
                           ['2764068'] if kind == 'resume' else
                           ['8028715'] if kind == 'job' else ['1403037'])
    for oid in ids:
        d = fetch_one(oid, ck, retries=2, kind=kind)
        if not d or d.get('_error'):
            print(oid, '-> 失败', d)
            continue
        cols = {'resume': DETAIL_COLS, 'job': ADMIN_JOB_COLS,
                'company': ADMIN_COMPANY_COLS, 'user': ADMIN_USER_COLS}[kind]
        print(f'\n=== {oid} ===')
        for c in cols:
            if d.get(c):
                print(f'  {c}: {str(d[c])[:120]}')
