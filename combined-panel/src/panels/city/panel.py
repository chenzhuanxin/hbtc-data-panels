#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
城市中国·湖北通城网 招聘频道 平台数据面板
==========================================
运行：python 面板.py   （自动打开浏览器 http://127.0.0.1:8787）

功能：
  1. 抓取 7 个类目数据（订单管理/全职岗位/兼职岗位/企业库/会员企业/简历库/注册用户）
  2. base64 电话图片自动 OCR（RapidOCR 离线识别）后填充到表格
  3. /Date(...)/ 时间统一转为 yyyy-mm-dd hh:mm:ss
  4. Excel 导出：勾选类目、选择目录、自动命名；单类目=一个 xlsx，多类目=zip 包
  5. Excel 保留全部字段，次要字段自动隐藏列（对照参考 Excel 的隐藏配置）
  6. 账号/密码/Cookie 保存在本目录 config.json，支持导入/导出
"""

import os
import re
import io
import sys
import json
import time
import base64
import hashlib
import zipfile
import threading
import traceback
import webbrowser
from datetime import datetime, timezone, timedelta

import requests

import resume_front as RF

sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))   # 本文件位于 panels/xx/


def _is_frozen():
    return bool(getattr(sys, 'frozen', False))


def _combo_root():
    """综合面板根目录：源码模式 = 项目根；exe 模式 = exe 所在目录。"""
    if _is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


# ----------------------------------------------------------------------
# 本模块是「综合数据面板」的内置子面板（源自独立项目的 面板.py）：
#   * static 资源放在本模块同目录，源码 / exe 解包后路径一致
#   * 配置文件与综合面板共享（根目录 config.json，两个子面板的键完全兼容，
#     后台 Cookie / 前台 Cookie / 保存目录 只需填一次）
#   * 各自的抓取缓存放在独立子目录，互不干扰
# ----------------------------------------------------------------------
RES_DIR = BASE_DIR
DATA_DIR = os.path.join(_combo_root(), '数据_平台面板')
CACHE_DIR = os.path.join(DATA_DIR, 'data_cache')
CONFIG_PATH = os.path.join(_combo_root(), 'config.json')
os.makedirs(CACHE_DIR, exist_ok=True)

BASE_URL = 'http://job.admin.ccoo.cn'
LOGIN_URL = BASE_URL + '/Home/login'
VALIDATE_CODE_URL = 'https://vcodeapi.bccoo.cn/home/getvalidcode'

TZ = timezone(timedelta(hours=8))
DT_MIN_MS = -62135596800000  # .NET DateTime.MinValue

# ----------------------------------------------------------------------
# 类目配置：url / params / 表头（英文字段 -> 中文）/ 隐藏列（参照参考 Excel）
# ----------------------------------------------------------------------
API_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
    'Referer': BASE_URL + '/',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'X-Requested-With': 'XMLHttpRequest',
}

PARAMS_ORDER = {
    '订单管理': {
        'url': '/Order/AllIncomeReportDataNew',
        'params': {'keyword': '', 'keyType': 0, 'orderSource': 0, 'orderType': 0,
                   'followUserId': 0, 'date': '', 'rank': 0, 'tjprice': 0,
                   'Integral': -1, 'invoice': -1},
    },
    '全职岗位': {
        'url': '/Post/GetYZJobList',
        'params': {'key': '', 'kindId': '', 'IsCmd': -1, 'fromtime': '2-365',
                   'fromtype': -1, 'IsExit': -1, 'istjusername': '', 'paixu': 0,
                   'selCcoochk': -1, 'classid': -1, 'hasSend': 0, 'ccoochk': 1,
                   'openSource': 0, 'username': '', 'zphjoblistType': 1, 'zphid': 0},
    },
    '兼职岗位': {
        'url': '/Post/GetJZ_YZJobList',
        'params': {'key': '', 'kindId': -1, 'IsCmd': -1, 'IsExit': -1,
                   'istjusername': '', 'paixu': 0, 'selCcoochk': -1, 'ccoochk': 1},
    },
    '企业库': {
        'url': '/Enterprise/GetCompanyListData',
        'params': {'keyword': '', 'keyType': 0, 'openSource': 0},
    },
    '会员企业': {
        'url': '/Enterprise/GetEnterpriseData',
        'params': {'keyword': '', 'keyType': ''},
    },
    '简历库': {
        'url': '/Post/GetPostJlList',
        'params': {'keyWord': '', 'kindId': '', 'ClassId': -1, 'record': 0,
                   'Edu': 0, 'iscmd': 0, 'Videochk': 0, 'jlstate': -1,
                   'jlPrivacy': 0, 'age': 0, 'lev': 0, 'label': '',
                   'istjusername': '', 'paixu': 0, 'checkstate': 0,
                   'mainSiteId': 0, 'telContact': 0, 'wechartContact': 0,
                   'sex': 0, 'fromtype': 0, 'isreport': 0, 'ischk': 2,
                   'openSource': 0},
    },
    '注册用户': {
        'url': '/Customer/GetUserInfoList',
        'params': {'openSource': 0},
    },
}

# 各类目字段（英文字段顺序 = 接口返回顺序，中文 = 后台表头）
HEAD_ORDER = {
    '订单管理': ['id', 'compName', 'CustId', 'crmId', 'Name', 'username', 'price',
              'mark', 'systime', 'IncomeType', 'title', 'tjday', 'incomesource',
              'OrderNo', 'JobTitle', 'jlTitle', 'Integral', 'tgjMoney',
              'LAY_CHECKED', 'useMoney', 'invoice'],
    '全职岗位': ['AsClassName', 'AsCmd', 'countjl', 'countRe', 'fromtype',
              'checkstate', 'hitnum', 'AsIsRecoms', 'ClassName', 'ccoochkReason',
              'bdTel', 'ComptradeStr', 'regSiteid', 'fbUserStr', 'fbIPStr',
              'fbUserTel', 'bdUserTel', 'telArea', 'VipType', 'unchk', 'iserror',
              'FromTypeStr', 'AddTimeStr', 'UpTimeStr', 'reportCount',
              'followUser', 'users_lh', 'bdtel_lh', 'tel_lh', 'id', 'ISTJING',
              'ccoochk', 'title', 'linkman', 'tel', 'email', 'fuli', 'ZC', 'ZC2',
              'username', 'edittime', 'systime', 'istjtime', 'istjDoTime', 'hit',
              'IsExit', 'jlIsViewCount', 'kind', 'classid', 'kind_back', 'info',
              'workarea', 'compname', 'Compaddr', 'Comptrade', 'compkind',
              'compsize', 'istjusername', 'callCount', 'IsCmd', 'compinfo',
              'siteid', 'IsSite', 'ip', 'education', 'record', 'number',
              'minsalary', 'maxsalary', 'salary', 'UpPic', 'infoNum', 'LCZtype',
              'isdel', 'AreaId', 'LCZusername', 'oMap', 'qyid', 'isHideTel',
              'updateTime', 'jobMap', 'jobComppointTag', 'airs', 'isxcxshow'],
    '兼职岗位': None,  # 与全职岗位相同
    '企业库': ['id', 'Siteid', 'CompName', 'oPx', 'SysTime', 'username',
             'jobCount', 'oLinkMan', 'oLinkTel', 'compinfo', 'VipType',
             'oldVipType', 'CmdTime', 'expTime', 'HasExpire', 'Areaid', 'AreaStr',
             'Compaddr', 'Comptrade', 'ComptradeStr', 'Compkind', 'CompkindStr',
             'CompSize', 'CompSizeStr', 'systime', 'isdel', 'fromtype', 'olrz',
             'isVerify', 'verifyTime', 'vipSystime', 'opendays', 'opendaysStr',
             'EditTime', 'IsPause', 'OEndTIme', 'verifyPic', 'CreateTime',
             'UseTime', 'HasUse'],
    '会员企业': None,  # 与企业库相同
    '简历库': ['AsImg', 'AsCmd', 'AsIsRecoms', 'livenessscore', 'reportcount',
             'areatitle', 'isBindWx', 'num', 'Report_state', 'Videochk',
             'VideoUrl', 'id', 'siteid', 'username', 'title', 'truename', 'sex',
             'bornyear', 'bornmonth', 'borndate', 'ismarry', 'record', 'edu',
             'jiguan', 'edukind', 'eduaddr', 'email', 'tel', 'qq', 'wish', 'pay',
             'workrecord', 'edurecord', 'merecord', 'systime', 'edittime',
             'jlstate', 'jlkind', 'ccoochk', 'UserImg', 'IsCmd', 'UpUserImg1',
             'UpUserImg2', 'UpUserImg3', 'ClassId', 'hit', 'UpPic', 'headPic',
             'fromtype', 'IsSite', 'istjtime', 'liangdian', 'juzhudi', 'areaId',
             'userType', 'postRemarks', 'IsUpby', 'LCZusername', 'LCZtype',
             'ischeckuser', 'ischecktime', 'ANum', 'ip', 'istjusername',
             'classidstr', 'jlkindstr', 'ccoochkReason', 'isContactTel',
             'isAddWx', 'checkstate', 'jl_level', 'jl_remark', 'jl_label',
             'jlPrivacy', 'telArea', 'aitags', 'inputkind'],
    '注册用户': ['site_name', 'bindtel', 'bindwx', 'qiyeku', 'jl', 'jlId',
              'stopTel', 'stopUser', 'isBZ', 'id', 'username', 'password', 'pwd',
              'type', 'regtime', 'regIP', 'atwhere', 'siteid', 'name', 'tel',
              'email', 'state', 'nick', 'Previous', 'LastTime', 'LastIp',
              'LogNum', 'JinFen', 'Coins', 'prov', 'city', 'town', 'RoleId',
              'LongId', 'friend_join', 'friend_total', 'mobile', 'QQID',
              'userface', 'sex', 'birthday', 'address', 'postcode', 'qq',
              'QQID2', 'position', 'positionName', 'lifeaddr', 'source'],
}
HEAD_ORDER['兼职岗位'] = HEAD_ORDER['全职岗位']
HEAD_ORDER['会员企业'] = HEAD_ORDER['企业库']

ZH_LABELS = {
    '订单管理': {'id': '订单ID', 'compName': '企业名称', 'CustId': '客户ID',
             'crmId': 'CRM ID', 'Name': '姓名', 'username': '用户名',
             'price': '金额', 'mark': '备注', 'systime': '下单时间',
             'IncomeType': '收入类型', 'title': '职位名称', 'tjday': '推荐天数',
             'incomesource': '收入来源', 'OrderNo': '订单编号',
             'JobTitle': '岗位标题', 'jlTitle': '简历标题', 'Integral': '积分',
             'tgjMoney': '推广金额', 'LAY_CHECKED': '是否选中（LAYUI专用）',
             'useMoney': '使用金额', 'invoice': '是否开票'},
    '全职岗位': {'AsClassName': '所属分类名称', 'AsCmd': '是否推荐（Cmd）',
              'countjl': '简历投递数', 'countRe': '查看数/回复数',
              'fromtype': '来源类型', 'checkstate': '审核状态', 'hitnum': '点击量',
              'AsIsRecoms': '是否推荐位', 'ClassName': '分类名称',
              'ccoochkReason': '审核原因', 'bdTel': '绑定电话',
              'ComptradeStr': '公司行业（字符串）', 'regSiteid': '注册站点ID',
              'fbUserStr': '发布人（字符串）', 'fbIPStr': '发布IP',
              'fbUserTel': '发布人电话', 'bdUserTel': '绑定用户电话',
              'telArea': '电话区号', 'VipType': 'VIP类型', 'unchk': '未审核标识',
              'iserror': '是否错误', 'FromTypeStr': '来源类型（字符串）',
              'AddTimeStr': '添加时间（字符串）', 'UpTimeStr': '更新时间（字符串）',
              'reportCount': '举报次数', 'followUser': '跟进人',
              'users_lh': '用户最后登录时间', 'bdtel_lh': '绑定电话最后操作时间',
              'tel_lh': '电话最后操作时间', 'id': '岗位ID',
              'ISTJING': '是否推荐中', 'ccoochk': '审核状态（数字）',
              'title': '标题', 'linkman': '联系人', 'tel': '联系电话',
              'email': '邮箱', 'fuli': '福利', 'ZC': '注册（或职位类型）',
              'ZC2': '注册2', 'username': '用户名', 'edittime': '编辑时间',
              'systime': '发布时间', 'istjtime': '推荐开始时间',
              'istjDoTime': '推荐处理时间', 'hit': '点击量', 'IsExit': '是否下线',
              'jlIsViewCount': '简历查看次数', 'kind': '职位类别ID',
              'classid': '分类ID', 'kind_back': '职位类别备用', 'info': '职位描述',
              'workarea': '工作区域', 'compname': '公司名称',
              'Compaddr': '公司地址', 'Comptrade': '行业ID',
              'compkind': '公司类型ID', 'compsize': '公司规模ID',
              'istjusername': '推荐人用户名', 'callCount': '电话咨询量',
              'IsCmd': '是否推荐', 'compinfo': '公司简介', 'siteid': '站点ID',
              'IsSite': '是否主站', 'ip': 'IP地址', 'education': '学历要求',
              'record': '工作经验', 'number': '招聘人数', 'minsalary': '最低薪资',
              'maxsalary': '最高薪资', 'salary': '薪资（文本）',
              'UpPic': '上传图片', 'infoNum': '信息编号',
              'LCZtype': '拉勾/城市类型', 'isdel': '是否删除',
              'AreaId': '区域ID', 'LCZusername': '操作人用户名',
              'oMap': '地图坐标', 'qyid': '企业ID', 'isHideTel': '是否隐藏电话',
              'updateTime': '更新时间戳', 'jobMap': '职位地图信息',
              'jobComppointTag': '公司标签', 'airs': 'airs',
              'isxcxshow': '是否小程序显示'},
    '企业库': {'id': '企业ID', 'Siteid': '站点ID', 'CompName': '公司名称',
             'oPx': '排序权重', 'SysTime': '注册时间', 'username': '用户名',
             'jobCount': '发布职位数', 'oLinkMan': '联系人',
             'oLinkTel': '联系电话', 'compinfo': '公司简介',
             'VipType': 'VIP类型', 'oldVipType': '原VIP类型',
             'CmdTime': '推荐时间', 'expTime': '过期时间', 'HasExpire': '是否过期',
             'Areaid': '区域ID', 'AreaStr': '区域名称', 'Compaddr': '公司地址',
             'Comptrade': '行业ID', 'ComptradeStr': '行业名称',
             'Compkind': '公司类型ID', 'CompkindStr': '公司类型名称',
             'CompSize': '公司规模ID', 'CompSizeStr': '公司规模名称',
             'systime': '注册时间（重复）', 'isdel': '是否删除',
             'fromtype': '来源类型', 'olrz': '是否认证', 'isVerify': '是否审核',
             'verifyTime': '审核时间', 'vipSystime': 'VIP开通时间',
             'opendays': '营业天数', 'opendaysStr': '营业天数（文本）',
             'EditTime': '编辑时间', 'IsPause': '是否暂停', 'OEndTIme': '到期时间',
             'verifyPic': '认证图片', 'CreateTime': '创建时间',
             'UseTime': '使用时间', 'HasUse': '是否已使用'},
    '简历库': {'AsImg': '头像图片', 'AsCmd': '是否推荐', 'AsIsRecoms': '是否推荐位',
             'livenessscore': '活跃度评分', 'reportcount': '举报次数',
             'areatitle': '区域标题', 'isBindWx': '是否绑定微信', 'num': '编号',
             'Report_state': '举报状态', 'Videochk': '视频审核',
             'VideoUrl': '视频地址', 'id': '简历ID', 'siteid': '站点ID',
             'username': '用户名', 'title': '简历标题', 'truename': '真实姓名',
             'sex': '性别', 'bornyear': '出生年份', 'bornmonth': '出生月份',
             'borndate': '出生日期', 'ismarry': '婚姻状况', 'record': '学历',
             'edu': '教育背景', 'jiguan': '籍贯', 'edukind': '教育类型',
             'eduaddr': '教育地址', 'email': '邮箱', 'tel': '电话', 'qq': 'QQ号',
             'wish': '期望职位', 'pay': '期望薪资', 'workrecord': '工作经历',
             'edurecord': '教育经历', 'merecord': '其他记录',
             'systime': '创建时间', 'edittime': '更新时间', 'jlstate': '简历状态',
             'jlkind': '简历类型', 'ccoochk': '审核状态', 'UserImg': '用户头像',
             'IsCmd': '是否推荐', 'UpUserImg1': '用户图片1',
             'UpUserImg2': '用户图片2', 'UpUserImg3': '用户图片3',
             'ClassId': '分类ID', 'hit': '点击量', 'UpPic': '上传图片',
             'headPic': '头像', 'fromtype': '来源类型', 'IsSite': '是否主站',
             'istjtime': '推荐时间', 'liangdian': '亮点', 'juzhudi': '居住地',
             'areaId': '区域ID', 'userType': '用户类型', 'postRemarks': '备注',
             'IsUpby': '是否升级', 'LCZusername': '操作人',
             'LCZtype': '操作类型', 'ischeckuser': '审核人',
             'ischecktime': '审核时间', 'ANum': '认证编号', 'ip': 'IP地址',
             'istjusername': '推荐人', 'classidstr': '分类名称',
             'jlkindstr': '简历类型名称', 'ccoochkReason': '审核原因',
             'isContactTel': '是否可联系手机', 'isAddWx': '是否添加微信',
             'checkstate': '审核状态', 'jl_level': '简历等级',
             'jl_remark': '简历备注', 'jl_label': '简历标签',
             'jlPrivacy': '隐私设置', 'telArea': '电话区号', 'aitags': 'AI标签',
             'inputkind': '录入方式'},
    '注册用户': {'site_name': '站点名称', 'bindtel': '绑定手机',
              'bindwx': '绑定微信', 'qiyeku': '企业库标识', 'jl': '简历标识',
              'jlId': '简历ID', 'stopTel': '禁用电话', 'stopUser': '禁用用户',
              'isBZ': '是否备注', 'id': '用户ID', 'username': '用户名',
              'password': '密码（加密）', 'pwd': '密码（备用）',
              'type': '用户类型', 'regtime': '注册时间', 'regIP': '注册IP',
              'atwhere': '注册位置', 'siteid': '站点ID', 'name': '真实姓名',
              'tel': '电话', 'email': '邮箱', 'state': '状态', 'nick': '昵称',
              'Previous': '上次登录时间', 'LastTime': '最后登录时间',
              'LastIp': '最后登录IP', 'LogNum': '登录次数', 'JinFen': '积分',
              'Coins': '金币', 'prov': '省份', 'city': '城市', 'town': '乡镇',
              'RoleId': '角色ID', 'LongId': '长ID', 'friend_join': '好友加入数',
              'friend_total': '好友总数', 'mobile': '手机号', 'QQID': 'QQ号',
              'userface': '头像', 'sex': '性别', 'birthday': '生日',
              'address': '地址', 'postcode': '邮编', 'qq': 'QQ号（重复）',
              'QQID2': 'QQ号2', 'position': '职位', 'positionName': '职位名称',
              'lifeaddr': '生活地址', 'source': '来源'},
}
ZH_LABELS['兼职岗位'] = ZH_LABELS['全职岗位']
ZH_LABELS['会员企业'] = ZH_LABELS['企业库']

# 隐藏列（保留不删除，仅 Excel 中隐藏）——对照参考 Excel 的隐藏列配置
HIDDEN_COLS = {
    '订单管理': ['id', 'crmId', 'mark', 'IncomeType', 'tjday', 'OrderNo',
              'jlTitle', 'LAY_CHECKED', 'invoice'],
    '全职岗位': ['AsCmd', 'countRe', 'ccoochkReason', 'ISTJING', 'linkman',
              'email', 'istjDoTime', 'hit', 'Compaddr', 'jobMap'],
    '兼职岗位': ['AsCmd', 'countRe', 'VipType', 'ISTJING', 'linkman', 'email',
              'istjDoTime', 'hit', 'jobMap'],
    '企业库': ['Siteid', 'oPx', 'CmdTime', 'HasExpire', 'isdel', 'verifyTime',
             'vipSystime', 'opendays', 'EditTime', 'IsPause', 'verifyPic',
             'CreateTime', 'UseTime'],
    '会员企业': ['id', 'oPx', 'compinfo', 'CmdTime', 'HasExpire', 'isdel',
              'opendays', 'opendaysStr', 'EditTime', 'IsPause'],
    '简历库': ['AsImg', 'Report_state', 'siteid', 'bornyear', 'qq', 'jlkind',
             'fromtype', 'liangdian', 'ischecktime', 'ANum', 'aitags'],
    '注册用户': ['site_name', 'stopTel', 'pwd', 'regIP', 'Previous', 'LastIp',
              'JinFen', 'QQID', 'birthday', 'address'],
}

CATEGORIES = list(PARAMS_ORDER.keys())
UNAVAILABLE = {'账户订单管理': '接口不可用（与订单管理相同数据源，后台已停用）'}

# ----------------------------------------------------------------------
# 配置
# ----------------------------------------------------------------------
CONFIG_LOCK = threading.RLock()
DEFAULT_CONFIG = {'username': '', 'password': '', 'cookie': '',
                  'front_cookie': '',
                  'save_dir': '', 'ocr_scale': 1, 'fetch_workers': 4,
                  'front_rate': 1.0,               # 前台简历请求最小间隔（秒），1.0=1请求/秒
                  'fetch_resume_detail': 1,        # 简历库是否补抓前台简历详情
                  'resume_detail_limit': 0}        # 0 = 不限条数


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
# 数据仓库：内存 + 磁盘缓存
# ----------------------------------------------------------------------
DATA = {}          # cat -> {'rows': [[...],...], 'cols': [...], 'fetched_at': float, 'total': int}
DATA_LOCK = threading.RLock()
PROGRESS = {}      # cat -> {...}
JOB_LOCK = threading.Lock()
FETCH_JOBS = set()  # 正在抓取的类目
OCR_CACHE_PATH = os.path.join(CACHE_DIR, 'ocr_cache.json')
OCR_CACHE = {}
OCR_CACHE_LOCK = threading.Lock()

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


def cache_path(cat):
    return os.path.join(CACHE_DIR, cat + '.json')


def save_cache(cat):
    with DATA_LOCK:
        d = DATA.get(cat)
        if not d:
            return
        tmp = cache_path(cat) + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False)
        os.replace(tmp, cache_path(cat))


def load_all_caches():
    for cat in CATEGORIES:
        p = cache_path(cat)
        if os.path.exists(p):
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                if d.get('rows') is not None:
                    DATA[cat] = d
                    print(f'[缓存] 已加载 {cat}: {len(d["rows"])} 行')
            except Exception as e:
                print(f'[缓存] {cat} 加载失败: {e}')


# ----------------------------------------------------------------------
# 时间与字段处理
# ----------------------------------------------------------------------
RE_DOTNET_DATE = re.compile(r'^/Date\((-?\d+)\)/$')
FW_DIGITS = str.maketrans('０１２３４５６７８９', '0123456789')
NUMERIC_DATE_FIELDS = {'birthday'}  # 毫秒时间戳形式的日期字段

# Excel(xlsx) 底层是 XML，不能包含 XML 1.0 非法控制字符。
# openpyxl 写单元格时若命中就抛 IllegalCharacterError: "xxx cannot be used in worksheets."
# 允许保留：制表符 \x09、换行 \x0A、回车 \x0D（Excel 支持）
ILLEGAL_XML_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
EXCEL_MAX_LEN = 32767  # 单元格文本上限，超出 Excel 会报错

CLEAN_STAT = {'cells': 0, 'chars': 0}   # 导出时统计被清洗的数量


def clean_text(s):
    """去掉 XML 非法控制字符，并截断到 Excel 允许的最大长度"""
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
    """保证任意值都能安全写入 Excel 单元格"""
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
        return str(ms) if ms is not None else ''
    if ms <= DT_MIN_MS + 1000 or ms <= 0:
        return ''
    if ms > 4102444800000:  # > 2100 年，视为异常
        return ''
    dt = datetime.fromtimestamp(ms / 1000, TZ)
    return dt.strftime('%Y-%m-%d %H:%M:%S')


def process_value(key, v, ocr_fn):
    """单字段处理：时间转换 / base64图片 OCR / 其他"""
    if v is None:
        return ''
    if isinstance(v, (dict, list)):
        return clean_text(json.dumps(v, ensure_ascii=False))
    if isinstance(v, bool):
        return 1 if v else 0
    if isinstance(v, (int, float)):
        if key in NUMERIC_DATE_FIELDS:
            return ms_to_str(v)
        return v
    s = str(v)
    if s.startswith('data:image'):
        # OCR 结果也可能带控制字符，统一清洗
        return clean_text(ocr_fn(s)) if ocr_fn else '[图片]'
    m = RE_DOTNET_DATE.match(s.strip())
    if m:
        return ms_to_str(int(m.group(1)))
    return clean_text(s)


# ----------------------------------------------------------------------
# OCR（RapidOCR 离线，首次使用时懒加载）
# ----------------------------------------------------------------------
_OCR_ENGINE = None
_OCR_LOCK = threading.Lock()


def get_ocr_engine():
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        with _OCR_LOCK:
            if _OCR_ENGINE is None:
                print('[OCR] 正在加载 RapidOCR 引擎...')
                from rapidocr_onnxruntime import RapidOCR
                _OCR_ENGINE = RapidOCR()
                print('[OCR] RapidOCR 就绪')
    return _OCR_ENGINE


def ocr_base64_image(b64_str, scale=1):
    """解码 base64 图片 -> （可选放大）-> OCR -> 返回文本"""
    with OCR_CACHE_LOCK:
        key = hashlib.md5(f'{scale}|{b64_str}'.encode('utf-8')).hexdigest()
        if key in OCR_CACHE:
            return OCR_CACHE[key]
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
        if not res:
            text = ''
        else:
            text = ''.join(x[1] for x in res).translate(FW_DIGITS).replace(' ', '')
        del raw
    except Exception as e:
        print('[OCR] 识别失败:', e)
        text = ''
    with OCR_CACHE_LOCK:
        OCR_CACHE[key] = text
    return text


# ----------------------------------------------------------------------
# 会话
# ----------------------------------------------------------------------
def make_session(cookie_str=None):
    s = requests.Session()
    s.headers.update(API_HEADERS)
    ck = cookie_str if cookie_str is not None else CONFIG.get('cookie', '')
    if ck:
        d = {}
        for item in ck.split(';'):
            item = item.strip()
            if '=' in item:
                k, v = item.split('=', 1)
                d[k] = v
        s.cookies.update(d)
    return s


def test_cookie(cookie_str=None):
    """验证 Cookie 是否有效"""
    try:
        s = make_session(cookie_str)
        r = s.get(BASE_URL + '/Order/AllIncomeReportDataNew',
                  params={'page': 1, 'limit': 1}, timeout=15)
        try:
            j = r.json()
        except ValueError:
            return False, 'Cookie 无效（后台返回了登录页而非数据），请重新登录或粘贴新 Cookie'
        return j.get('code') == 0, f"连接成功，订单总数 {j.get('count', '?')}"
    except Exception as e:
        return False, f'验证失败: {e}'


# ----------------------------------------------------------------------
# 抓取任务
# ----------------------------------------------------------------------
def _set_progress(cat, **kw):
    with JOB_LOCK:
        p = PROGRESS.setdefault(cat, {})
        p.update(kw)


def fetch_category(cat, max_pages=0, use_ocr=True, with_detail=True):
    """后台抓取一个类目（含 OCR），写入 DATA 与磁盘缓存"""
    if cat not in PARAMS_ORDER:
        _set_progress(cat, phase='error', msg='未知类目')
        return
    ok, msg = test_cookie()
    if not ok:
        _set_progress(cat, phase='error', msg='Cookie 无效，请先在「连接设置」中配置有效的 Cookie 或重新登录')
        return

    scale = int(CONFIG.get('ocr_scale') or 1)
    ocr_fn = (lambda s: ocr_base64_image(s, scale)) if use_ocr else None

    cfg = PARAMS_ORDER[cat]
    canon = list(HEAD_ORDER[cat])
    cols = list(canon)
    rows = []
    session = make_session()
    ocr_total = 0
    t0 = time.time()
    _set_progress(cat, phase='fetching', pages_done=0, pages_total=None, rows=0,
                  ocr_done=0, ocr_total=0, msg='开始抓取…', started_at=t0)

    import concurrent.futures

    WORKERS = max(1, int(CONFIG.get('fetch_workers') or 6))

    def _download_page(pg):
        """下载单页（自带重试）；返回 (页号, 数据, 错误)"""
        params = dict(cfg['params'])
        params['page'] = pg
        params['limit'] = 100
        last = None
        for attempt in range(3):
            try:
                jj = make_session().get(BASE_URL + cfg['url'],
                                        params=params, timeout=30).json()
                return pg, (jj.get('data') or []), None
            except Exception as e:
                last = e
                time.sleep(0.4 * (attempt + 1))
        return pg, [], last

    def _consume_row(row, canon, ocr_fn):
        """处理一行：字段按 canon 顺序取值 + 时间/OCR 转换"""
        out = []
        for k in canon:
            v = row.get(k)
            if isinstance(v, str) and v.startswith('data:image'):
                n_img[0] += 1
            out.append(process_value(k, v, ocr_fn))
        return out

    n_img = [0]

    try:
        # ---- 第 1 页：探明总数、字段与页数 ----
        params = dict(cfg['params'])
        params['page'] = 1
        params['limit'] = 100
        j = session.get(BASE_URL + cfg['url'], params=params, timeout=30).json()
        if j.get('code') not in (0, '0'):
            raise RuntimeError(f'接口返回错误: {j.get("msg", j)[:120]}')
        first = j.get('data') or []
        total = int(j.get('count') or 0)
        pages_total = (total + 99) // 100 if total else 1
        if max_pages:
            pages_total = min(pages_total, max_pages)
        # 补充接口新增字段
        for k in (first[0].keys() if first else []):
            if k not in cols:
                cols.append(k)
                canon.append(k)
        _set_progress(cat, pages_total=pages_total,
                      msg=f'共 {total} 条 / {pages_total} 页，并发抓取中…')

        # ---- 按页序流式消费：并发下载，顺序处理，内存只保留少量页 ----
        def _stream():
            yield 1, first
            if pages_total > 1:
                with concurrent.futures.ThreadPoolExecutor(max_workers=WORKERS) as ex:
                    for pg, d, err in ex.map(_download_page, range(2, pages_total + 1)):
                        if err is not None:
                            print(f'[{cat}] 第 {pg} 页 3 次重试仍失败: {err}')
                        yield pg, d

        pages_seen = 0
        last_len = len(first)
        for pg, data in _stream():
            for row in data:
                rows.append(_consume_row(row, canon, ocr_fn))
                del row
            last_len = len(data)
            del data
            pages_seen += 1
            _set_progress(cat, pages_done=pages_seen, rows=len(rows),
                          ocr_total=n_img[0],
                          msg=f'已处理 {len(rows)}/{total} 条'
                              + (f'，OCR {n_img[0]} 张' if n_img[0] else ''))
            # 每 20 页落一次盘，防中断丢失
            if pages_seen % 20 == 0:
                with DATA_LOCK:
                    DATA[cat] = {'rows': rows, 'cols': cols,
                                 'fetched_at': time.time(), 'total': total}
                save_cache(cat)

        # ---- 兜底：接口 count 偏小时，继续顺序补拉 ----
        pg = pages_total + 1
        while last_len >= 100 and not max_pages:
            params = dict(cfg['params'])
            params['page'] = pg
            params['limit'] = 100
            try:
                jj = session.get(BASE_URL + cfg['url'], params=params, timeout=30).json()
                data = jj.get('data') or []
            except Exception as e:
                print(f'[{cat}] 补拉第 {pg} 页失败: {e}')
                break
            if not data:
                break
            last_len = len(data)
            for row in data:
                rows.append(_consume_row(row, canon, ocr_fn))
                del row
            total = max(total, len(rows))
            pages_seen += 1
            _set_progress(cat, pages_done=pages_seen, rows=len(rows),
                          ocr_total=n_img[0], msg=f'补充抓取，已 {len(rows)} 条')
            pg += 1

        ocr_total = n_img[0]
        with DATA_LOCK:
            DATA[cat] = {'rows': rows, 'cols': cols,
                         'fetched_at': time.time(), 'total': len(rows)}
        save_cache(cat)
        _persist_ocr_cache()

        detail_info = ''
        if cat == '简历库' and with_detail:
            detail_info = enrich_resume_details(cat)
            with DATA_LOCK:
                rows = DATA[cat]['rows']

        _set_progress(cat, phase='done', pages_done=pages_seen, rows=len(rows),
                      ocr_total=ocr_total,
                      msg=f'完成：{len(rows)} 条'
                          + (f'，OCR 识别 {ocr_total} 张电话图片' if ocr_total else '')
                          + detail_info
                          + f'，用时 {int(time.time()-t0)} 秒')
        print(f'[抓取] {cat} 完成 {len(rows)} 行')
    except Exception as e:
        traceback.print_exc()
        # 已抓到的部分也保留
        if rows:
            with DATA_LOCK:
                DATA[cat] = {'rows': rows, 'cols': cols,
                             'fetched_at': time.time(), 'total': len(rows)}
            save_cache(cat)
        _set_progress(cat, phase='error', msg=f'{e}')


# ----------------------------------------------------------------------
# 简历库：补抓「前台简历详情」全量字段
#
# 列表接口 /Post/GetPostJlList 只给列表级字段（电话还是图片，需 OCR）；
# 完整简历内容（工作经历 1-10 / 教育经历 1-5 / 语言能力 1-5 / 获得证书 1-5 /
# 职位类别 / 现居地 / 自我描述标签与详情 …）只能来自前台简历详情页
#     http://www.hbtc.ccoo.cn/post/qiuzhi/{id}x.html
# 该页**必须带前台登录 Cookie**（无 Cookie 会返回登录页）。
# ----------------------------------------------------------------------
DETAIL_COLS = list(RF.DETAIL_COLS)

# 简历详情列中默认隐藏的次要列（Excel 中可手动取消隐藏）
RESUME_DETAIL_HIDDEN = {'照片', '工作经历7', '工作经历8', '工作经历9', '工作经历10',
                        '教育经历3', '教育经历4', '教育经历5',
                        '语言能力2', '语言能力3', '语言能力4', '语言能力5',
                        '获得证书3', '获得证书4', '获得证书5'}


def enrich_resume_details(cat='简历库', workers=None, limit=None, only_missing=False):
    """给简历库补抓前台简历详情，结果作为新列追加到每行末尾。

    only_missing=True 时**跳过已经有详情的行**，只补那些还空着的
    （用于被 WAF 熔断中断后，稍后续抓剩下的部分，不必重跑全部）。

    返回一段用于展示的说明文本。
    """
    with DATA_LOCK:
        d = DATA.get(cat)
        if not d or not d.get('rows'):
            return ''
        cols = list(d['cols'])
        rows = d['rows']
        fetched_at = d.get('fetched_at') or time.time()

    if 'id' not in cols:
        return '（该类目无 id 字段，跳过简历详情）'

    fck = CONFIG.get('front_cookie') or ''
    if not RF.normalize_cookie(fck):
        return '（未配置前台 Cookie，已跳过简历详情；配置后可获取完整简历内容）'

    idx = cols.index('id')
    # 判断"是否已有详情"的锚点列：姓名（前台详情 100% 有值）
    name_idx = cols.index('姓名') if '姓名' in cols else None

    # 起始覆盖率：以「姓名」为锚点，统计已带详情的行数（含未去重前）
    already = 0
    if name_idx is not None:
        already = sum(1 for r in rows
                      if name_idx < len(r) and str(r[name_idx] or '').strip())

    jlids, seen, skipped = [], set(), 0
    for r in rows:
        rid = r[idx] if idx < len(r) else ''
        s = '' if rid in (None, '') else str(rid).strip()
        if not s or s in seen:
            continue
        seen.add(s)
        if only_missing and name_idx is not None:
            has = name_idx < len(r) and str(r[name_idx] or '').strip()
            if has:
                skipped += 1
                continue
        jlids.append(s)

    if limit is None:
        try:
            limit = int(CONFIG.get('resume_detail_limit') or 0)
        except (TypeError, ValueError):
            limit = 0
    if limit and limit > 0:
        jlids = jlids[:limit]

    if not jlids:
        if only_missing and skipped:
            return f'（全部 {skipped} 份都已有详情，无需补抓）'
        return ''

    workers = workers or max(1, min(8, int(CONFIG.get('fetch_workers') or 4)))
    try:
        rate = float(CONFIG.get('front_rate') or 1.0)
    except (TypeError, ValueError):
        rate = 1.0
    total = len(jlids)
    n_rows = len(rows)
    # 补抓期间对外显示的"覆盖"数：已有 + 本轮新增（回填模式下总数不变）
    base = already if only_missing else 0
    est = total * max(rate, 0.05)
    print(f'[简历详情] 开始抓取 {total} 份（跳过已有 {skipped} 份；'
          f'{workers} 线程，间隔 {rate}s ≈ {1.0 / max(rate, 0.01):.1f} 请求/秒，'
          f'预计 {int(est // 60)} 分 {int(est % 60)} 秒）…')
    _set_progress(cat, msg=f'正在抓取 {total} 份简历详情（前台页）…'
                           f'当前详情覆盖 {base}/{n_rows}',
                  detail_done=base, detail_total=n_rows)

    def onp(done, tot, ok, fail):
        if done % 10 == 0 or done == tot:
            _set_progress(cat, detail_done=min(n_rows, base + done),
                          detail_total=n_rows,
                          msg=f'简历详情 {done}/{tot}（成功 {ok} / 失败 {fail}），'
                              f'当前详情覆盖 {min(n_rows, base + done)}/{n_rows}')

    t0 = time.time()
    got, stat = RF.fetch_many(jlids, fck, workers=workers, on_progress=onp,
                              min_interval=rate, max_inflight=2,
                              waf_stop_after=3, retry_rounds=2, retry_delay=8.0)
    el = max(0.1, time.time() - t0)
    print(f'[简历详情] 完成 成功 {stat["ok"]} 失败 {stat["fail"]} '
          f'需登录 {stat["need_login"]} 被拦截 {stat["waf_block"]}，'
          f'{stat["rounds"]} 轮，用时 {int(el)}s（{stat["done"] / el:.2f} 请求/秒）')

    add = [c for c in DETAIL_COLS if c not in cols]
    if not add:
        # 列已存在：按 id 回填缺失单元格
        for r in rows:
            rid = r[idx] if idx < len(r) else ''
            s = '' if rid in (None, '') else str(rid).strip()
            det = got.get(s) or {}
            for c in DETAIL_COLS:
                if c not in cols:
                    continue
                ci = cols.index(c)
                if ci >= len(r):
                    continue
                if not str(r[ci] or '').strip() and det.get(c):
                    r[ci] = det[c]
    else:
        for r in rows:
            rid = r[idx] if idx < len(r) else ''
            s = '' if rid in (None, '') else str(rid).strip()
            det = got.get(s) or {}
            for c in add:
                r.append(det.get(c, ''))
        cols.extend(add)

    with DATA_LOCK:
        DATA[cat] = {'rows': rows, 'cols': cols, 'fetched_at': fetched_at,
                     'total': len(rows), 'detail_at': time.time()}
    save_cache(cat)

    # 统计最终详情覆盖率（以「姓名」为锚点）
    filled = 0
    if '姓名' in cols:
        ci = cols.index('姓名')
        filled = sum(1 for r in rows if ci < len(r) and str(r[ci] or '').strip())
    cover = f'，详情覆盖 {filled}/{len(rows)}'

    extra = ''
    if stat.get('waf_block'):
        extra = (f'（{stat["waf_block"]} 份被站点安全策略拦截，已自动熔断；'
                 f'请等十几分钟（期间不要继续请求）后点「只补缺失详情」续抓，'
                 f'或把「前台请求间隔」调大）')
    elif stat.get('need_login'):
        extra = f'（{stat["need_login"]} 份返回登录页，前台 Cookie 可能已失效）'
    if stat.get('stopped') and not stat.get('waf_block'):
        extra += '（已提前停止）'
    if skipped:
        extra += f'（跳过已有详情 {skipped} 份）'
    return (f'；简历详情 成功 {stat["ok"]} / 失败 {stat["fail"]}{extra}{cover}'
            f'，新增 {len(add)} 列')


def detail_coverage(cat='简历库'):
    """返回 (已有详情数, 总行数)，用于界面提示"""
    with DATA_LOCK:
        d = DATA.get(cat)
        if not d or not d.get('rows'):
            return 0, 0
        cols, rows = d['cols'], d['rows']
        if '姓名' not in cols:
            return 0, len(rows)
        ci = cols.index('姓名')
        return sum(1 for r in rows if ci < len(r) and str(r[ci] or '').strip()), len(rows)


def start_fetch(cat, max_pages=0, use_ocr=True, with_detail=True):
    with JOB_LOCK:
        if cat in FETCH_JOBS:
            return False, '该类目正在抓取中'
        FETCH_JOBS.add(cat)

    def runner():
        try:
            fetch_category(cat, max_pages, use_ocr, with_detail)
        finally:
            with JOB_LOCK:
                FETCH_JOBS.discard(cat)

    threading.Thread(target=runner, daemon=True).start()
    return True, '已开始'


# ----------------------------------------------------------------------
# Excel 导出
# ----------------------------------------------------------------------
EXPORT = {'phase': 'idle', 'msg': '', 'file': '', 'written': 0, 'total': 0}
EXPORT_LOCK = threading.Lock()


def _col_width(zh, en):
    base = max(len(str(zh or en)) * 2 + 2, 9)
    return min(base, 32)


def export_excel(cat, rows, cols, path):
    """write_only 模式写单个类目 xlsx：两行表头(英文/中文) + 隐藏列"""
    from openpyxl import Workbook
    hidden = set(HIDDEN_COLS.get(cat, []))
    if cat == '简历库':
        hidden |= RESUME_DETAIL_HIDDEN
    zh_map = ZH_LABELS.get(cat, {})
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(cat)
    widths = [_col_width(zh_map.get(c, c), c) for c in cols]
    for i, c in enumerate(cols):
        letter = openpyxl_get_letter(i)
        cd = ws.column_dimensions[letter]
        cd.width = widths[i]
        if c in hidden:
            cd.hidden = True
    ws.append([clean_cell(c) for c in cols])                    # 第1行：英文字段
    ws.append([clean_cell(zh_map.get(c, c)) for c in cols])     # 第2行：中文表头
    n = 0
    safe_cols = list(cols)
    for row in rows:
        # 兜底清洗：旧缓存里可能残留 XML 非法控制字符
        ws.append([clean_cell(row[i]) if i < len(row) else '' for i in range(len(safe_cols))])
        n += 1
        if n % 5000 == 0:
            with EXPORT_LOCK:
                EXPORT['written'] += 5000
    wb.save(path)
    return n


def openpyxl_get_letter(idx):
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx + 1)


def auto_filename(cats):
    tag = '+'.join(cats) if len(cats) <= 3 else f'{cats[0]}等{len(cats)}类'
    stamp = datetime.now().strftime('%Y%m%d_%H%M')
    return f'城市中国平台数据_{tag}_{stamp}'


def do_export(cats, save_dir):
    global EXPORT
    try:
        with EXPORT_LOCK:
            EXPORT = {'phase': 'running', 'msg': '准备导出…', 'file': '',
                      'written': 0, 'total': 0}
        CLEAN_STAT['cells'] = 0
        CLEAN_STAT['chars'] = 0
        if not save_dir or not os.path.isdir(save_dir):
            raise RuntimeError(f'保存目录无效: {save_dir}')
        todo = []
        total_rows = 0
        for cat in cats:
            with DATA_LOCK:
                d = DATA.get(cat)
                if d and d['rows']:
                    todo.append((cat, d['rows'], d['cols']))
                    total_rows += len(d['rows'])
            with EXPORT_LOCK:
                EXPORT['total'] = total_rows
        if not todo:
            raise RuntimeError('所选类目暂无数据，请先抓取')

        name = auto_filename([t[0] for t in todo])
        produced = []
        if len(todo) == 1:
            cat, rows, cols = todo[0]
            path = os.path.join(save_dir, name + '.xlsx')
            with EXPORT_LOCK:
                EXPORT['msg'] = f'正在写入 {os.path.basename(path)} …'
            n = export_excel(cat, rows, cols, path)
            produced.append(path)
            print(f'[导出] {cat}: {n} 行 -> {path}')
        else:
            base = os.path.join(save_dir, name)
            tmp_files = []
            for cat, rows, cols in todo:
                p = f'{base}_{cat}.xlsx'
                with EXPORT_LOCK:
                    EXPORT['msg'] = f'正在写入 {os.path.basename(p)} …'
                export_excel(cat, rows, cols, p)
                tmp_files.append(p)
                produced.append(p)
                print(f'[导出] {cat}: -> {p}')
            zip_path = base + '.zip'
            with EXPORT_LOCK:
                EXPORT['msg'] = '正在打包 zip …'
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for p in tmp_files:
                    zf.write(p, os.path.basename(p))
            produced.append(zip_path)
            # 打包完成后删除散装的 xlsx，只留 zip
            for p in tmp_files:
                try:
                    os.remove(p)
                except OSError:
                    pass
        msg = '导出完成'
        if CLEAN_STAT['cells']:
            msg += f"（已自动清洗 {CLEAN_STAT['cells']} 处非法字符/超长文本，共移除 {CLEAN_STAT['chars']} 个控制字符）"
            print(f'[导出] 清洗统计: {CLEAN_STAT}')
        with EXPORT_LOCK:
            EXPORT.update(phase='done', msg=msg, file=produced[-1],
                          files=produced, written=EXPORT.get('written', 0),
                          cleaned=CLEAN_STAT['cells'])
    except Exception as e:
        traceback.print_exc()
        with EXPORT_LOCK:
            EXPORT = {'phase': 'error', 'msg': str(e), 'file': '',
                      'written': 0, 'total': 0}


def start_export(cats, save_dir):
    with EXPORT_LOCK:
        if EXPORT.get('phase') == 'running':
            return False, '已有导出任务进行中'
    threading.Thread(target=do_export, args=(cats, save_dir), daemon=True).start()
    return True, '已开始'


# ----------------------------------------------------------------------
# Flask
# ----------------------------------------------------------------------
from flask import Flask, request, jsonify, send_from_directory, send_file, Response

app = Flask(__name__, static_folder=os.path.join(RES_DIR, 'static'),
            static_url_path='/static')
_picking = threading.Lock()


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
@app.route('/api/config', methods=['GET'])
def api_config_get():
    global CONFIG
    with CONFIG_LOCK:
        CONFIG = load_config()
    return jsonify(dict(CONFIG))


@app.route('/api/config', methods=['POST'])
def api_config_set():
    global CONFIG
    d = request.get_json(force=True, silent=True) or {}
    with CONFIG_LOCK:
        for k in ('username', 'password', 'cookie', 'front_cookie', 'save_dir'):
            if k in d:
                CONFIG[k] = str(d[k] or '')
        if 'ocr_scale' in d:
            try:
                CONFIG['ocr_scale'] = max(1, min(20, int(d['ocr_scale'])))
            except (TypeError, ValueError):
                pass
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
        if 'fetch_resume_detail' in d:
            CONFIG['fetch_resume_detail'] = 1 if d['fetch_resume_detail'] else 0
        if 'resume_detail_limit' in d:
            try:
                CONFIG['resume_detail_limit'] = max(0, int(d['resume_detail_limit']))
            except (TypeError, ValueError):
                pass
        save_config(CONFIG)
        CONFIG = load_config()
    return jsonify({'ok': True})


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


@app.route('/api/enrich_resume', methods=['POST'])
def api_enrich_resume():
    """对已抓取的简历库单独补抓前台简历详情（不重新拉列表）

    body 可选：{"only_missing": true} —— 只补还没有详情的那些（断点续抓）
    """
    d = request.get_json(force=True, silent=True) or {}
    only_missing = bool(d.get('only_missing'))
    with JOB_LOCK:
        if '简历库' in FETCH_JOBS:
            return jsonify({'ok': False, 'msg': '简历库正在抓取中'})
        FETCH_JOBS.add('简历库')

    def runner():
        try:
            info = enrich_resume_details('简历库', only_missing=only_missing)
            filled, tot = detail_coverage('简历库')
            _set_progress('简历库', phase='done',
                          msg=f'简历详情补抓完成{info}（详情覆盖 {filled}/{tot}）')
        except Exception as e:
            traceback.print_exc()
            _set_progress('简历库', phase='error', msg=f'简历详情抓取失败：{e}')
        finally:
            with JOB_LOCK:
                FETCH_JOBS.discard('简历库')

    threading.Thread(target=runner, daemon=True).start()
    return jsonify({'ok': True, 'msg': '已开始补抓简历详情'})


@app.route('/api/config/export', methods=['GET'])
def api_config_export():
    data = json.dumps(CONFIG, ensure_ascii=False, indent=2)
    return Response(data, mimetype='application/json',
                    headers={'Content-Disposition':
                             'attachment; filename=ccoocrm_config.json'})


@app.route('/api/config/import', methods=['POST'])
def api_config_import():
    global CONFIG
    d = request.get_json(force=True, silent=True)
    if not isinstance(d, dict):
        return jsonify({'ok': False, 'msg': '不是有效的 JSON 配置文件'})
    with CONFIG_LOCK:
        for k in ('username', 'password', 'cookie', 'save_dir'):
            if k in d:
                CONFIG[k] = str(d[k] or '')
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
        s = make_session('')
        s.get(BASE_URL + '/', timeout=15)
        r = s.get(VALIDATE_CODE_URL, timeout=15)
        from PIL import Image
        img = Image.open(io.BytesIO(r.content))
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        b64 = base64.b64encode(buf.getvalue()).decode()
        # 记住这个会话供登录使用
        _LOGIN_SESSIONS[threading.get_ident()] = s
        return jsonify({'ok': True, 'image': 'data:image/png;base64,' + b64})
    except Exception as e:
        return jsonify({'ok': False, 'msg': str(e)})


_LOGIN_SESSIONS = {}


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


@app.route('/api/categories', methods=['GET'])
def api_categories():
    with DATA_LOCK:
        info = {}
        for c in CATEGORIES:
            d = DATA.get(c)
            info[c] = {
                'rows': len(d['rows']) if d else 0,
                'fetched_at': (datetime.fromtimestamp(d['fetched_at'], TZ)
                               .strftime('%m-%d %H:%M:%S')) if d else '',
                'ocr': c in ('简历库', '注册用户'),
            }
    return jsonify({'cats': CATEGORIES, 'info': info,
                    'unavailable': UNAVAILABLE, 'hidden': HIDDEN_COLS,
                    'zh': {c: [ZH_LABELS.get(c, {}).get(k, k)
                               for k in HEAD_ORDER[c]] for c in CATEGORIES},
                    'en': {c: HEAD_ORDER[c] for c in CATEGORIES}})


@app.route('/api/fetch', methods=['POST'])
def api_fetch():
    d = request.get_json(force=True, silent=True) or {}
    cat = d.get('cat')
    max_pages = int(d.get('max_pages') or 0)
    use_ocr = bool(d.get('use_ocr', True))
    with_detail = bool(d.get('with_detail', True))
    if cat not in PARAMS_ORDER:
        return jsonify({'ok': False, 'msg': '未知类目'})
    ok, msg = start_fetch(cat, max_pages, use_ocr, with_detail)
    return jsonify({'ok': ok, 'msg': msg})


@app.route('/api/progress', methods=['GET'])
def api_progress():
    with JOB_LOCK:
        snap = json.dumps(PROGRESS)
    return Response(snap, mimetype='application/json')


@app.route('/api/data', methods=['GET'])
def api_data():
    cat = request.args.get('cat', '')
    page = max(1, int(request.args.get('page', 1)))
    size = min(200, max(10, int(request.args.get('size', 50))))
    kw = (request.args.get('kw') or '').strip()
    with DATA_LOCK:
        d = DATA.get(cat)
        if not d:
            return jsonify({'ok': False, 'rows': [], 'total': 0,
                            'cols': [], 'fetched_at': ''})
        rows, cols = d['rows'], d['cols']
        total_all = len(rows)
        if kw:
            idxs = [i for i, c in enumerate(cols)
                    if c not in HIDDEN_COLS.get(cat, [])]
            def hit(r):
                for i in idxs:
                    if i < len(r) and kw in str(r[i]):
                        return True
                return False
            rows = [r for r in rows if hit(r)]
        total = len(rows)
        start = (page - 1) * size
        page_rows = rows[start:start + size]
        return jsonify({'ok': True, 'cols': cols, 'rows': page_rows,
                        'total': total, 'total_all': total_all,
                        'fetched_at': datetime.fromtimestamp(
                            d['fetched_at'], TZ).strftime('%Y-%m-%d %H:%M:%S')})


@app.route('/api/pickdir', methods=['POST'])
def api_pickdir():
    if not _picking.acquire(blocking=False):
        return jsonify({'ok': False, 'msg': '已有目录选择窗口打开'})
    path = ''
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        path = filedialog.askdirectory(title='选择 Excel 保存目录') or ''
        root.destroy()
    except Exception as e:
        _picking.release()
        return jsonify({'ok': False, 'msg': f'无法打开目录选择器: {e}'})
    _picking.release()
    if path:
        with CONFIG_LOCK:
            CONFIG['save_dir'] = path
            save_config(CONFIG)
    return jsonify({'ok': True, 'path': path})


@app.route('/api/export', methods=['POST'])
def api_export():
    d = request.get_json(force=True, silent=True) or {}
    cats = d.get('cats') or []
    cats = [c for c in cats if c in PARAMS_ORDER]
    save_dir = str(d.get('dir') or CONFIG.get('save_dir') or '')
    if not cats:
        return jsonify({'ok': False, 'msg': '请至少勾选一个类目'})
    if not save_dir:
        return jsonify({'ok': False, 'msg': '请先选择保存目录'})
    ok, msg = start_export(cats, save_dir)
    return jsonify({'ok': ok, 'msg': msg})


@app.route('/api/export_progress', methods=['GET'])
def api_export_progress():
    with EXPORT_LOCK:
        snap = json.dumps(EXPORT)
    return Response(snap, mimetype='application/json')


@app.route('/api/open_folder', methods=['POST'])
def api_open_folder():
    d = request.get_json(force=True, silent=True) or {}
    path = str(d.get('path') or '')
    if not path or not os.path.exists(path):
        return jsonify({'ok': False, 'msg': '路径不存在'})
    try:
        if os.path.isfile(path):
            os.startfile(os.path.dirname(path))   # noqa
        else:
            os.startfile(path)                    # noqa
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'msg': str(e)})


# ----------------------------------------------------------------------
def main():
    load_all_caches()
    port = 8787
    for p in (8787, 8788, 8790, 8791):
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
    print('=' * 60)
    print('城市中国·湖北通城网 招聘频道 平台数据面板')
    print(f'请用浏览器访问: {url}  （将自动打开）')
    print(f'配置/缓存目录: {DATA_DIR}')
    print('=' * 60)
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host='127.0.0.1', port=port, threaded=True, debug=False)


if __name__ == '__main__':
    main()
