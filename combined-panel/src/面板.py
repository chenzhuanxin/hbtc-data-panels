#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
城市中国·湖北通城网 招聘频道 综合数据面板
==========================================
一个 exe / 一个端口，集成两个工具（门户 + 两个子面板）：

  /      综合面板首页（门户导航）
  /city  平台数据面板 —— 7 类目全量抓取 + 电话图片 OCR + 简历/岗位/企业/用户全量字段补抓
  /jd    企业投递简历采集面板 —— 企业 → 岗位 → 投递简历 → 完整简历内容

运行：python 面板.py    （自动打开 http://127.0.0.1:8795）

  * 配置共享：根目录 config.json —— 后台 Cookie / 总后台 Cookie / 保存目录只需填一次
  * 数据分存：数据_平台面板/ 与 数据_简历面板/ 互不干扰
  * 详情来源：简历 JlEdit / 岗位 JobEdit / 企业 CompanyAdd（job.admin.ccoo.cn，后台 Cookie）
              用户 user_set.asp（www.admin.ccoo.cn，总后台 Cookie，纯文本无需 OCR）
"""

import os
import sys
import threading
import webbrowser

sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 源码模式：保证 panels 包与 admin_detail 可导入
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from flask import Flask, send_from_directory                    # noqa: E402
from werkzeug.middleware.dispatcher import DispatcherMiddleware  # noqa: E402
from werkzeug.serving import run_simple                         # noqa: E402

# 两个子面板（静态导入，PyInstaller 依赖分析可追踪）
from panels.city import panel as city_panel                     # noqa: E402
from panels.jd import panel as jd_panel                         # noqa: E402

PORT = 8795


def _root():
    """可写根目录：源码模式 = 本项目目录；exe 模式 = exe 所在目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return BASE_DIR


# ----------------------------------------------------------------------
# 门户（/）
# ----------------------------------------------------------------------
app = Flask(__name__, static_folder=os.path.join(BASE_DIR, 'static'),
            static_url_path='/static')


@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')


# 一个进程挂载两个子面板：/city/** 与 /jd/**
# DispatcherMiddleware 会把挂载前缀写进 SCRIPT_NAME，
# 子面板内部仍以 / 与 /api/... 的视角工作，互不冲突。
application = DispatcherMiddleware(app, {
    '/city': city_panel.app,
    '/jd': jd_panel.app,
})


def main():
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
    print('=' * 66)
    print('  城市中国·湖北通城网 招聘频道 综合数据面板')
    print(f'  首页     :  {url}/')
    print(f'  平台数据 :  {url}/city/    （7 类目 + OCR + 简历全量字段）')
    print(f'  企业投递 :  {url}/jd/      （企业 → 岗位 → 投递 → 完整简历）')
    print(f'  共享配置 :  {os.path.join(_root(), "config.json")}')
    print(f'  数据目录 :  {os.path.join(_root(), "数据_平台面板")} / '
          f'{os.path.join(_root(), "数据_简历面板")}')
    print('  按 Ctrl+C 退出')
    print('=' * 66)
    try:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    except Exception:
        pass
    run_simple('127.0.0.1', port, application, threaded=True)


if __name__ == '__main__':
    main()
