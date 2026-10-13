# -*- coding: utf-8 -*-
"""
一键打包 exe：企业投递简历采集面板
用法：python 打包exe.py    （需要已 pip install pyinstaller）
产物：dist/企业投递简历采集面板.exe
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)

NAME = '企业投递简历采集面板'

cmd = [
    sys.executable, '-X', 'utf8', '-m', 'PyInstaller',
    '--noconfirm', '--clean', '--onefile', '--console',
    '--name', NAME,
    '--icon', 'app.ico',
    '--add-data', 'static;static',
    # RapidOCR + onnxruntime + 模型文件（个别电话字段是 base64 图片时启用）
    '--collect-all', 'rapidocr_onnxruntime',
    '面板.py',
]

print('执行：')
print('  ' + ' '.join(cmd))
r = subprocess.run(cmd)
if r.returncode == 0:
    out = os.path.join(HERE, 'dist', NAME + '.exe')
    size = os.path.getsize(out) / 1024 / 1024
    print(f'\n[完成] {out}  ({size:.1f} MB)')
else:
    print('\n[失败] 请查看 build/ 目录下的日志或上方输出')
sys.exit(r.returncode)
