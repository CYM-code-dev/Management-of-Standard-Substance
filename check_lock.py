# -*- coding: utf-8 -*-
"""检查目标 Excel 是否被 Excel 占用（锁文件），并测试本机写权限。"""
import os, sys, tempfile
sys.stdout.reconfigure(encoding='utf-8')

d = r'\\files.cirs-ck.com\轻工公盘\有机\3-有机耗品'
target = os.path.join(d, '有机标准品-20250910.xlsx')
lock = os.path.join(d, '~$有机标准品-20250910.xlsx')

print('目录可访问:', os.path.exists(d))
print('目标文件存在:', os.path.exists(target))
print('Excel 锁文件(~$)存在:', os.path.exists(lock), '→ 存在=有人正用 Excel 打开')

# 列出相关文件
try:
    rel = [n for n in os.listdir(d) if '有机标准品' in n or n.startswith('~$')]
    print('相关条目:', rel)
except Exception as e:
    print('listdir 失败:', e)

# 本机写权限测试：在该目录创建一个临时文件
try:
    tp = os.path.join(d, '__wtest_' + tempfile._get_default_tempdir()[-4:] + '.tmp')
    with open(tp, 'w') as f:
        f.write('x')
    os.remove(tp)
    print('本机写测试: 成功（你的账号可写该目录）')
except Exception as e:
    print('本机写测试: 失败 →', repr(e))
