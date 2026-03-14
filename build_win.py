"""
Windows build helper: forces UTF-8 then calls flet pack via CLI entry point.
"""
import sys
import io
import os
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

env = os.environ.copy()
env['PYTHONUTF8'] = '1'
env['PYTHONIOENCODING'] = 'utf-8'

# flet installs a 'flet' console script - find and call it directly
import shutil
flet_bin = shutil.which('flet')
if not flet_bin:
    print("ERROR: flet command not found")
    sys.exit(1)

cmd = [
    sys.executable, flet_bin,
    'pack', 'launcher.py',
    '--name', 'VoucherTool',
    '--add-data', 'app.py;.',
    '--add-data', 'company_manager.py;.',
    '--add-data', 'rules_manager.py;.',
    '--add-data', 'processor;processor',
    '--add-data', 'utils;utils',
]

result = subprocess.run(cmd, env=env)
sys.exit(result.returncode)
