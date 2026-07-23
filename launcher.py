import subprocess, sys, os
script_dir = 'C:/Users/wy98k/Documents/片区分类开发工具。'
python_exe = 'C:/Users/wy98k/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
os.chdir(script_dir)
env = os.environ.copy()
env["GAODE_KEY"] = env.get("GAODE_KEY") or "617e58986e6e6d1872bb9464924692b0"
proc = subprocess.Popen(
    [python_exe, 'app.py'],
    stdout=open(os.path.join(script_dir, 'server_out.log'), 'w'),
    stderr=open(os.path.join(script_dir, 'server_err.log'), 'w'),
    creationflags=subprocess.CREATE_NO_WINDOW,
    env=env
)
print(f'Started PID: {proc.pid}')
