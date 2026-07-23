import subprocess, os
script_dir = 'C:/Users/wy98k/Documents/片区分类开发工具。'
python_exe = 'C:/Users/wy98k/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
os.chdir(script_dir)
env = os.environ.copy()
env["GAODE_KEY"] = env.get("GAODE_KEY") or "617e58986e6e6d1872bb9464924692b0"
# Use DETACHED_PROCESS flag to prevent parent from killing child
proc = subprocess.Popen(
    [python_exe, os.path.join(script_dir, 'app.py')],
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
    creationflags=subprocess.DETACHED_PROCESS,
    env=env,
    close_fds=True
)
print(f'Started PID: {proc.pid}')
