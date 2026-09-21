import os
import sys
import json
import shutil
import subprocess
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# -------------------- 配置记忆 --------------------
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
if getattr(sys, 'frozen', False):
    CONFIG_FILE = os.path.join(os.path.dirname(sys.executable), 'config.json')

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {'work_dir': '', 'port': 5000}

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f)
    except:
        pass

# -------------------- GUI 主窗口 --------------------
class LauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title('网站巡检系统')
        self.root.resizable(False, False)
        self.root.protocol('WM_DELETE_WINDOW', self.on_close)

        self.process = None
        self.config = load_config()

        # 居中显示
        self.root.update_idletasks()
        w, h = 420, 280
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f'{w}x{h}+{x}+{y}')

        self._build_ui()

    def _build_ui(self):
        pad = {'padx': 12, 'pady': 6}

        # 标题
        tk.Label(self.root, text='网站巡检系统', font=('Microsoft YaHei', 16, 'bold')).pack(pady=(20, 10))

        # 端口
        port_frame = tk.Frame(self.root)
        port_frame.pack(fill='x', **pad)
        tk.Label(port_frame, text='端口:', width=8, anchor='e').pack(side='left')
        self.port_var = tk.StringVar(value=str(self.config.get('port', 5000)))
        tk.Entry(port_frame, textvariable=self.port_var, width=10).pack(side='left', padx=(4, 0))

        # 工作目录
        dir_frame = tk.Frame(self.root)
        dir_frame.pack(fill='x', **pad)
        tk.Label(dir_frame, text='工作目录:', width=8, anchor='e').pack(side='left')
        self.dir_var = tk.StringVar(value=self.config.get('work_dir', ''))
        tk.Entry(dir_frame, textvariable=self.dir_var, width=28).pack(side='left', padx=(4, 4))
        tk.Button(dir_frame, text='选择', command=self.choose_dir).pack(side='left')

        # 按钮
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=16)
        self.start_btn = tk.Button(btn_frame, text='启动服务', width=12, command=self.start_service,
                                   bg='#409eff', fg='white', font=('Microsoft YaHei', 10))
        self.start_btn.pack(side='left', padx=8)
        self.stop_btn = tk.Button(btn_frame, text='停止服务', width=12, command=self.stop_service,
                                  bg='#f56c6c', fg='white', font=('Microsoft YaHei', 10), state='disabled')
        self.stop_btn.pack(side='left', padx=8)

        # 状态
        self.status_var = tk.StringVar(value='状态: 已停止')
        tk.Label(self.root, textvariable=self.status_var, font=('Microsoft YaHei', 10), fg='#909399').pack()

        self.url_var = tk.StringVar(value='')
        self.url_label = tk.Label(self.root, textvariable=self.url_var, font=('Microsoft YaHei', 9),
                                  fg='#409eff', cursor='hand2')
        self.url_label.pack()
        self.url_label.bind('<Button-1>', lambda e: self.open_browser())

        self.browser_btn = tk.Button(self.root, text='打开浏览器', command=self.open_browser, state='disabled')
        self.browser_btn.pack(pady=(4, 12))

    def choose_dir(self):
        d = filedialog.askdirectory(title='选择工作目录')
        if d:
            self.dir_var.set(d)

    def start_service(self):
        work_dir = self.dir_var.get().strip()
        if not work_dir:
            messagebox.showwarning('提示', '请先选择工作目录')
            return
        try:
            port = int(self.port_var.get().strip())
        except ValueError:
            messagebox.showwarning('提示', '端口必须是数字')
            return

        os.makedirs(work_dir, exist_ok=True)

        # 保存配置
        self.config['work_dir'] = work_dir
        self.config['port'] = port
        save_config(self.config)

        # 设置环境变量
        env = os.environ.copy()
        env['WORK_DIR'] = work_dir
        env['FLASK_PORT'] = str(port)
        env['FLASK_DEBUG'] = '0'  # GUI 模式下关闭 debug，避免 reloader 双进程

        # 启动 Flask 子进程
        if getattr(sys, 'frozen', False):
            # PyInstaller 打包后，用自身 exe 运行 app 模块
            self.process = subprocess.Popen(
                [sys.executable, '--flask'],
                env=env,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        else:
            # 开发模式
            app_py = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.py')
            self.process = subprocess.Popen(
                [sys.executable, app_py],
                env=env,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )

        # 更新 UI
        self.start_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.browser_btn.config(state='normal')
        self.status_var.set('状态: 服务运行中')
        self.url_var.set(f'http://localhost:{port}')

        # 延迟打开浏览器
        self.root.after(2000, self.open_browser)

    def stop_service(self):
        if not self.process:
            return

        # 先调用停止巡检接口（如果有正在运行的巡检）
        port = self.port_var.get().strip()
        try:
            import urllib.request
            urllib.request.urlopen(
                f'http://localhost:{port}/api/inspect/stop',
                data=b'{}',
                timeout=3
            )
        except:
            pass

        # 等待 2 秒让巡检停止
        import time
        time.sleep(2)

        # 终止 Flask 进程
        try:
            self.process.terminate()
            self.process.wait(timeout=5)
        except:
            try:
                self.process.kill()
                self.process.wait(timeout=3)
            except:
                pass
        self.process = None

        # 清理工作目录
        work_dir = self.dir_var.get().strip()
        if work_dir and os.path.exists(work_dir):
            try:
                for item in os.listdir(work_dir):
                    item_path = os.path.join(work_dir, item)
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path, ignore_errors=True)
                    else:
                        try:
                            os.remove(item_path)
                        except:
                            pass
            except:
                pass

        # 更新 UI
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.browser_btn.config(state='disabled')
        self.status_var.set('状态: 已停止')
        self.url_var.set('')

    def open_browser(self):
        port = self.port_var.get().strip()
        if port:
            webbrowser.open(f'http://localhost:{port}')

    def on_close(self):
        if self.process:
            if messagebox.askokcancel('确认', '服务正在运行，关闭窗口将停止服务并清理数据。确定关闭吗？'):
                self.stop_service()
                self.root.destroy()
            else:
                return
        else:
            self.root.destroy()


# -------------------- PyInstaller 打包后入口 --------------------
def run_flask():
    """PyInstaller 打包后，通过 --flask 参数启动 Flask"""
    from app import app
    port = int(os.environ.get('FLASK_PORT', '5000'))
    app.run(host='0.0.0.0', port=port, debug=False)


if __name__ == '__main__':
    if '--flask' in sys.argv:
        # 子进程模式：运行 Flask
        run_flask()
    else:
        # 主进程：显示 GUI
        root = tk.Tk()
        LauncherApp(root)
        root.mainloop()
