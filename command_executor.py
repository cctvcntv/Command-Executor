import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import threading
import json
import os
import sys
import time

# 数据文件路径
DATA_FILE = os.path.join(os.path.dirname(sys.argv[0]), "data.json")

class CommandApp:
    def __init__(self, root):
        self.root = root
        self.root.title("命令执行器 开发：一兵，88389917@qq.com")
        self.root.geometry("900x700")
        self.center_window()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 数据存储: list of dict, 包含 description, note, commands, status
        self.data = []
        self.filtered_indices = []   # 过滤后的索引(相对于self.data)
        self.load_data()

        # 当前执行任务相关
        self.current_task_index = None   # 正在执行的任务索引
        self.current_process = None      # 当前子进程
        self.stop_flag = False            # 停止当前任务的标志

        # 创建界面
        self.create_widgets()
        self.refresh_treeview()
        self.rebuild_task_tabs()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        # 顶部查询区域
        query_frame = tk.Frame(self.root)
        query_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(query_frame, text="查询:").pack(side=tk.LEFT)
        self.query_var = tk.StringVar()
        self.query_entry = tk.Entry(query_frame, textvariable=self.query_var, width=30)
        self.query_entry.pack(side=tk.LEFT, padx=5)
        self.query_btn = tk.Button(query_frame, text="查询", command=self.query)
        self.query_btn.pack(side=tk.LEFT)
        self.clear_query_btn = tk.Button(query_frame, text="清除", command=self.clear_query)
        self.clear_query_btn.pack(side=tk.LEFT, padx=5)

        # 中间主区域：左侧任务表格 + 右侧按钮
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 左侧表格（带滚动条）
        tree_frame = tk.Frame(main_frame)
        tree_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        columns = ('name', 'status', 'note')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='browse')
        self.tree.heading('name', text='任务名称')
        self.tree.heading('status', text='运行情况')
        self.tree.heading('note', text='说明')
        self.tree.column('name', width=200, anchor='w')
        self.tree.column('status', width=100, anchor='center')
        self.tree.column('note', width=200, anchor='w')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # 右侧按钮区域
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))

        self.add_btn = tk.Button(btn_frame, text="添加", width=12, command=self.add_command)
        self.add_btn.pack(pady=5)

        self.edit_btn = tk.Button(btn_frame, text="修改", width=12, command=self.edit_command, state=tk.DISABLED)
        self.edit_btn.pack(pady=5)

        self.delete_btn = tk.Button(btn_frame, text="删除", width=12, command=self.delete_command, state=tk.DISABLED)
        self.delete_btn.pack(pady=5)

        self.execute_btn = tk.Button(btn_frame, text="执行", width=12, command=self.start_execution, state=tk.DISABLED)
        self.execute_btn.pack(pady=5)

        self.stop_btn = tk.Button(btn_frame, text="停止", width=12, command=self.stop_execution, state=tk.DISABLED)
        self.stop_btn.pack(pady=5)

        # 底部Tab控件（执行输出）
        nb_frame = tk.Frame(self.root)
        nb_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tk.Label(nb_frame, text="执行情况:").pack(anchor=tk.W)
        self.notebook = ttk.Notebook(nb_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ---------- 数据操作 ----------
    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    raw = json.load(f)
                self.data = []
                for item in raw:
                    new_item = {
                        "description": item.get("description", ""),
                        "commands": item.get("commands", []),
                        "note": item.get("note", ""),
                        "status": "未执行"
                    }
                    self.data.append(new_item)
            except Exception as e:
                messagebox.showerror("错误", f"加载数据失败: {e}")
                self.data = []
        else:
            self.data = []

    def save_data(self):
        to_save = []
        for item in self.data:
            to_save.append({
                "description": item["description"],
                "commands": item["commands"],
                "note": item.get("note", "")
            })
        try:
            with open(DATA_FILE, 'w', encoding='utf-8') as f:
                json.dump(to_save, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("错误", f"保存数据失败: {e}")

    # ---------- 界面刷新 ----------
    def refresh_treeview(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        if not self.filtered_indices:
            indices = range(len(self.data))
        else:
            indices = self.filtered_indices

        for i in indices:
            item = self.data[i]
            name = item.get('description', '').strip() or "(无名称)"
            status = item.get('status', '未执行')
            note = item.get('note', '')
            self.tree.insert('', tk.END, iid=str(i), values=(name, status, note))

        if not indices:
            self.tree.insert('', tk.END, values=("(无任务)", "", ""))

        self.tree.selection_remove(*self.tree.selection())
        self.update_buttons_for_selection()

    def rebuild_task_tabs(self):
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)

        self.task_tabs = {}
        if not self.data:
            frame = ttk.Frame(self.notebook)
            text = scrolledtext.ScrolledText(frame, wrap=tk.WORD)
            text.pack(fill=tk.BOTH, expand=True)
            text.insert(tk.END, "没有任务，请先添加。")
            text.config(state=tk.DISABLED)
            self.notebook.add(frame, text="空")
            return

        for idx, item in enumerate(self.data):
            frame = ttk.Frame(self.notebook)
            text = scrolledtext.ScrolledText(frame, wrap=tk.WORD)
            text.pack(fill=tk.BOTH, expand=True)
            text.config(state=tk.NORMAL)
            text.delete(1.0, tk.END)
            text.insert(tk.END, "尚未执行。\n")
            text.config(state=tk.DISABLED)
            tab_title = f"任务{idx+1}: {item.get('description','')[:20]}"
            self.notebook.add(frame, text=tab_title)
            self.task_tabs[idx] = text

    def clear_task_tab(self, idx):
        if idx in self.task_tabs:
            text = self.task_tabs[idx]
            text.config(state=tk.NORMAL)
            text.delete(1.0, tk.END)
            text.config(state=tk.DISABLED)

    def append_output_to_task(self, idx, text_content):
        def do_append():
            if idx in self.task_tabs:
                text_widget = self.task_tabs[idx]
                text_widget.config(state=tk.NORMAL)
                text_widget.insert(tk.END, text_content)
                text_widget.see(tk.END)
                text_widget.config(state=tk.DISABLED)
        self.root.after(0, do_append)

    def update_task_status(self, idx, status):
        if 0 <= idx < len(self.data):
            self.data[idx]['status'] = status
            item_id = str(idx)
            if self.tree.exists(item_id):
                self.tree.set(item_id, 'status', status)

    # ---------- 按钮状态管理 ----------
    def on_select(self, event):
        self.update_buttons_for_selection()

    def update_buttons_for_selection(self):
        selected = self.tree.selection()
        if not selected:
            # 无选中项：只有添加可用
            self.add_btn.config(state=tk.NORMAL)
            self.edit_btn.config(state=tk.DISABLED)
            self.delete_btn.config(state=tk.DISABLED)
            self.execute_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.DISABLED)
            return

        idx = int(selected[0])
        task_status = self.data[idx]['status']

        if self.current_task_index is not None:
            # 有任务在执行
            if idx == self.current_task_index:
                # 选中正在执行的任务
                self.add_btn.config(state=tk.NORMAL)
                self.edit_btn.config(state=tk.DISABLED)
                self.delete_btn.config(state=tk.DISABLED)
                self.execute_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.NORMAL)
            else:
                # 选中其他任务
                self.add_btn.config(state=tk.NORMAL)
                # 修改：只要不是运行中（实际不会是运行中）即可修改
                can_edit = task_status != "运行中"
                self.edit_btn.config(state=tk.NORMAL if can_edit else tk.DISABLED)
                # 删除：为避免索引变化，始终禁用
                self.delete_btn.config(state=tk.DISABLED)
                # 执行：始终禁用（不能同时执行两个任务）
                self.execute_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.DISABLED)
        else:
            # 没有任务执行
            self.add_btn.config(state=tk.NORMAL)
            can_operate = task_status != "运行中"
            self.edit_btn.config(state=tk.NORMAL if can_operate else tk.DISABLED)
            self.delete_btn.config(state=tk.NORMAL if can_operate else tk.DISABLED)
            self.execute_btn.config(state=tk.NORMAL if can_operate else tk.DISABLED)
            self.stop_btn.config(state=tk.DISABLED)

    # ---------- 查询功能 ----------
    def query(self):
        query_text = self.query_var.get().strip().lower()
        if not query_text:
            self.filtered_indices = []
        else:
            filtered = []
            for i, item in enumerate(self.data):
                if query_text in item.get('description', '').lower():
                    filtered.append(i)
                    continue
                if query_text in item.get('note', '').lower():
                    filtered.append(i)
                    continue
                for cmd in item.get('commands', []):
                    if query_text in cmd.lower():
                        filtered.append(i)
                        break
            self.filtered_indices = filtered
        self.refresh_treeview()

    def clear_query(self):
        self.query_var.set("")
        self.filtered_indices = []
        self.refresh_treeview()

    def get_selected_index(self):
        selected = self.tree.selection()
        if not selected:
            return None
        try:
            return int(selected[0])
        except ValueError:
            return None

    # ---------- 增删改 ----------
    def add_command(self):
        dialog = CommandDialog(self.root, "添加任务")
        if dialog.result is not None:
            desc, cmd_list, note = dialog.result
            new_item = {
                "description": desc,
                "commands": cmd_list,
                "note": note,
                "status": "未执行"
            }
            self.data.append(new_item)
            self.save_data()
            self.clear_query()
            self.rebuild_task_tabs()

    def edit_command(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        if self.current_task_index is not None and idx == self.current_task_index:
            messagebox.showwarning("警告", "任务正在执行，无法修改")
            return
        item = self.data[idx]
        dialog = CommandDialog(self.root, "修改任务",
                               description=item['description'],
                               commands=item['commands'],
                               note=item.get('note', ''))
        if dialog.result is not None:
            desc, cmd_list, note = dialog.result
            self.data[idx]['description'] = desc
            self.data[idx]['commands'] = cmd_list
            self.data[idx]['note'] = note
            self.data[idx]['status'] = "未执行"   # 修改后重置状态
            self.save_data()
            self.refresh_treeview()
            self.rebuild_task_tabs()

    def delete_command(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        if self.current_task_index is not None:
            messagebox.showwarning("警告", "有任务正在执行，无法删除")
            return
        if messagebox.askyesno("确认", "确定要删除该任务吗？"):
            del self.data[idx]
            self.save_data()
            self.clear_query()
            self.rebuild_task_tabs()

    # ---------- 执行与停止 ----------
    def start_execution(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        if self.current_task_index is not None:
            messagebox.showinfo("提示", "已有任务正在执行，请等待完成或停止")
            return
        if self.data[idx]['status'] == "运行中":
            return

        self.current_task_index = idx
        self.stop_flag = False
        self.update_task_status(idx, "运行中")
        self.clear_task_tab(idx)
        self.append_output_to_task(idx, f"=== 开始执行任务: {self.data[idx]['description']} ===\n")
        self.update_buttons_for_selection()

        thread = threading.Thread(target=self.run_task, args=(idx,), daemon=True)
        thread.start()

    def stop_execution(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        if self.current_task_index != idx:
            return

        # 如果进程对象已不存在，说明进程已经结束或被清理
        if self.current_process is None:
            self.append_output_to_task(idx, "\n--- 用户请求停止 (进程已结束) ---\n")
            self.stop_flag = True
            self.clear_current_task()
            return

        if self.current_process.poll() is None:
            self.stop_flag = True
            self.append_output_to_task(idx, "\n--- 用户请求停止 ---\n")
            pid = self.current_process.pid
            try:
                self.current_process.terminate()
                time.sleep(0.3)

                if self.current_process.poll() is None:
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    subprocess.run(
                        ['taskkill', '/F', '/T', '/PID', str(pid)],
                        capture_output=True,
                        startupinfo=startupinfo,
                        text=True
                    )
                    self.append_output_to_task(idx, "进程树已强制终止\n")
            except Exception as e:
                self.append_output_to_task(idx, f"终止进程时出错: {e}\n")
            finally:
                self.current_process = None
        else:
            self.append_output_to_task(idx, "\n--- 用户请求停止 (进程已自然结束) ---\n")
            self.current_process = None

        self.update_buttons_for_selection()

    def run_task(self, idx):
        try:
            item = self.data[idx]
            cmds = item['commands']
            success = True

            for cmd_num, cmd in enumerate(cmds, 1):
                if self.stop_flag:
                    self.append_output_to_task(idx, "任务已被用户停止\n")
                    success = False
                    break

                self.append_output_to_task(idx, f"> 执行命令 {cmd_num}: {cmd}\n")

                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

                try:
                    self.current_process = subprocess.Popen(
                        cmd,
                        shell=True,
                        startupinfo=startupinfo,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        stdin=subprocess.PIPE,
                        text=True,
                        encoding='gbk',
                        bufsize=1
                    )

                    def read_stream(stream, prefix=""):
                        for line in iter(stream.readline, ''):
                            if self.stop_flag:
                                break
                            self.append_output_to_task(idx, f"{prefix}{line}")
                        stream.close()

                    t_out = threading.Thread(target=read_stream, args=(self.current_process.stdout, ""))
                    t_err = threading.Thread(target=read_stream, args=(self.current_process.stderr, "错误: "))
                    t_out.daemon = True
                    t_err.daemon = True
                    t_out.start()
                    t_err.start()

                    self.current_process.wait()
                    t_out.join(timeout=1)
                    t_err.join(timeout=1)

                    if self.current_process.returncode != 0 and not self.stop_flag:
                        self.append_output_to_task(idx, f"命令返回非零值: {self.current_process.returncode}\n")
                        success = False

                except Exception as e:
                    self.append_output_to_task(idx, f"执行异常: {str(e)}\n")
                    success = False
                    break
                finally:
                    self.current_process = None

                self.append_output_to_task(idx, "\n")

            if self.stop_flag:
                final_status = "已停止"
            elif success:
                final_status = "已完成"
            else:
                final_status = "出错"

        except Exception as e:
            self.append_output_to_task(idx, f"任务执行过程中发生严重错误: {e}\n")
            final_status = "出错"
        finally:
            self.root.after(0, lambda: self.update_task_status(idx, final_status))
            self.root.after(0, self.clear_current_task)

    def clear_current_task(self):
        self.current_task_index = None
        self.current_process = None
        self.stop_flag = False
        self.update_buttons_for_selection()
        self.status_var.set("就绪")

    def on_closing(self):
        if self.current_task_index is not None:
            self.stop_execution()
            time.sleep(0.5)
        self.save_data()
        self.root.destroy()


class CommandDialog:
    """添加/修改任务的对话框"""
    def __init__(self, parent, title, description="", commands=None, note=""):
        self.result = None
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(False, False)
        self.center_dialog(parent)

        tk.Label(self.dialog, text="任务名称:").pack(anchor=tk.W, padx=10, pady=(10, 0))
        self.name_var = tk.StringVar(value=description)
        self.name_entry = tk.Entry(self.dialog, textvariable=self.name_var, width=50)
        self.name_entry.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(self.dialog, text="说明:").pack(anchor=tk.W, padx=10, pady=(10, 0))
        self.note_var = tk.StringVar(value=note)
        self.note_entry = tk.Entry(self.dialog, textvariable=self.note_var, width=50)
        self.note_entry.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(self.dialog, text="命令行 (每行一条命令):").pack(anchor=tk.W, padx=10, pady=(10, 0))
        self.cmd_text = tk.Text(self.dialog, height=12, width=60)
        self.cmd_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        if commands:
            self.cmd_text.insert(1.0, "\n".join(commands))

        btn_frame = tk.Frame(self.dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Button(btn_frame, text="确定", width=10, command=self.ok).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="取消", width=10, command=self.cancel).pack(side=tk.RIGHT)

        self.dialog.bind('<Return>', lambda e: self.ok())
        self.dialog.bind('<Escape>', lambda e: self.cancel())

        parent.wait_window(self.dialog)

    def center_dialog(self, parent):
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def ok(self):
        name = self.name_var.get().strip()
        note = self.note_var.get().strip()
        cmd_text = self.cmd_text.get(1.0, tk.END).strip()
        if not cmd_text:
            messagebox.showwarning("警告", "请输入至少一条命令", parent=self.dialog)
            return

        cmd_list = [line.strip() for line in cmd_text.splitlines() if line.strip()]
        if not cmd_list:
            messagebox.showwarning("警告", "命令行不能为空", parent=self.dialog)
            return

        self.result = (name, cmd_list, note)
        self.dialog.destroy()

    def cancel(self):
        self.dialog.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = CommandApp(root)
    root.mainloop()