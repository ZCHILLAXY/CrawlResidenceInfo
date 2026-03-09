"""
GUI
Made with ❤️by Z🐻
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import logging
from pathlib import Path
from datetime import datetime
import sys
import os

from crawler import ResidencePointsCrawler
from excel_handler import ExcelHandler
from config import Config

class TextHandler(logging.Handler):

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.after(0, self._append_text, msg)

    def _append_text(self, msg):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

class ResidencePointGUI:

    def __init__(self, root):
        self.root = root
        self.root.title("居住证查询工具 v1.0 - Made with ❤️by Z🐻")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        self.config = Config()

        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.session_id = tk.StringVar(value=self.config.get('session_id', ''))
        self.max_retries = tk.IntVar(value=self.config.get('max_retries', 20))

        self.is_running = False
        self.crawler = None
        self.excel_handler = ExcelHandler()

        self._setup_logging()

        self._create_widgets()

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _setup_logging(self):
        logger = logging.getLogger()
        logger.handlers.clear()

        logger.setLevel(logging.INFO)

        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%H:%M:%S'
        )

        log_dir = Path('logs')
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / f'query_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'

        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

        self.logger = logger

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)

        self._create_file_section(main_frame)

        self._create_config_section(main_frame)

        self._create_control_section(main_frame)

        self._create_log_section(main_frame)

        self._create_status_bar(main_frame)

    def _create_file_section(self, parent):
        file_frame = ttk.LabelFrame(parent, text="文件设置", padding="10")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="输入文件：").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.input_file_path, state='readonly').grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        ttk.Button(file_frame, text="选择Excel", command=self._select_input_file).grid(
            row=0, column=2, pady=5
        )

        ttk.Label(file_frame, text="输出文件：").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_file_path, state='readonly').grid(
            row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        ttk.Button(file_frame, text="选择保存位置", command=self._select_output_file).grid(
            row=1, column=2, pady=5
        )

    def _create_config_section(self, parent):
        config_frame = ttk.LabelFrame(parent, text="高级设置", padding="10")
        config_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="Session ID: ").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(config_frame, textvariable=self.session_id).grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        ttk.Label(config_frame, text="（如果Cookie失效可以留空）", font=('', 8)).grid(
            row=0, column=2, sticky=tk.W, pady=5
        )
        ttk.Label(config_frame, text="最大重试次数: ").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(
            config_frame,
            from_=5,
            to=50,
            textvariable=self.max_retries,
            width=10
        ).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

    def _create_control_section(self, parent):
        control_frame = ttk.Frame(parent)
        control_frame.grid(row=2, column=0, pady=(0, 10))

        self.start_button = ttk.Button(
            control_frame,
            text="开始查询",
            command=self._start_query,
            width=15
        )
        self.start_button.grid(row=0, column=0, padx=5)

        self.stop_button = ttk.Button(
            control_frame,
            text="停止查询",
            command=self._stop_query,
            width=15,
            state='disabled'
        )
        self.stop_button.grid(row=0, column=1, padx=5)

        ttk.Button(
            control_frame,
            text='清空日志',
            command=self._clear_log,
            width=15
        ).grid(row=0, column=2, padx=5)

    def _create_log_section(self, parent):
        log_frame = ttk.LabelFrame(parent, text="运行日志", padding="5")
        log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            width=80,
            height=15,
            state='disabled',
            font=('Consolas', 9)
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(leveltime)s - %(message)s', datefmt='%H:%M:%S'))
        logging.getLogger().addHandler(text_handler)

    def _create_status_bar(self, parent):
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        status_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(status_frame, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E))

        author_label = ttk.Label(status_frame, text="Made with ❤️by Z🐻", foreground="gray")
        author_label.grid(row=0, column=1, padx=(10, 5))

        self.progress_bar = ttk.Progressbar(status_frame, mode='determinate')
        self.progress_bar.grid(row=0, column=2, padx=(5, 0), sticky=tk.E)
        self.progress_bar.grid_remove()

    def _select_input_file(self):
        filename = filedialog.askopenfilename(
            title="选择家长信息Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
            initialdir=self.config.get('last_input_dir', os.path.expanduser("~"))
        )

        if filename:
            self.input_file_path.set(filename)
            self.config.set('last_input_dir', str(Path(filename).parent))

            is_valid, error_msg = self.excel_handler.validate_excel_file(filename)
            if is_valid:
                self.logger.info(f"已选择输入文件：{filename}")

                if not self.output_file_path.get():
                    input_path = Path(filename)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_path = input_path.parent / f"查询结果_{timestamp}.xlsx"
                    self.output_file_path.set(str(output_path))

            else:
                messagebox.showerror("文件验证失败", f"选择的文件格式不正确：\n{error_msg}")
                self.input_file_path.set("")

    def _select_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="选择结果保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
            initialdir=self.config.get('last_output_dir', os.path.expanduser("~")),
            initialfile=f'查询结果_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        if filename:
            self.output_file_path.set(filename)
            self.config.set('last_output_dir', str(Path(filename).parent))
            self.logger.info(f"已设置输出文件：{filename}")

    def _start_query(self):
        if not self.input_file_path.get():
            messagebox.showwarning("提示", "请先选择输入文件")
            return
        if not self.output_file_path.get():
            messagebox.showwarning("提示", "请先选择输出文件保存位置")
            return

        self.config.set("session_id", self.session_id.get())
        self.config.set("max_retries", self.max_retries.get())
        self.config.save()

        self.is_running = True
        self.start_button.config(state='disabled')
        self.stop_button.config(state='normal')
        self.progress_bar.grid()
        self.progress_bar['value'] = 0

        thread = threading.Thread(target=self._run_query, daemon=True)
        thread.start()

    def _stop_query(self):
        self.is_running = False
        self.logger.warning("用户请求停止查询...")
        self.status_label.config(text="正在停止...")

    def _run_query(self):
        try:
            self.logger.info("正在读取信息...")
            self.status_label.config(text="正在读取Excel文件...")

            parent_info_list = self.excel_handler.read_parent_info(self.input_file_path.get())
            total_count = len(parent_info_list)

            if total_count == 0:
                messagebox.showwarning("提示", "输入文件没有有效的家长信息")
                return

            self.logger.info(f"共读取到 {total_count} 条信息")

            self.progress_bar['maximum'] = total_count

            session_id = self.session_id.get().strip() if self.session_id.get().strip() else None
            self.crawler = ResidencePointsCrawler(
                session_id = session_id,
                max_retries= self.max_retries.get()
            )

            query_results = []
            for idx, parent_info in enumerate(parent_info_list, start=1):
                if not self.is_running:
                    self.logger.warning("查询中止")
                    break

                name = parent_info['name']
                pid = parent_info['pid']

                self.logger.info(f"[{idx}/{total_count}] 开始查询：{name}")
                self.status_label.config(text=f"正在查询：{name} ({idx}/{total_count})")

                def progress_calllback(attempt, max_retries, message):
                    self.logger.debug(f"[{name}] {message}")

                result = self.crawler.query_points(name, pid, progress_calllback)
                query_results.append(result)

                self.progress_bar['value'] = idx

                if result['status'] == 'success':
                    self.logger.info(f"[{idx}/{total_count}] 查询成功：{name}")
                elif result['status'] == 'not_found':
                    self.logger.warning(f"[{idx}/{total_count}] 未找到记录：{name}")
                else:
                    self.logger.info(f"[{idx}/{total_count}] 查询失败：{name} - {result.get('error')}")

            if query_results:
                self.logger.info("正在保存查询结果...")
                self.status_label.config(text="正在保存结果...")

                self.excel_handler.write_results(
                    self.output_file_path.get(),
                    parent_info_list[:len(query_results)],
                    query_results
                )

                success_count = sum(1 for r in query_results if r['status'] == 'success')
                failed_count = sum(1 for r in query_results if r['status'] == 'failed')
                not_found_count = sum(1 for r in query_results if r['status'] == 'not_found')

                summary = (
                    f"查询完成！\n"
                    f"总共 {len(query_results)} 条！\n"
                    f"成功 {success_count} 条！\n"
                    f"未找到 {not_found_count} 条！\n"
                    f"失败 {failed_count} 条！\n"
                    f"结果保存到：\n{self.output_file_path.get()}"
                )

                self.logger.info(f"查询完成 - 成功：{success_count}, 未找到：{not_found_count}, 失败：{failed_count}")
                messagebox.showinfo("查询完成", summary)

                self.status_label.config(text="查询完成")

            else:
                self.logger.warning("没有完成任何查询")
                self.status_label.config(text="查询已停止")

        except Exception as e:
            self.logger.error(f"查询发生错误：{str(e)}", exc_info=True)
            messagebox.showerror("错误", f"查询发生错误：\n{str(e)}")
            self.status_label.config(text="查询出错")

        finally:
            if self.crawler:
                self.crawler.close()

            self.is_running = False
            self.start_button.config(state='normal')
            self.stop_button.config(state='disabled')

    def _clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')


    def _on_closing(self):
        if self.is_running:
            if messagebox.askokcancel("确认退出", "查询正在进行，确认要退出吗？"):
                self.is_running = False
                self.root.destroy()
        else:
            self.root.destroy()

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = tk.Tk()
    app = ResidencePointGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()






