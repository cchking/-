import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import zipfile
import rarfile
import py7zr
import tarfile
import gzip
import bz2
import lzma
import os
import sys
import traceback
import logging
import winreg

# 设置日志
logging.basicConfig(filename='unzip_tool.log', level=logging.DEBUG)

class SettingsWindow:
    def __init__(self, master):
        self.window = tk.Toplevel(master)
        self.window.title("设置")
        self.window.geometry("300x150")
        self.window.resizable(False, False)

        self.label = ttk.Label(self.window, text="设置为默认打开程序:", font=("Arial", 10))
        self.label.pack(pady=(20, 10))

        self.set_default_button = ttk.Button(self.window, text="设置为默认", command=self.set_as_default)
        self.set_default_button.pack()

    def set_as_default(self):
        file_types = ['.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tar.bz2', '.gz', '.bz2', '.xz']
        program_path = sys.argv[0]  # 使用当前脚本的路径
        icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "icon.ico"))

        try:
            self._set_as_default_internal(file_types, program_path, icon_path)
        except Exception as e:
            logging.error(f"设置默认程序失败: {str(e)}")
            logging.error(traceback.format_exc())
            messagebox.showerror("错误", f"设置失败: {str(e)}")

    def _set_as_default_internal(self, file_types, program_path, icon_path):
        for ext in file_types:
            key_path = f'Software\\Classes\\{ext}'
            try:
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    winreg.SetValue(key, '', winreg.REG_SZ, f'CompressedFile{ext}')
                    with winreg.CreateKey(key, 'DefaultIcon') as icon_key:
                        winreg.SetValue(icon_key, '', winreg.REG_SZ, f'{icon_path},0')
                    with winreg.CreateKey(key, 'shell\\open\\command') as command_key:
                        winreg.SetValue(command_key, '', winreg.REG_SZ, f'python "{program_path}" "%1"')
            except Exception as e:
                logging.error(f"设置 {ext} 失败: {str(e)}")
                raise

        messagebox.showinfo("成功", "已尝试设置为默认打开程序并修改图标。请刷新文件夹查看效果。")

class UnzipTool:
    def __init__(self, master):
        self.master = master
        master.title("yt解压")
        master.geometry("500x550")
        master.resizable(False, False)

        style = ttk.Style()
        style.theme_use("clam")

        main_frame = ttk.Frame(master, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.label = ttk.Label(main_frame, text="选择要解压的文件:", font=("Arial", 12))
        self.label.pack(pady=(0, 10))

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 20))

        self.select_button = ttk.Button(button_frame, text="选择文件", command=self.select_file, width=15)
        self.select_button.pack(side=tk.LEFT, padx=(0, 10))

        self.preview_button = ttk.Button(button_frame, text="预览文件", command=self.preview_file, width=15)
        self.preview_button.pack(side=tk.LEFT, padx=(0, 10))

        self.extract_button = ttk.Button(button_frame, text="解压文件", command=self.extract_file, width=15)
        self.extract_button.pack(side=tk.LEFT)

        self.settings_button = ttk.Button(main_frame, text="设置", command=self.open_settings, width=15)
        self.settings_button.pack(pady=(0, 20))

        self.status_label = ttk.Label(main_frame, text="", font=("Arial", 10), wraplength=460)
        self.status_label.pack(pady=(0, 20))

        self.preview_text = tk.Text(main_frame, height=10, width=60)
        self.preview_text.pack(pady=(0, 20))

        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=460, mode="determinate")
        self.progress_bar.pack(pady=(0, 10))

        self.progress_label = ttk.Label(main_frame, text="", font=("Arial", 10))
        self.progress_label.pack()

        self.file_path = None

        # 设置拖放功能
        self.master.drop_target_register(DND_FILES)
        self.master.dnd_bind('<<Drop>>', self.handle_drop)

    def open_settings(self):
        SettingsWindow(self.master)

    def handle_drop(self, event):
        file_path = event.data
        if file_path:
            self.file_path = file_path.strip('{}')  # 移除可能的大括号
            self.status_label.config(text=f"已选择文件: {os.path.basename(self.file_path)}")
            self.preview_file()

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[
            ("压缩文件", "*.zip;*.rar;*.7z;*.tar;*.tar.gz;*.tar.bz2;*.gz;*.bz2;*.xz")
        ])
        if self.file_path:
            self.status_label.config(text=f"已选择文件: {os.path.basename(self.file_path)}")

    def preview_file(self):
        if not self.file_path:
            self.status_label.config(text="请先选择一个文件")
            return

        try:
            file_list = self.get_file_list(self.file_path)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, "\n".join(file_list))
        except Exception as e:
            self.status_label.config(text=f"预览失败: {str(e)}")

    def get_file_list(self, file_path):
        if file_path.endswith('.zip'):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                return zip_ref.namelist()
        elif file_path.endswith('.rar'):
            with rarfile.RarFile(file_path, 'r') as rar_ref:
                return rar_ref.namelist()
        elif file_path.endswith('.7z'):
            with py7zr.SevenZipFile(file_path, 'r') as sz_ref:
                return sz_ref.getnames()
        elif file_path.endswith(('.tar', '.tar.gz', '.tar.bz2')):
            with tarfile.open(file_path, 'r:*') as tar_ref:
                return tar_ref.getnames()
        elif file_path.endswith('.gz'):
            return [os.path.basename(file_path[:-3])]
        elif file_path.endswith('.bz2'):
            return [os.path.basename(file_path[:-4])]
        elif file_path.endswith('.xz'):
            return [os.path.basename(file_path[:-3])]
        else:
            raise ValueError("不支持的文件格式")

    def extract_file(self):
        if not self.file_path:
            self.status_label.config(text="请先选择一个文件")
            return

        output_dir = filedialog.askdirectory(title="选择解压目标文件夹")
        if not output_dir:
            return

        try:
            total_size = self.get_file_size(self.file_path)
            extracted_size = 0

            if self.file_path.endswith('.zip'):
                with zipfile.ZipFile(self.file_path, 'r') as zip_ref:
                    for file in zip_ref.infolist():
                        zip_ref.extract(file, output_dir)
                        extracted_size += file.file_size
                        self.update_progress(extracted_size, total_size)

            elif self.file_path.endswith('.rar'):
                with rarfile.RarFile(self.file_path, 'r') as rar_ref:
                    for file in rar_ref.infolist():
                        rar_ref.extract(file, output_dir)
                        extracted_size += file.file_size
                        self.update_progress(extracted_size, total_size)

            elif self.file_path.endswith('.7z'):
                with py7zr.SevenZipFile(self.file_path, 'r') as sz_ref:
                    for file in sz_ref.getnames():
                        sz_ref.extract(output_dir, [file])
                        extracted_size += sz_ref.getmember(file).uncompressed
                        self.update_progress(extracted_size, total_size)

            elif self.file_path.endswith(('.tar', '.tar.gz', '.tar.bz2')):
                with tarfile.open(self.file_path, 'r:*') as tar_ref:
                    for member in tar_ref:
                        tar_ref.extract(member, output_dir)
                        extracted_size += member.size
                        self.update_progress(extracted_size, total_size)

            elif self.file_path.endswith('.gz'):
                with gzip.open(self.file_path, 'rb') as gz_ref:
                    with open(os.path.join(output_dir, os.path.basename(self.file_path[:-3])), 'wb') as out_f:
                        out_f.write(gz_ref.read())

            elif self.file_path.endswith('.bz2'):
                with bz2.open(self.file_path, 'rb') as bz2_ref:
                    with open(os.path.join(output_dir, os.path.basename(self.file_path[:-4])), 'wb') as out_f:
                        out_f.write(bz2_ref.read())

            elif self.file_path.endswith('.xz'):
                with lzma.open(self.file_path, 'rb') as xz_ref:
                    with open(os.path.join(output_dir, os.path.basename(self.file_path[:-3])), 'wb') as out_f:
                        out_f.write(xz_ref.read())

            else:
                raise ValueError("不支持的文件格式")

            self.status_label.config(text=f"文件已成功解压到: {output_dir}")
        except Exception as e:
            self.status_label.config(text=f"解压失败: {str(e)}")
        finally:
            self.progress_bar["value"] = 0
            self.progress_label.config(text="")

    def get_file_size(self, file_path):
        return os.path.getsize(file_path)

    def update_progress(self, extracted_size, total_size):
        progress = (extracted_size / total_size) * 100
        self.progress_bar["value"] = progress
        self.progress_label.config(text=f"{progress:.1f}%")
        self.master.update_idletasks()

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == '--set_default':
        SettingsWindow(None).set_as_default()
    else:
        root = TkinterDnD.Tk()
        # 设置窗口图标
        icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "icon.ico"))
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
        unzip_tool = UnzipTool(root)
        root.mainloop()
