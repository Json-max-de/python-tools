#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
刀具数据提取与写入工具（GUI 界面版）
功能：
1. 界面化操作，提供 TXT 和 Excel 文件列表供用户选择
2. 强制 1对1 转化：选择一份 TXT 和一份 Excel，合并数据
3. 保护原文件：通过复制原 Excel 生成新文件并在新文件上修改
4. 依赖 WPS COM 接口
"""

import os
import re
import sys
import shutil
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# 尝试导入 win32com.client
try:
    import win32com.client
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("错误", "未检测到 pywin32 库，请先安装：pip install pywin32")
    root.destroy()
    sys.exit(1)


def get_script_directory():
    """获取当前脚本或 exe 所在的目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def extract_tool_data_from_txt(txt_path):
    """从单个 TXT 文件中提取刀具名称、L 值和 FL 值"""
    tool_data = {}
    try:
        with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        raise Exception(f"读取 TXT 文件失败: {e}")

    lines = content.split('\n')
    for line in lines:
        tool_match = re.search(r'\(T\d+=([^\s)]+)', line)
        if tool_match:
            tool_name = tool_match.group(1).strip()
            l_value, fl_value = None, None

            l_match = re.search(r'(?<!F)L=([0-9]+\.?[0-9]*)', line)
            if l_match:
                l_value = l_match.group(1).rstrip('.')

            fl_match = re.search(r'FL=([0-9]+\.?[0-9]*)', line)
            if fl_match:
                fl_value = fl_match.group(1).rstrip('.')

            if tool_name and (l_value is not None or fl_value is not None):
                tool_data[tool_name] = {'L': l_value, 'FL': fl_value}

    return tool_data


def find_wps_application():
    """尝试启动 WPS 的 COM 接口"""
    wps_progid_list = ['Ket.Application', 'ET.Application', 'Kwps.Application']
    for progid in wps_progid_list:
        try:
            return win32com.client.Dispatch(progid)
        except Exception:
            continue
    return None


def find_column_headers(ws, max_rows=30):
    """在表格前 max_rows 行中查找表头"""
    tool_name_col = None
    tool_length_col = None
    fl_length_col = None
    header_row = None

    for row in range(1, max_rows + 1):
        for col in range(1, 21):
            try:
                cell_value = ws.Cells(row, col).Value
                if cell_value is None:
                    continue

                cell_str = str(cell_value).strip()
                if cell_str == '刀具名称':
                    tool_name_col = col
                    header_row = row
                elif cell_str == '装刀长':
                    tool_length_col = col
                    header_row = row
                elif cell_str == '刃长':
                    fl_length_col = col
                    header_row = row

                if tool_name_col is not None and tool_length_col is not None and fl_length_col is not None:
                    return tool_name_col, tool_length_col, fl_length_col, header_row
            except Exception:
                continue

    if tool_name_col is not None and (tool_length_col is not None or fl_length_col is not None):
        return tool_name_col, tool_length_col, fl_length_col, header_row

    return None


def process_excel_file(excel_path, tool_data):
    """处理目标 Excel 文件（注意：传入的 excel_path 已经是新复制的独立文件）"""
    app = find_wps_application()
    if app is None:
        raise RuntimeError("未检测到 WPS Office 环境，请确认已安装 WPS。")

    wb = None
    updated_count = 0

    try:
        app.Visible = False
        app.DisplayAlerts = False

        # 打开复制后的新工作簿
        wb = app.Workbooks.Open(os.path.abspath(excel_path))
        ws = wb.ActiveSheet

        header_result = find_column_headers(ws)
        if header_result is None:
            raise Exception("在 Excel 中未找到'刀具名称'及对应的长度列头，请检查表格格式。")

        tool_name_col, tool_length_col, fl_length_col, header_row = header_result
        data_start_row = header_row + 1
        max_search_row = header_row + 100

        for row in range(data_start_row, max_search_row + 1):
            cell_value = ws.Cells(row, tool_name_col).Value
            if cell_value is None:
                continue

            tool_name = str(cell_value).strip()
            if not tool_name or tool_name not in tool_data:
                continue

            data = tool_data[tool_name]
            l_value_str = data.get('L')
            fl_value_str = data.get('FL')
            row_updated = False

            if l_value_str is not None and tool_length_col is not None:
                try:
                    ws.Cells(row, tool_length_col).Value = float(l_value_str)
                    row_updated = True
                except ValueError:
                    pass

            if fl_value_str is not None and fl_length_col is not None:
                try:
                    ws.Cells(row, fl_length_col).Value = float(fl_value_str)
                    row_updated = True
                except ValueError:
                    pass

            if row_updated:
                updated_count += 1

        wb.Save()

    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False) # 已经 Save 过了，安全关闭
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass

    return updated_count


class ToolDataApp:
    def __init__(self, root):
        self.root = root
        self.root.title("刀具数据转化工具")
        self.root.geometry("650x450")
        self.root.configure(bg="#E0E0E0")  # 浅灰背景，贴近设计图
        
        # 定义目录
        self.script_dir = get_script_directory()
        self.txt_dir = os.path.join(self.script_dir, "TXT文件")
        self.excel_dir = os.path.join(self.script_dir, "Excel文件")
        
        # 确保目录存在
        os.makedirs(self.txt_dir, exist_ok=True)
        os.makedirs(self.excel_dir, exist_ok=True)

        self.setup_ui()
        self.refresh_lists()

    def setup_ui(self):
        # 容器框架
        main_frame = tk.Frame(self.root, bg="#E0E0E0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # ====== 容器 1 (TXT) ======
        frame_txt = tk.Frame(main_frame, bg="white", bd=0)
        frame_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        tk.Label(frame_txt, text="容器 1", bg="#E0E0E0", fg="gray", anchor="w").pack(fill=tk.X)
        tk.Label(frame_txt, text="■ 文件。txt", bg="white", font=("微软雅黑", 14)).pack(pady=10)
        
        # TXT 列表框（加入 exportselection=False 允许两边同时选中）
        self.listbox_txt = tk.Listbox(frame_txt, font=("微软雅黑", 10), selectbackground="#A0A0A0", relief=tk.FLAT, exportselection=False)
        self.listbox_txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # ====== 容器 2 (Excel) ======
        frame_excel = tk.Frame(main_frame, bg="white", bd=0)
        frame_excel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 10))
        
        tk.Label(frame_excel, text="容器 2", bg="#E0E0E0", fg="gray", anchor="w").pack(fill=tk.X)
        tk.Label(frame_excel, text="■ 文件。excel", bg="white", font=("微软雅黑", 14)).pack(pady=10)
        
        # Excel 列表框（加入 exportselection=False 允许两边同时选中）
        self.listbox_excel = tk.Listbox(frame_excel, font=("微软雅黑", 10), selectbackground="#A0A0A0", relief=tk.FLAT, exportselection=False)
        self.listbox_excel.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # ====== 右侧操作区 (按钮) ======
        frame_action = tk.Frame(main_frame, bg="#E0E0E0", width=120)
        frame_action.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0))
        frame_action.pack_propagate(False) # 固定宽度
        
        # 红色转换按钮
        self.btn_convert = tk.Button(
            frame_action, 
            text="开始转化", 
            bg="#D32F2F",     # 红色
            fg="white", 
            font=("微软雅黑", 12, "bold"),
            relief=tk.FLAT,
            command=self.execute_conversion
        )
        self.btn_convert.pack(pady=35, fill=tk.X)
        
        # 刷新按钮 (方便用户放入文件后直接刷新列表)
        btn_refresh = tk.Button(
            frame_action, 
            text="刷新列表", 
            command=self.refresh_lists
        )
        btn_refresh.pack(fill=tk.X, pady=10)

    def refresh_lists(self):
        """读取文件夹并刷新列表显示"""
        self.listbox_txt.delete(0, tk.END)
        self.listbox_excel.delete(0, tk.END)

        # 读取 txt 文件
        if os.path.exists(self.txt_dir):
            for file in os.listdir(self.txt_dir):
                if file.lower().endswith('.txt'):
                    self.listbox_txt.insert(tk.END, file)

        # 读取 excel 文件
        if os.path.exists(self.excel_dir):
            for file in os.listdir(self.excel_dir):
                if file.lower().endswith(('.xlsx', '.xls')):
                    self.listbox_excel.insert(tk.END, file)

    def execute_conversion(self):
        """执行转换逻辑"""
        # 1. 获取选中的文件
        txt_sel = self.listbox_txt.curselection()
        excel_sel = self.listbox_excel.curselection()

        if not txt_sel:
            messagebox.showwarning("提示", "请在左侧选择一份 TXT 文件！")
            return
        if not excel_sel:
            messagebox.showwarning("提示", "请在中间选择一份 Excel 文件！")
            return

        txt_filename = self.listbox_txt.get(txt_sel[0])
        excel_filename = self.listbox_excel.get(excel_sel[0])

        txt_path = os.path.join(self.txt_dir, txt_filename)
        source_excel_path = os.path.join(self.excel_dir, excel_filename)

        # 2. 提取 TXT 数据
        try:
            tool_data = extract_tool_data_from_txt(txt_path)
            if not tool_data:
                messagebox.showwarning("提示", f"从 {txt_filename} 中未提取到任何刀具数据，请检查文件格式。")
                return
        except Exception as e:
            messagebox.showerror("读取 TXT 失败", str(e))
            return

        # 3. 构建输出文件路径并安全复制原文件
        # 输出格式：原文件名+年月日时分.xlsx (例如: filename202608230835.xlsx)
        timestamp = datetime.now().strftime("%Y%m%d%H%M")
        base_name = os.path.splitext(excel_filename)[0]
        # 强制将新文件存为 .xlsx
        output_filename = f"{base_name}{timestamp}.xlsx"
        output_path = os.path.join(self.script_dir, output_filename)

        try:
            # 使用 shutil.copy2 复制源文件到根目录，保证不破坏源文件夹内的模板
            shutil.copy2(source_excel_path, output_path)
        except Exception as e:
            messagebox.showerror("文件复制失败", f"无法创建新文件：\n{str(e)}")
            return

        # 4. 操作新生成的 Excel 文件
        self.btn_convert.config(text="处理中...", state=tk.DISABLED, bg="gray")
        self.root.update()

        try:
            updated_count = process_excel_file(output_path, tool_data)
            messagebox.showinfo(
                "处理完成", 
                f"转化成功！\n\n共更新 {updated_count} 条记录。\n新文件已生成在程序目录下：\n{output_filename}"
            )
        except Exception as e:
            messagebox.showerror("Excel 处理错误", str(e))
            # 如果出错，删除复制出来的残次文件
            if os.path.exists(output_path):
                os.remove(output_path)
        finally:
            self.btn_convert.config(text="开始转化", state=tk.NORMAL, bg="#D32F2F")


if __name__ == '__main__':
    root = tk.Tk()
    app = ToolDataApp(root)
    root.mainloop()