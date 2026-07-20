#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
刀具数据提取与写入工具（无控制台后台静默运行版本）
功能：
1. 读取脚本所在目录下的所有 .txt 文件，提取刀具名称和对应的 L 值、FL 值
2. 遍历同目录下的 .xlsx 和 .xls 文件，在表格中查找对应的刀具名称
3. 将 L 值写入到对应的"装刀长"列中，将 FL 值写入到对应的"刃长"列中
4. 使用 WPS COM 接口操作 Excel，防止图片丢失
5. 使用 Tkinter 弹窗替代控制台输出，适配无黑框环境
"""

import os
import re
import sys
import glob
import tkinter
from tkinter import messagebox

# 尝试导入 win32com.client
try:
    import win32com.client
except ImportError:
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showerror(
        "错误",
        "未检测到 pywin32 库，请先安装：pip install pywin32"
    )
    root.destroy()
    sys.exit(1)


def get_script_directory():
    """获取当前脚本所在的目录"""
    if getattr(sys, 'frozen', False):
        # 如果是打包后的 exe，使用 sys.executable 的路径
        return os.path.dirname(sys.executable)
    else:
        # 如果是脚本直接运行，使用脚本文件的路径
        return os.path.dirname(os.path.abspath(__file__))


def extract_tool_data_from_txt(txt_path):
    """
    从单个 TXT 文件中提取刀具名称、L 值和 FL 值
    返回字典：{刀具名称: {'L': L值, 'FL': FL值}}
    """
    tool_data = {}

    try:
        with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print(f"读取文件 {txt_path} 失败: {e}")
        return tool_data

    # 按行处理，查找包含 T数字=、L= 和 FL= 的行
    lines = content.split('\n')

    for line in lines:
        # 匹配 T数字= 后面的刀具名称
        # 示例: (T01=CMF-4X90L40 D=4.      R=0.      FL=2.    L=40.    MinZ=47.03 Time=0.72M  )
        tool_match = re.search(r'\(T\d+=([^\s)]+)', line)
        if tool_match:
            tool_name = tool_match.group(1).strip()
            
            l_value = None
            fl_value = None

            # 精确匹配 L= 后面的数值，使用负向前瞻 (?<!F) 确保前面不是 F
            l_match = re.search(r'(?<!F)L=([0-9]+\.?[0-9]*)', line)
            if l_match:
                l_value = l_match.group(1).rstrip('.')

            # 匹配 FL= 后面的数值
            fl_match = re.search(r'FL=([0-9]+\.?[0-9]*)', line)
            if fl_match:
                fl_value = fl_match.group(1).rstrip('.')

            # 只要匹配到了名称，且 L 或 FL 中有一个有值，就存入字典
            if tool_name and (l_value is not None or fl_value is not None):
                tool_data[tool_name] = {'L': l_value, 'FL': fl_value}

    return tool_data


def extract_all_tool_data(directory):
    """
    读取目录下所有 .txt 文件，提取刀具数据
    返回合并后的字典
    """
    all_tool_data = {}

    # 查找所有 txt 文件
    txt_pattern = os.path.join(directory, '*.txt')
    txt_files = glob.glob(txt_pattern)

    if not txt_files:
        return all_tool_data

    for txt_file in txt_files:
        file_data = extract_tool_data_from_txt(txt_file)
        all_tool_data.update(file_data)

    return all_tool_data


def find_wps_application():
    """
    尝试启动 WPS 的 COM 接口
    依次尝试: Ket.Application, ET.Application, Kwps.Application
    返回成功启动的 application 对象，如果都失败则返回 None
    """
    wps_progid_list = [
        'Ket.Application',   # WPS 表格（旧版）
        'ET.Application',    # WPS 表格（新版）
        'Kwps.Application',   # WPS Office（通用）
    ]

    for progid in wps_progid_list:
        try:
            app = win32com.client.Dispatch(progid)
            return app
        except Exception:
            continue

    return None


def find_column_headers(ws, max_rows=30):
    """
    在表格前 max_rows 行中查找"刀具名称"、"装刀长"和"刃长"的列号
    返回: (刀具名称列号, 装刀长列号, 刃长列号, 表头行号)
    如果找不到"刀具名称"，返回 None
    """
    tool_name_col = None
    tool_length_col = None
    fl_length_col = None
    header_row = None

    # 遍历前 max_rows 行，前 20 列
    for row in range(1, max_rows + 1):
        for col in range(1, 21):
            try:
                cell_value = ws.Cells(row, col).Value
                if cell_value is None:
                    continue

                # 转换为字符串进行比较
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

                # 如果三个都找到了，提前返回
                if tool_name_col is not None and tool_length_col is not None and fl_length_col is not None:
                    return tool_name_col, tool_length_col, fl_length_col, header_row

            except Exception:
                continue

    # 只要找到了刀具名称和（装刀长或刃长其一），就返回
    if tool_name_col is not None and (tool_length_col is not None or fl_length_col is not None):
        return tool_name_col, tool_length_col, fl_length_col, header_row

    return None


def process_excel_file(excel_path, tool_data):
    """
    处理单个 Excel 文件
    在表格中查找刀具名称，并将 L 值写入装刀长列，FL值写入刃长列
    返回: (成功标志, 更新记录数)
    """
    # 启动 WPS
    app = find_wps_application()
    if app is None:
        raise RuntimeError("未检测到 WPS Office 环境，请确认已安装 WPS。")

    wb = None
    updated_count = 0

    try:
        # 设置后台静默运行
        app.Visible = False
        app.DisplayAlerts = False

        # 打开工作簿
        wb = app.Workbooks.Open(os.path.abspath(excel_path))
        ws = wb.ActiveSheet

        # 查找表头
        header_result = find_column_headers(ws)
        if header_result is None:
            return False, 0

        tool_name_col, tool_length_col, fl_length_col, header_row = header_result

        # 从表头下一行开始遍历数据
        data_start_row = header_row + 1
        max_search_row = header_row + 100  # 最多搜索 100 行数据

        for row in range(data_start_row, max_search_row + 1):
            # 获取当前行的刀具名称
            cell_value = ws.Cells(row, tool_name_col).Value
            if cell_value is None:
                continue

            tool_name = str(cell_value).strip()
            if not tool_name:
                continue

            # 在字典中查找对应的数据
            if tool_name in tool_data:
                data = tool_data[tool_name]
                l_value_str = data.get('L')
                fl_value_str = data.get('FL')
                
                row_updated = False

                # 写入 L 值到装刀长列
                if l_value_str is not None and tool_length_col is not None:
                    try:
                        ws.Cells(row, tool_length_col).Value = float(l_value_str)
                        row_updated = True
                    except ValueError:
                        pass
                
                # 写入 FL 值到刃长列
                if fl_value_str is not None and fl_length_col is not None:
                    try:
                        ws.Cells(row, fl_length_col).Value = float(fl_value_str)
                        row_updated = True
                    except ValueError:
                        pass
                
                if row_updated:
                    updated_count += 1

    finally:
        # 确保保存并关闭工作簿
        try:
            if wb is not None:
                wb.Save()
                wb.Close()
        except Exception:
            pass

        # 退出 WPS 应用
        try:
            app.Quit()
        except Exception:
            pass

    return True, updated_count


def main():
    """主函数"""
    # 初始化 Tkinter（隐藏主窗口）
    root = tkinter.Tk()
    root.withdraw()

    try:
        # 获取脚本所在目录
        script_dir = get_script_directory()

        # 第一步：提取 TXT 文件中的刀具数据
        tool_data = extract_all_tool_data(script_dir)

        if not tool_data:
            messagebox.showerror(
                "错误",
                "未提取到任何刀具数据，请检查 TXT 文件格式。"
            )
            return

        # 第二步：查找并处理 Excel 文件
        excel_files = []
        for pattern in ['*.xlsx', '*.xls']:
            excel_files.extend(glob.glob(os.path.join(script_dir, pattern)))

        # 去重并排序
        excel_files = sorted(list(set(excel_files)))

        if not excel_files:
            messagebox.showerror(
                "错误",
                f"在 {script_dir} 中未找到 .xlsx 或 .xls 文件"
            )
            return

        # 第三步：逐个处理 Excel 文件
        success_count = 0
        total_updated = 0

        for excel_file in excel_files:
            try:
                success, updated = process_excel_file(excel_file, tool_data)
                if success:
                    success_count += 1
                    total_updated += updated
            except Exception as e:
                messagebox.showerror(
                    "处理错误",
                    f"处理文件 {os.path.basename(excel_file)} 时出错：\n{str(e)}"
                )
                return

        # 处理完成，弹出成功提示
        messagebox.showinfo(
            "处理完成",
            f"处理完成！成功更新了 {success_count} 个 Excel 文件，共 {total_updated} 条记录。"
        )

    except Exception as e:
        messagebox.showerror(
            "程序错误",
            f"程序运行过程中发生错误：\n{str(e)}"
        )

    finally:
        # 销毁 Tkinter 根窗口
        try:
            root.destroy()
        except Exception:
            pass


if __name__ == '__main__':
    main()