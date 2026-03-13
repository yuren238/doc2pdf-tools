# -*- coding: utf-8 -*-
"""
文档转PDF工具（Windows 专用）

支持将以下文件转换为PDF：
- Excel文件：面积汇总表.xls/xlsx（工作表1）
- Word文档：审查申请表.doc/.docx
- Word文档：地籍调查报告中的"质量承诺书"页面
- Word文档：地籍调查报告中的"地籍调查成果验收报告"页面

输出为黑白PDF，文件保存到脚本所在目录的 out 文件夹。
"""

import os
import sys
import glob
import traceback
from datetime import datetime

# ============= 配置区域 =============
# exe 打包后 __file__ 指向临时目录，需用 sys.executable 获取 exe 所在目录
if getattr(sys, 'frozen', False):
    ROOT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(ROOT_DIR, 'out')
ERROR_LOG_FILE = os.path.join(ROOT_DIR, 'error.txt')
# ============= 配置区域结束 =============


def write_error_log(error_message):
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ERROR_LOG_FILE, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write("文档转PDF工具 - 错误日志\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"错误发生时间： {timestamp}\n")
        f.write(f"脚本位置： {ROOT_DIR}\n")
        f.write(f"Python版本： {sys.version}\n")
        f.write(f"操作系统： {sys.platform}\n")
        f.write("\n" + "-" * 60 + "\n")
        f.write("错误详情：\n")
        f.write("-" * 60 + "\n\n")
        f.write(error_message)
        f.write("\n\n" + "-" * 60 + "\n")
        f.write("建议检查：\n")
        f.write("-" * 60 + "\n")
        f.write("1. 确认已安装 Microsoft Office (Excel/Word)\n")
        f.write("2. 确认已安装 Python 3.7 及以上版本\n")
        f.write("3. 尝试手动运行: pip install pywin32\n")
        f.write("4. 检查文件是否被其他程序占用\n")
        f.write("5. 确认文件格式正确（xls/xlsx/doc/docx）\n")
    print(f"\nError log saved to: {ERROR_LOG_FILE}")


def pause_and_exit(code=0, error_msg=None):
    
    if error_msg and code != 0:
        write_error_log(error_msg)
    print()
    os.system("pause")
    sys.exit(code)



# ─────────────────────────────────────────────
#  黑白转换
# ─────────────────────────────────────────────

def strip_color_from_stream(data: bytes) -> bytes:
    """
    扫描 PDF 页面内容流，将所有颜色设置指令替换为黑色。

    PDF 颜色操作符：
      rg / RG   — DeviceRGB 填充/描边（3个参数）
      k  / K    — DeviceCMYK 填充/描边（4个参数）
      scn/ SCN  — 通用颜色（参数数量不定，此处处理1-4个数字的情况）
      sc / SC   — 同上简化版
    全部替换为 DeviceGray 的等效操作符 g/G（0 = 黑色）。
    """
    import re

    # RGB 填充：r g b rg  → 0 g
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg', b'0 g', data)
    # RGB 描边：r g b RG  → 0 G
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+RG', b'0 G', data)
    # CMYK 填充：c m y k  → 0 g
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+k', b'0 g', data)
    # CMYK 描边：c m y k  → 0 G
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+K', b'0 G', data)
    # scn/sc 填充（1-4个参数）→ 0 g
    data = re.sub(rb'(?:[\d.]+\s+){1,4}sc(?:n)?', b'0 g', data)
    # SCN/SC 描边（1-4个参数）→ 0 G
    data = re.sub(rb'(?:[\d.]+\s+){1,4}SC(?:N)?', b'0 G', data)
    return data


def convert_pdf_to_grayscale(input_pdf, output_pdf):
    """将彩色 PDF 转换为黑白：修改页面内容流中的颜色操作符。"""
    import pypdf

    reader = pypdf.PdfReader(input_pdf)
    writer = pypdf.PdfWriter()

    for page in reader.pages:
        if '/Contents' in page:
            contents = page['/Contents']
            if isinstance(contents, pypdf.generic.ArrayObject):
                obj_list = [c.get_object() for c in contents]
            else:
                obj_list = [contents.get_object()]

            for stream_obj in obj_list:
                try:
                    raw = stream_obj.get_data()
                    new_raw = strip_color_from_stream(raw)
                    if new_raw != raw:
                        stream_obj.set_data(new_raw)
                except Exception:
                    pass

        writer.add_page(page)

    with open(output_pdf, 'wb') as f:
        writer.write(f)


# ─────────────────────────────────────────────
#  菜单
# ─────────────────────────────────────────────

def show_menu():
    
    print("\n" + "=" * 50)
    print("  文档转PDF转换工具")
    print("=" * 50)
    print("  请选择要转换的文件类型：")
    print()
    print("  1. 转换面积汇总表（Excel 文件）")
    print("  2. 转换审查申请表（Word 文档）")
    print("  3. 转换地籍调查报告-质量承诺书页面")
    print("  4. 转换地籍调查报告-地籍调查成果验收报告页面")
    print()
    print("  支持多选：")
    print("    输入 1,2,3,4 → 转换所有")
    print("    输入 34      → 转换质量承诺书和验收报告")
    print("    直接回车     → 转换全部")
    print("    输入 0       → 退出")
    print()

    while True:
        try:
            choice = input("  请输入选项：").strip()
        except KeyboardInterrupt:
            print()
            return ['0']

        if choice == '':
            return ['1', '2', '3', '4']
        if choice == '0':
            return ['0']

        normalized = choice.replace(',', '').replace('，', '').replace(' ', '')
        choices = list(dict.fromkeys(normalized))

        if choices and all(c in ('1', '2', '3', '4') for c in choices):
            return choices

        print("  输入无效，请重新输入（可输入 1、2、3、4 或组合如 1234、1,2,3,4）")


# ─────────────────────────────────────────────
#  转换函数
# ─────────────────────────────────────────────

def excel_to_pdf(excel_file, output_pdf):
    """将 Excel 工作表1导出为黑白 PDF。"""
    filename = os.path.basename(excel_file)
    print(f"▶ 处理 Excel：{filename}")

    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(excel_file))
        ws = wb.Worksheets(1)
        print(f"   工作表名称：{ws.Name}")

        tmp_pdf = output_pdf + ".tmp.pdf"
        ws.ExportAsFixedFormat(
            Type=0,
            Filename=os.path.abspath(tmp_pdf),
            Quality=0,
            IncludeDocProperties=False,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )

        print("   正在转换为黑白...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   ✅ 已保存（黑白）：{output_pdf}\n")
        return True

    except Exception as e:
        print(f"   ❌ 失败：{e}\n")
        traceback.print_exc()
        return False
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def word_to_pdf(word_file, output_pdf):
    """将 Word 文档导出为黑白 PDF。"""
    filename = os.path.basename(word_file)
    print(f"▶ 处理 Word：{filename}")

    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(os.path.abspath(word_file))

        tmp_pdf = output_pdf + ".tmp.pdf"
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(tmp_pdf),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            IncludeDocProps=False,
        )

        print("   正在转换为黑白...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   ✅ 已保存（黑白）：{output_pdf}\n")
        return True

    except Exception as e:
        print(f"   ❌ 失败：{e}\n")
        traceback.print_exc()
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        import time; time.sleep(1)  # 等待 Word 进程完全退出


def word_page_to_pdf(word_file, output_pdf, page_num, page_name=""):
    """将 Word 文档的指定页面导出为黑白 PDF。"""
    filename = os.path.basename(word_file)
    print(f"▶ 处理 Word：{filename}（{page_name} - 第{page_num}页）")

    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(os.path.abspath(word_file))

        tmp_pdf = output_pdf + ".tmp.pdf"
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(tmp_pdf),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=3,
            From=page_num,
            To=page_num,
            IncludeDocProps=False,
        )

        print("   正在转换为黑白...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   ✅ 已保存（黑白）：{output_pdf}\n")
        return True

    except Exception as e:
        print(f"   ❌ 失败：{e}\n")
        traceback.print_exc()
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        import time; time.sleep(1)  # 等待 Word 进程完全退出


def find_page_by_keyword(word_file, keyword, occurrence=1):
    """
    用 Word Find 对象搜索关键词所在的绝对页码。

    Args:
        word_file: Word文件路径
        keyword: 搜索关键词
        occurrence: 第几次出现（默认1，即第一次）

    Returns:
        页码（整数）或 None（未找到）
    """
    import win32com.client
    import pythoncom
    import time

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        max_retries = 3
        for attempt in range(max_retries):
            try:
                doc = word.Documents.Open(os.path.abspath(word_file), ReadOnly=True)
                break
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"   文件打开失败，等待1秒后重试... ({attempt + 1}/{max_retries})")
                    time.sleep(1)
                else:
                    raise e

        find = doc.Content.Find
        find.Text = keyword
        find.Forward = True
        find.Wrap = 0

        found_count = 0
        while find.Execute():
            found_count += 1
            if found_count == occurrence:
                page = find.Parent.Information(1)
                return page

        return None

    except Exception as e:
        print(f"   查找页面时出错：{e}")
        return None
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        time.sleep(0.5)


# ─────────────────────────────────────────────
#  主流程
# ─────────────────────────────────────────────

def main():
    os.chdir(ROOT_DIR)

    choices = show_menu()
    if '0' in choices:
        print("\nProgram exited.")
        pause_and_exit(0)

    need_excel = '1' in choices
    need_word = '2' in choices or '3' in choices or '4' in choices


    def find_files(pattern):
        return [f for f in glob.glob(pattern, recursive=True)
                if not os.path.basename(f).startswith('~')]

    
    excel_files = find_files(os.path.join(ROOT_DIR, '**', '*面积汇总表*.xls*')) if need_excel else []
    word_files = find_files(os.path.join(ROOT_DIR, '**', '*审查申请表*.doc*')) if '2' in choices else []
    survey_files_3 = find_files(os.path.join(ROOT_DIR, '**', '*地籍调查报告*.doc*')) if '3' in choices else []
    survey_files_4 = find_files(os.path.join(ROOT_DIR, '**', '*地籍调查报告*.doc*')) if '4' in choices else []

    # 查找质量承诺书页码
    quality_tasks = []
    if '3' in choices and survey_files_3:
        print("\n正在查找地籍调查报告中的'质量承诺书'页面...")
        for f in survey_files_3:
            page_num = find_page_by_keyword(f, "质量承诺书", occurrence=1)
            if page_num:
                print(f"   ✅ {os.path.basename(f)} - 质量承诺书在第 {page_num} 页")
                quality_tasks.append((f, page_num))
            else:
                print(f"   ❌ {os.path.basename(f)} - 未找到'质量承诺书'")

    # 查找验收报告页码
    acceptance_tasks = []
    if '4' in choices and survey_files_4:
        print("\n正在查找地籍调查报告中的'地籍调查成果验收报告'页面...")
        for f in survey_files_4:
            page_num = find_page_by_keyword(f, "地籍调查成果验收报告", occurrence=2)
            if page_num:
                print(f"   ✅ {os.path.basename(f)} - 质量承诺书在第 {page_num} 页")
                acceptance_tasks.append((f, page_num))
            else:
                print(f"   ❌ {os.path.basename(f)} - 未找到'质量承诺书'")

    # 构建任务列表
    tasks = []
    for f in excel_files:
        tasks.append({'file': f, 'type': 'excel', 'convert_fn': excel_to_pdf, 'extra': None, 'desc': '面积汇总表'})
    for f in word_files:
        tasks.append({'file': f, 'type': 'word', 'convert_fn': word_to_pdf, 'extra': None, 'desc': '审查申请表'})
    for f, page in quality_tasks:
        tasks.append({'file': f, 'type': 'quality', 'convert_fn': word_page_to_pdf, 'extra': page, 'desc': f'质量承诺书(第{page}页)'})
    for f, page in acceptance_tasks:
        tasks.append({'file': f, 'type': 'acceptance', 'convert_fn': word_page_to_pdf, 'extra': page, 'desc': f'验收报告(第{page}页)'})

    if not tasks:
        print("\nNo matching files found.")
        pause_and_exit(0)

    print(f"\nFound {len(tasks)} files to convert:")
    print("-" * 50)
    for i, task in enumerate(tasks, 1):
        print(f"  {i}. [{task['desc']}] {os.path.basename(task['file'])}")
    print("-" * 50)
    print()

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    ok_list = []
    fail_list = []

    for i, task in enumerate(tasks, 1):
        print(f"[{i}/{len(tasks)}] ", end="")
        stem = os.path.splitext(os.path.basename(task['file']))[0]

        if task['type'] == 'quality':
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}_质量承诺书.pdf")
            success = task['convert_fn'](task['file'], output_pdf, task['extra'], "质量承诺书")
        elif task['type'] == 'acceptance':
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}_地籍调查成果验收报告.pdf")
            success = task['convert_fn'](task['file'], output_pdf, task['extra'], "地籍调查成果验收报告")
        else:
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}.pdf")
            success = task['convert_fn'](task['file'], output_pdf)

        if success:
            ok_list.append(os.path.basename(task['file']))
        else:
            fail_list.append(os.path.basename(task['file']))

    print("=" * 50)
    print("  转换结果汇总")
    print("=" * 50)
    print(f"  成功：{len(ok_list)} 个")
    print(f"  失败：{len(fail_list)} 个")
    if fail_list:
        print("\n  失败的文件：")
        for f in fail_list:
            print(f"    - {f}")
    print()
    print(f"  PDF 文件保存位置：{OUTPUT_DIR}")
    print("=" * 50)

    pause_and_exit(0)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = f"错误类型：{type(e).__name__}\n错误信息：{str(e)}\n\n详细堆栈信息：\n{traceback.format_exc()}"
        print(f"\nProgram error: {e}")
        pause_and_exit(1, error_msg)# -*- coding: utf-8 -*-
"""
文档转PDF工具（Windows 专用）

支持将以下文件转换为PDF：
- Excel文件：面积汇总表.xls/xlsx（工作表1）
- Word文档：审查申请表.doc/.docx
- Word文档：地籍调查报告中的"质量承诺书"页面
- Word文档：地籍调查报告中的"地籍调查成果验收报告"页面

输出为黑白PDF，文件保存到脚本所在目录的 out 文件夹。
"""

import os
import sys
import glob
import traceback
from datetime import datetime

# ============= 配置区域 =============
# exe 打包后 __file__ 指向临时目录，需用 sys.executable 获取 exe 所在目录
if getattr(sys, 'frozen', False):
    ROOT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(ROOT_DIR, 'out')
ERROR_LOG_FILE = os.path.join(ROOT_DIR, 'error.txt')
# ============= 配置区域结束 =============


def write_error_log(error_message):
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ERROR_LOG_FILE, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write("文档转PDF工具 - 错误日志\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"错误发生时间： {timestamp}\n")
        f.write(f"脚本位置： {ROOT_DIR}\n")
        f.write(f"Python版本： {sys.version}\n")
        f.write(f"操作系统： {sys.platform}\n")
        f.write("\n" + "-" * 60 + "\n")
        f.write("错误详情：\n")
        f.write("-" * 60 + "\n\n")
        f.write(error_message)
        f.write("\n\n" + "-" * 60 + "\n")
        f.write("建议检查：\n")
        f.write("-" * 60 + "\n")
        f.write("1. 确认已安装 Microsoft Office (Excel/Word)\n")
        f.write("2. 确认已安装 Python 3.7 及以上版本\n")
        f.write("3. 尝试手动运行: pip install pywin32\n")
        f.write("4. 检查文件是否被其他程序占用\n")
        f.write("5. 确认文件格式正确（xls/xlsx/doc/docx）\n")
    print(f"\nError log saved to: {ERROR_LOG_FILE}")


def pause_and_exit(code=0, error_msg=None):
    
    if error_msg and code != 0:
        write_error_log(error_msg)
    print()
    os.system("pause")
    sys.exit(code)



# ─────────────────────────────────────────────
#  黑白转换
# ─────────────────────────────────────────────

def strip_color_from_stream(data: bytes) -> bytes:
    """
    扫描 PDF 页面内容流，将所有颜色设置指令替换为黑色。

    PDF 颜色操作符：
      rg / RG   — DeviceRGB 填充/描边（3个参数）
      k  / K    — DeviceCMYK 填充/描边（4个参数）
      scn/ SCN  — 通用颜色（参数数量不定，此处处理1-4个数字的情况）
      sc / SC   — 同上简化版
    全部替换为 DeviceGray 的等效操作符 g/G（0 = 黑色）。
    """
    import re

    # RGB 填充：r g b rg  → 0 g
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg', b'0 g', data)
    # RGB 描边：r g b RG  → 0 G
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+RG', b'0 G', data)
    # CMYK 填充：c m y k  → 0 g
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+k', b'0 g', data)
    # CMYK 描边：c m y k  → 0 G
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+K', b'0 G', data)
    # scn/sc 填充（1-4个参数）→ 0 g
    data = re.sub(rb'(?:[\d.]+\s+){1,4}sc(?:n)?', b'0 g', data)
    # SCN/SC 描边（1-4个参数）→ 0 G
    data = re.sub(rb'(?:[\d.]+\s+){1,4}SC(?:N)?', b'0 G', data)
    return data


def convert_pdf_to_grayscale(input_pdf, output_pdf):
    """将彩色 PDF 转换为黑白：修改页面内容流中的颜色操作符。"""
    import pypdf

    reader = pypdf.PdfReader(input_pdf)
    writer = pypdf.PdfWriter()

    for page in reader.pages:
        if '/Contents' in page:
            contents = page['/Contents']
            if isinstance(contents, pypdf.generic.ArrayObject):
                obj_list = [c.get_object() for c in contents]
            else:
                obj_list = [contents.get_object()]

            for stream_obj in obj_list:
                try:
                    raw = stream_obj.get_data()
                    new_raw = strip_color_from_stream(raw)
                    if new_raw != raw:
                        stream_obj.set_data(new_raw)
                except Exception:
                    pass

        writer.add_page(page)

    with open(output_pdf, 'wb') as f:
        writer.write(f)


# ─────────────────────────────────────────────
#  菜单
# ─────────────────────────────────────────────

def show_menu():
    
    print("\n" + "=" * 50)
    print("  文档转PDF转换工具")
    print("=" * 50)
    print("  请选择要转换的文件类型：")
    print()
    print("  1. 转换面积汇总表（Excel 文件）")
    print("  2. 转换审查申请表（Word 文档）")
    print("  3. 转换地籍调查报告-质量承诺书页面")
    print("  4. 转换地籍调查报告-地籍调查成果验收报告页面")
    print()
    print("  支持多选：")
    print("    输入 1,2,3,4 → 转换所有")
    print("    输入 34      → 转换质量承诺书和验收报告")
    print("    直接回车     → 转换全部")
    print("    输入 0       → 退出")
    print()

    while True:
        try:
            choice = input("  请输入选项：").strip()
        except KeyboardInterrupt:
            print()
            return ['0']

        if choice == '':
            return ['1', '2', '3', '4']
        if choice == '0':
            return ['0']

        normalized = choice.replace(',', '').replace('，', '').replace(' ', '')
        choices = list(dict.fromkeys(normalized))

        if choices and all(c in ('1', '2', '3', '4') for c in choices):
            return choices

        print("  输入无效，请重新输入（可输入 1、2、3、4 或组合如 1234、1,2,3,4）")


# ─────────────────────────────────────────────
#  转换函数
# ─────────────────────────────────────────────

def excel_to_pdf(excel_file, output_pdf):
    """将 Excel 工作表1导出为黑白 PDF。"""
    filename = os.path.basename(excel_file)
    print(f"▶ 处理 Excel：{filename}")

    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(excel_file))
        ws = wb.Worksheets(1)
        print(f"   工作表名称：{ws.Name}")

        tmp_pdf = output_pdf + ".tmp.pdf"
        ws.ExportAsFixedFormat(
            Type=0,
            Filename=os.path.abspath(tmp_pdf),
            Quality=0,
            IncludeDocProperties=False,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )

        print("   正在转换为黑白...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   ✅ 已保存（黑白）：{output_pdf}\n")
        return True

    except Exception as e:
        print(f"   ❌ 失败：{e}\n")
        traceback.print_exc()
        return False
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def word_to_pdf(word_file, output_pdf):
    """将 Word 文档导出为黑白 PDF。"""
    filename = os.path.basename(word_file)
    print(f"▶ 处理 Word：{filename}")

    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(os.path.abspath(word_file))

        tmp_pdf = output_pdf + ".tmp.pdf"
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(tmp_pdf),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            IncludeDocProps=False,
        )

        print("   正在转换为黑白...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   ✅ 已保存（黑白）：{output_pdf}\n")
        return True

    except Exception as e:
        print(f"   ❌ 失败：{e}\n")
        traceback.print_exc()
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def word_page_to_pdf(word_file, output_pdf, page_num, page_name=""):
    """将 Word 文档的指定页面导出为黑白 PDF。"""
    filename = os.path.basename(word_file)
    print(f"▶ 处理 Word：{filename}（{page_name} - 第{page_num}页）")

    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(os.path.abspath(word_file))

        tmp_pdf = output_pdf + ".tmp.pdf"
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(tmp_pdf),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=3,
            From=page_num,
            To=page_num,
            IncludeDocProps=False,
        )

        print("   正在转换为黑白...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   ✅ 已保存（黑白）：{output_pdf}\n")
        return True

    except Exception as e:
        print(f"   ❌ 失败：{e}\n")
        traceback.print_exc()
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def find_page_by_keyword(word_file, keyword, occurrence=1):
    """
    用 Word Find 对象搜索关键词所在的绝对页码。

    Args:
        word_file: Word文件路径
        keyword: 搜索关键词
        occurrence: 第几次出现（默认1，即第一次）

    Returns:
        页码（整数）或 None（未找到）
    """
    import win32com.client
    import pythoncom
    import time

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        max_retries = 3
        for attempt in range(max_retries):
            try:
                doc = word.Documents.Open(os.path.abspath(word_file), ReadOnly=True)
                break
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"   文件打开失败，等待1秒后重试... ({attempt + 1}/{max_retries})")
                    time.sleep(1)
                else:
                    raise e

        find = doc.Content.Find
        find.Text = keyword
        find.Forward = True
        find.Wrap = 0

        found_count = 0
        while find.Execute():
            found_count += 1
            if found_count == occurrence:
                page = find.Parent.Information(1)
                return page

        return None

    except Exception as e:
        print(f"   查找页面时出错：{e}")
        return None
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        time.sleep(0.5)


# ─────────────────────────────────────────────
#  主流程
# ─────────────────────────────────────────────

def main():
    os.chdir(ROOT_DIR)

    choices = show_menu()
    if '0' in choices:
        print("\nProgram exited.")
        pause_and_exit(0)

    need_excel = '1' in choices
    need_word = '2' in choices or '3' in choices or '4' in choices


    def find_files(pattern):
        return [f for f in glob.glob(pattern, recursive=True)
                if not os.path.basename(f).startswith('~')]

    
    excel_files = find_files(os.path.join(ROOT_DIR, '**', '*面积汇总表*.xls*')) if need_excel else []
    word_files = find_files(os.path.join(ROOT_DIR, '**', '*审查申请表*.doc*')) if '2' in choices else []
    survey_files_3 = find_files(os.path.join(ROOT_DIR, '**', '*地籍调查报告*.doc*')) if '3' in choices else []
    survey_files_4 = find_files(os.path.join(ROOT_DIR, '**', '*地籍调查报告*.doc*')) if '4' in choices else []

    # 查找质量承诺书页码
    quality_tasks = []
    if '3' in choices and survey_files_3:
        print("\n正在查找地籍调查报告中的'质量承诺书'页面...")
        for f in survey_files_3:
            page_num = find_page_by_keyword(f, "质量承诺书", occurrence=1)
            if page_num:
                print(f"   ✅ {os.path.basename(f)} - 质量承诺书在第 {page_num} 页")
                quality_tasks.append((f, page_num))
            else:
                print(f"   ❌ {os.path.basename(f)} - 未找到'质量承诺书'")

    # 查找验收报告页码
    acceptance_tasks = []
    if '4' in choices and survey_files_4:
        print("\n正在查找地籍调查报告中的'地籍调查成果验收报告'页面...")
        for f in survey_files_4:
            page_num = find_page_by_keyword(f, "地籍调查成果验收报告", occurrence=2)
            if page_num:
                print(f"   ✅ {os.path.basename(f)} - 质量承诺书在第 {page_num} 页")
                acceptance_tasks.append((f, page_num))
            else:
                print(f"   ❌ {os.path.basename(f)} - 未找到'质量承诺书'")

    # 构建任务列表
    tasks = []
    for f in excel_files:
        tasks.append({'file': f, 'type': 'excel', 'convert_fn': excel_to_pdf, 'extra': None, 'desc': '面积汇总表'})
    for f in word_files:
        tasks.append({'file': f, 'type': 'word', 'convert_fn': word_to_pdf, 'extra': None, 'desc': '审查申请表'})
    for f, page in quality_tasks:
        tasks.append({'file': f, 'type': 'quality', 'convert_fn': word_page_to_pdf, 'extra': page, 'desc': f'质量承诺书(第{page}页)'})
    for f, page in acceptance_tasks:
        tasks.append({'file': f, 'type': 'acceptance', 'convert_fn': word_page_to_pdf, 'extra': page, 'desc': f'验收报告(第{page}页)'})

    if not tasks:
        print("\nNo matching files found.")
        pause_and_exit(0)

    print(f"\nFound {len(tasks)} files to convert:")
    print("-" * 50)
    for i, task in enumerate(tasks, 1):
        print(f"  {i}. [{task['desc']}] {os.path.basename(task['file'])}")
    print("-" * 50)
    print()

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    ok_list = []
    fail_list = []

    for i, task in enumerate(tasks, 1):
        print(f"[{i}/{len(tasks)}] ", end="")
        stem = os.path.splitext(os.path.basename(task['file']))[0]

        if task['type'] == 'quality':
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}_质量承诺书.pdf")
            success = task['convert_fn'](task['file'], output_pdf, task['extra'], "质量承诺书")
        elif task['type'] == 'acceptance':
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}_地籍调查成果验收报告.pdf")
            success = task['convert_fn'](task['file'], output_pdf, task['extra'], "地籍调查成果验收报告")
        else:
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}.pdf")
            success = task['convert_fn'](task['file'], output_pdf)

        if success:
            ok_list.append(os.path.basename(task['file']))
        else:
            fail_list.append(os.path.basename(task['file']))

    print("=" * 50)
    print("  转换结果汇总")
    print("=" * 50)
    print(f"  成功：{len(ok_list)} 个")
    print(f"  失败：{len(fail_list)} 个")
    if fail_list:
        print("\n  失败的文件：")
        for f in fail_list:
            print(f"    - {f}")
    print()
    print(f"  PDF 文件保存位置：{OUTPUT_DIR}")
    print("=" * 50)

    pause_and_exit(0)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = f"错误类型：{type(e).__name__}\n错误信息：{str(e)}\n\n详细堆栈信息：\n{traceback.format_exc()}"
        print(f"\nProgram error: {e}")
        pause_and_exit(1, error_msg)
