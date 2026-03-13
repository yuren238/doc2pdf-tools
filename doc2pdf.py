# -*- coding: utf-8 -*-
"""
Document to PDF Converter (Windows Only)

Converts the following files to PDF:
- Excel files: Area summary tables (Sheet 1)
- Word documents: Application forms
- Word documents: Quality commitment page from survey reports
- Word documents: Acceptance report page from survey reports

Output: Grayscale PDF files saved to 'out' folder.

Author: Auto-generated
Version: 1.0.0
"""

import os
import sys
import glob
import subprocess
import winreg
import traceback
from datetime import datetime

# ============= Configuration =============
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(ROOT_DIR, 'out')
ERROR_LOG_FILE = os.path.join(ROOT_DIR, 'error.txt')
# ============= End Configuration =============


def write_error_log(error_message):
    """Write error details to error.txt file."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ERROR_LOG_FILE, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write("Document to PDF Converter - Error Log\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"Error Time: {timestamp}\n")
        f.write(f"Script Location: {ROOT_DIR}\n")
        f.write(f"Python Version: {sys.version}\n")
        f.write(f"OS: {sys.platform}\n")
        f.write("\n" + "-" * 60 + "\n")
        f.write("Error Details:\n")
        f.write("-" * 60 + "\n\n")
        f.write(error_message)
        f.write("\n\n" + "-" * 60 + "\n")
        f.write("Suggestions:\n")
        f.write("-" * 60 + "\n")
        f.write("1. Ensure Microsoft Office (Excel/Word) is installed\n")
        f.write("2. Ensure Python 3.7+ is installed\n")
        f.write("3. Try running: pip install pywin32 pypdf\n")
        f.write("4. Check if file is locked by another program\n")
        f.write("5. Verify file format (xls/xlsx/doc/docx)\n")
    print(f"\nError log saved to: {ERROR_LOG_FILE}")


def pause_and_exit(code=0, error_msg=None):
    """Pause before exit and write error log if needed."""
    if error_msg and code != 0:
        write_error_log(error_msg)
    print()
    os.system("pause")
    sys.exit(code)


# ─────────────────────────────────────────────
#  Environment Check & Auto-fix
# ─────────────────────────────────────────────

def check_python_version():
    """Check if Python version is 3.7+."""
    major, minor = sys.version_info[:2]
    if (major, minor) < (3, 7):
        print(f"[X] Python version too low: {major}.{minor}, requires 3.7+")
        print("    Download from https://www.python.org/downloads/")
        return False
    print(f"[OK] Python {major}.{minor}.{sys.version_info[2]}")
    return True


def check_pip_available():
    """Check if pip is available, try to fix if not."""
    try:
        subprocess.check_call(
            [sys.executable, '-m', 'pip', '--version'],
            stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT,
        )
        print("[OK] pip available")
        return True
    except subprocess.CalledProcessError:
        print("[!] pip not available, trying to fix...")
        try:
            subprocess.check_call(
                [sys.executable, '-m', 'ensurepip', '--upgrade'],
                stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT,
            )
            print("[OK] pip fixed")
            return True
        except Exception as e:
            print(f"[X] pip fix failed: {e}")
            print("    Please reinstall Python with 'Add pip' option.")
            return False


def check_and_install_pywin32():
    """Check and install pywin32 if needed."""
    try:
        import win32com.client
        import pythoncom
        print("[OK] pywin32 installed")
        return True
    except ImportError:
        print("[!] pywin32 not installed, installing...")
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32'])
            scripts_dir = os.path.join(os.path.dirname(sys.executable), 'Scripts')
            post_install = os.path.join(scripts_dir, 'pywin32_postinstall.py')
            if os.path.exists(post_install):
                subprocess.call(
                    [sys.executable, post_install, '-install'],
                    stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT,
                )
            print("[OK] pywin32 installed")
            return True
        except Exception as e:
            print(f"[X] pywin32 install failed: {e}")
            print("    Run manually: pip install pywin32")
            return False


def check_and_install_pypdf():
    """Check and install pypdf if needed."""
    try:
        import pypdf
        print("[OK] pypdf installed")
        return True
    except ImportError:
        print("[!] pypdf not installed, installing...")
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pypdf'])
            print("[OK] pypdf installed")
            return True
        except Exception as e:
            print(f"[X] pypdf install failed: {e}")
            print("    Run manually: pip install pypdf")
            return False


def check_office_via_registry(app_name):
    """Check if Office app is installed via registry."""
    reg_paths = [
        (winreg.HKEY_LOCAL_MACHINE,
         rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{app_name.upper()}.EXE"),
        (winreg.HKEY_LOCAL_MACHINE,
         rf"SOFTWARE\Classes\{app_name}.Application\CurVer"),
        (winreg.HKEY_LOCAL_MACHINE,
         rf"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\{app_name.upper()}.EXE"),
    ]
    for hive, path in reg_paths:
        try:
            with winreg.OpenKey(hive, path):
                return True
        except FileNotFoundError:
            continue
    return False


def check_office_installed(need_excel=True, need_word=True):
    """Check if required Office apps are installed."""
    all_ok = True
    if need_excel:
        if check_office_via_registry('Excel'):
            print("[OK] Microsoft Excel installed")
        else:
            print("[X] Microsoft Excel not detected, please install Office.")
            all_ok = False
    if need_word:
        if check_office_via_registry('Word'):
            print("[OK] Microsoft Word installed")
        else:
            print("[X] Microsoft Word not detected, please install Office.")
            all_ok = False
    return all_ok


def run_environment_check(need_excel, need_word):
    """Run all environment checks."""
    print("=" * 50)
    print("  Environment Check")
    print("=" * 50)
    results = [
        check_python_version(),
        check_pip_available(),
        check_and_install_pywin32(),
        check_and_install_pypdf(),
        check_office_installed(need_excel, need_word),
    ]
    all_ok = all(results)
    print("=" * 50)
    if all_ok:
        print("  Environment check passed, starting conversion...\n")
    else:
        print("  Environment check failed, please fix issues above.")
    return all_ok


# ─────────────────────────────────────────────
#  Grayscale Conversion
# ─────────────────────────────────────────────

def strip_color_from_stream(data: bytes) -> bytes:
    """
    Strip color commands from PDF content stream.
    Replace all color operators with grayscale equivalents.
    """
    import re

    # RGB fill: r g b rg -> 0 g
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg', b'0 g', data)
    # RGB stroke: r g b RG -> 0 G
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+RG', b'0 G', data)
    # CMYK fill -> 0 g
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+k', b'0 g', data)
    # CMYK stroke -> 0 G
    data = re.sub(rb'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+K', b'0 G', data)
    # scn/sc fill -> 0 g
    data = re.sub(rb'(?:[\d.]+\s+){1,4}sc(?:n)?', b'0 g', data)
    # SCN/SC stroke -> 0 G
    data = re.sub(rb'(?:[\d.]+\s+){1,4}SC(?:N)?', b'0 G', data)
    return data


def convert_pdf_to_grayscale(input_pdf, output_pdf):
    """Convert color PDF to grayscale."""
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
#  Menu
# ─────────────────────────────────────────────

def show_menu():
    """Display menu and return user selection."""
    print("\n" + "=" * 50)
    print("  Document to PDF Converter")
    print("=" * 50)
    print("  Select file type to convert:")
    print()
    print("  1. Area Summary Table (Excel, Sheet 1)")
    print("  2. Application Form (Word document)")
    print("  3. Quality Commitment Page (from Survey Report)")
    print("  4. Acceptance Report Page (from Survey Report)")
    print()
    print("  Multi-select supported:")
    print("    1,2,3,4  -> Convert all")
    print("    34       -> Convert options 3 and 4")
    print("    Enter    -> Convert all (default)")
    print("    0        -> Exit")
    print()

    while True:
        try:
            choice = input("  Enter selection: ").strip()
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

        print("  Invalid input. Enter 1, 2, 3, 4 or combinations like 1234, 1,2,3,4")


# ─────────────────────────────────────────────
#  Conversion Functions
# ─────────────────────────────────────────────

def excel_to_pdf(excel_file, output_pdf):
    """Convert Excel Sheet 1 to grayscale PDF."""
    filename = os.path.basename(excel_file)
    print(f"> Processing Excel: {filename}")

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
        print(f"   Sheet name: {ws.Name}")

        tmp_pdf = output_pdf + ".tmp.pdf"
        ws.ExportAsFixedFormat(
            Type=0,
            Filename=os.path.abspath(tmp_pdf),
            Quality=0,
            IncludeDocProperties=False,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )

        print("   Converting to grayscale...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   [OK] Saved: {output_pdf}\n")
        return True

    except Exception as e:
        print(f"   [FAILED] {e}\n")
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
    """Convert Word document to grayscale PDF."""
    filename = os.path.basename(word_file)
    print(f"> Processing Word: {filename}")

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

        print("   Converting to grayscale...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   [OK] Saved: {output_pdf}\n")
        return True

    except Exception as e:
        print(f"   [FAILED] {e}\n")
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
    """Convert specific page of Word document to grayscale PDF."""
    filename = os.path.basename(word_file)
    print(f"> Processing Word: {filename} ({page_name} - Page {page_num})")

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

        print("   Converting to grayscale...")
        convert_pdf_to_grayscale(tmp_pdf, output_pdf)
        os.remove(tmp_pdf)

        print(f"   [OK] Saved: {output_pdf}\n")
        return True

    except Exception as e:
        print(f"   [FAILED] {e}\n")
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
    Find page number by searching keyword in Word document.
    
    Args:
        word_file: Path to Word document
        keyword: Text to search for
        occurrence: Which occurrence to find (default: 1)
    
    Returns:
        Page number (int) or None if not found
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
                    print(f"   File open failed, retrying... ({attempt + 1}/{max_retries})")
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
        print(f"   Error finding page: {e}")
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
#  Main
# ─────────────────────────────────────────────

def main():
    os.chdir(ROOT_DIR)

    choices = show_menu()
    if '0' in choices:
        print("\nProgram exited.")
        pause_and_exit(0)

    need_excel = '1' in choices
    need_word = '2' in choices or '3' in choices or '4' in choices

    if not run_environment_check(need_excel, need_word):
        pause_and_exit(1, "Environment check failed")

    def find_files(pattern):
        return [f for f in glob.glob(pattern, recursive=True)
                if not os.path.basename(f).startswith('~')]

    # File patterns (adjust these for your needs)
    excel_files = find_files(os.path.join(ROOT_DIR, '**', '*面积汇总表*.xls*')) if need_excel else []
    word_files = find_files(os.path.join(ROOT_DIR, '**', '*审查申请表*.doc*')) if '2' in choices else []
    survey_files_3 = find_files(os.path.join(ROOT_DIR, '**', '*地籍调查报告*.doc*')) if '3' in choices else []
    survey_files_4 = find_files(os.path.join(ROOT_DIR, '**', '*地籍调查报告*.doc*')) if '4' in choices else []

    # Find Quality Commitment pages
    quality_tasks = []
    if '3' in choices and survey_files_3:
        print("\nSearching for 'Quality Commitment' pages...")
        for f in survey_files_3:
            page_num = find_page_by_keyword(f, "质量承诺书", occurrence=1)
            if page_num:
                print(f"   [OK] {os.path.basename(f)} - Found at page {page_num}")
                quality_tasks.append((f, page_num))
            else:
                print(f"   [X] {os.path.basename(f)} - Not found")

    # Find Acceptance Report pages (2nd occurrence)
    acceptance_tasks = []
    if '4' in choices and survey_files_4:
        print("\nSearching for 'Acceptance Report' pages...")
        for f in survey_files_4:
            page_num = find_page_by_keyword(f, "地籍调查成果验收报告", occurrence=2)
            if page_num:
                print(f"   [OK] {os.path.basename(f)} - Found at page {page_num}")
                acceptance_tasks.append((f, page_num))
            else:
                print(f"   [X] {os.path.basename(f)} - Not found")

    # Build task list
    tasks = []
    for f in excel_files:
        tasks.append({'file': f, 'type': 'excel', 'convert_fn': excel_to_pdf, 'extra': None, 'desc': 'Area Summary'})
    for f in word_files:
        tasks.append({'file': f, 'type': 'word', 'convert_fn': word_to_pdf, 'extra': None, 'desc': 'Application Form'})
    for f, page in quality_tasks:
        tasks.append({'file': f, 'type': 'quality', 'convert_fn': word_page_to_pdf, 'extra': page, 'desc': f'Quality Commitment(Page {page})'})
    for f, page in acceptance_tasks:
        tasks.append({'file': f, 'type': 'acceptance', 'convert_fn': word_page_to_pdf, 'extra': page, 'desc': f'Acceptance Report(Page {page})'})

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
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}_quality_commitment.pdf")
            success = task['convert_fn'](task['file'], output_pdf, task['extra'], "Quality Commitment")
        elif task['type'] == 'acceptance':
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}_acceptance_report.pdf")
            success = task['convert_fn'](task['file'], output_pdf, task['extra'], "Acceptance Report")
        else:
            output_pdf = os.path.join(OUTPUT_DIR, f"{stem}.pdf")
            success = task['convert_fn'](task['file'], output_pdf)

        if success:
            ok_list.append(os.path.basename(task['file']))
        else:
            fail_list.append(os.path.basename(task['file']))

    print("=" * 50)
    print("  Conversion Summary")
    print("=" * 50)
    print(f"  Success: {len(ok_list)}")
    print(f"  Failed: {len(fail_list)}")
    if fail_list:
        print("\n  Failed files:")
        for f in fail_list:
            print(f"    - {f}")
    print()
    print(f"  Output folder: {OUTPUT_DIR}")
    print("=" * 50)

    pause_and_exit(0)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = f"Error Type: {type(e).__name__}\nError: {str(e)}\n\nStack Trace:\n{traceback.format_exc()}"
        print(f"\nProgram error: {e}")
        pause_and_exit(1, error_msg)
