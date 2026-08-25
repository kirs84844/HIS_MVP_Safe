import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pyperclip
import ctypes
import ctypes.wintypes
import threading
import time
import calendar
from datetime import datetime, timedelta

# [核心防御] 强制系统 DPI 感知
try:
    ctypes.windll.user32.SetProcessDPIAware()
except AttributeError:
    pass

# ==================== RPA 物理锚点配置 (V5.8) ====================
RPA_CONFIG = {
    "BTN_SWITCH_PATIENT": (32, 96),      
    "PATIENT_FIRST_ROW": (86, 215),      
    "LINE_HEIGHT": 18,                   
    
    "READY_PIXEL_POS": (28, 95),         
    "READY_PIXEL_RGB": (245, 245, 245),  
    "AREA_SAFE_BLANK": (500, 642),       
    
    # --- V5.8 核心修改锚点 ---
    "AREA_PROGRESS_RECORD": (60, 267),   # 1. 预定焦点区 (先单击此处)
    "BTN_NEW_RECORD": (240, 80),         # 2. 新建按钮 (再单击此处)
    "TPL_OPTION": (825, 370),            # 3. 模板列表具体位置
}

DB_FILE = "his_data.db"
APP_VERSION = "6.4"
is_running_auto = False
is_paused_auto = False
stop_requested = False
locked_tpl_name = ""
HOTKEY_PAUSE_ID = 1
HOTKEY_STOP_ID = 2
VK_F8 = 0x77
VK_F9 = 0x78
WM_HOTKEY = 0x0312
MOD_NOREPEAT = 0x4000
DEFAULT_SUPPLEMENT_TITLE = "高一鸣主治医师查房记录"
SCOPE_ALL = "全部"
SCOPE_INCLUDE = "指定范围"
SCOPE_EXCLUDE = "跳过床位"
INTERVAL_DAYS = {"7天": 7, "14天": 14}
INTERVAL_MONTHS = {"1个月": 1, "2个月": 2, "3个月": 3}
WORD_TEMPLATE_NAME = "中医会诊单"
WORD_DEPARTMENT = "浦二"
WORD_FIXED_ADMIT_DATE = "2022.9.20"
WORD_PAGE_BREAK = "\f"
DEFAULT_WORD_TEMPLATE = """会诊请求
拟请科室：  中医科                     2026 年  6月 30 日
患者姓名：{{name}}   性别：{{gender}}  年龄：{{age}}  科室：{{department}}   床号：{{bed}}
诊断：{{diagnosis}}
病情摘要：患者主因 “{{complaint}}”于{{admit_date}}入院，目前精神检查：意识清，定向可，仪态尚整，注意力集中，接触合作，对答切题，未引出明显幻觉、妄想，思维贫乏，情感反应淡漠，意志要求减退，智能可，自知力无。
目前体格检查：暂无阳性症状。
会诊目的：请中医科协助治疗。

请求科室：{{department}}    医师：李晗       主治医师：高一鸣
                                                                 
会诊意见:

                         
"""

# ==================== 1. 硬件模拟引擎 ====================
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

def mouse_click(x, y):
    user32.SetCursorPos(x, y)
    user32.mouse_event(2, 0, 0, 0, 0)
    user32.mouse_event(4, 0, 0, 0, 0)

def mouse_double_click(x, y):
    mouse_click(x, y)
    time.sleep(0.1)
    mouse_click(x, y)

def get_pixel_color(x, y):
    hdc = user32.GetDC(0)
    pixel = gdi32.GetPixel(hdc, x, y)
    user32.ReleaseDC(0, hdc)
    r = pixel & 0x0000ff
    g = (pixel & 0x00ff00) >> 8
    b = (pixel & 0xff0000) >> 16
    return (r, g, b)

def press_key(vk_code):
    user32.keybd_event(vk_code, 0, 0, 0)
    user32.keybd_event(vk_code, 0, 2, 0)

def paste_text(text):
    pyperclip.copy(text)
    time.sleep(0.2)
    user32.keybd_event(0x11, 0, 0, 0)
    press_key(0x56)
    user32.keybd_event(0x11, 0, 2, 0)

def normalize_bed(value):
    bed = clean_text(value).strip()
    return str(int(bed)) if bed.isdigit() else bed

def parse_bed_spec(spec):
    beds = set()
    text = spec.replace("，", ",").strip()
    if not text:
        raise ValueError("请输入床位范围。")
    for item in text.split(","):
        item = item.strip()
        if not item:
            raise ValueError("床位范围中存在空项。")
        if "-" in item:
            parts = [part.strip() for part in item.split("-")]
            if len(parts) != 2 or not all(part.isdigit() for part in parts):
                raise ValueError(f"无法识别床位范围: {item}")
            start, end = int(parts[0]), int(parts[1])
            if start < 1 or end < start:
                raise ValueError(f"床位范围无效: {item}")
            beds.update(str(bed) for bed in range(start, end + 1))
        elif item.isdigit() and int(item) >= 1:
            beds.add(str(int(item)))
        else:
            raise ValueError(f"无法识别床位: {item}")
    return beds

def select_scope_patients(patients, mode, spec):
    if mode == SCOPE_ALL:
        return patients
    selected_beds = parse_bed_spec(spec)
    if mode == SCOPE_INCLUDE:
        return [patient for patient in patients if normalize_bed(patient[0]) in selected_beds]
    if mode == SCOPE_EXCLUDE:
        return [patient for patient in patients if normalize_bed(patient[0]) not in selected_beds]
    raise ValueError("未知的患者范围模式。")

def add_calendar_months(start_time, months):
    month_index = start_time.year * 12 + start_time.month - 1 + months
    year, zero_based_month = divmod(month_index, 12)
    month = zero_based_month + 1
    day = min(start_time.day, calendar.monthrange(year, month)[1])
    return start_time.replace(year=year, month=month, day=day)

def build_supplement_times(start_time, interval_name, end_time):
    times = []
    step = 0
    while True:
        if interval_name in INTERVAL_DAYS:
            current = start_time + timedelta(days=INTERVAL_DAYS[interval_name] * step)
        elif interval_name in INTERVAL_MONTHS:
            current = add_calendar_months(start_time, INTERVAL_MONTHS[interval_name] * step)
        else:
            raise ValueError("未知的补病史时间间隔。")
        if current > end_time:
            break
        times.append(current)
        step += 1
    return times

# ==================== 2. 数据库模块 ====================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS patients (
                        bed TEXT PRIMARY KEY, name TEXT, gender TEXT, 
                        age TEXT, admit_date TEXT, complaint TEXT, 
                        admit_diag TEXT, current_diag TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS templates (
                        name TEXT PRIMARY KEY, content TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS word_templates (
                        name TEXT PRIMARY KEY, content TEXT)''')
    cursor.execute("INSERT OR IGNORE INTO word_templates VALUES (?,?)", (WORD_TEMPLATE_NAME, DEFAULT_WORD_TEMPLATE))
    conn.commit()
    conn.close()

def refresh_all_data():
    for row in mgr_tree.get_children(): mgr_tree.delete(row)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM patients")
    for row in cursor.fetchall():
        mgr_tree.insert("", "end", values=row)
    tpl_listbox.delete(0, tk.END)
    mgr_tpl_listbox.delete(0, tk.END)
    cursor.execute("SELECT name FROM templates")
    template_names = [row[0] for row in cursor.fetchall()]
    for template_name in template_names:
        tpl_listbox.insert(tk.END, template_name)
        mgr_tpl_listbox.insert(tk.END, template_name)
    if 'supp_template_combos' in globals():
        for combo in supp_template_combos:
            combo['values'] = [""] + template_names
    conn.close()

def load_all_patients():
    conn = sqlite3.connect(DB_FILE)
    patients = conn.execute("SELECT * FROM patients").fetchall()
    conn.close()
    return sorted(patients, key=patient_sort_key)

# ==================== 3. 核心全自动逻辑引擎 (V6.4) ====================
def status_update(msg, color=None):
    def apply_update():
        status_label.config(text=msg)
        if color:
            status_label.config(fg=color)
    if threading.current_thread() is threading.main_thread():
        apply_update()
    else:
        root.after(0, apply_update)

def wait_until_resumed():
    while is_paused_auto and not stop_requested:
        time.sleep(0.1)
    return not stop_requested

def find_patient_from_window(patients):
    hwnd = user32.GetForegroundWindow()
    length = user32.GetWindowTextLengthW(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buff, length + 1)
    window_title = buff.value
    matches = [patient for patient in patients if clean_text(patient[1]) and clean_text(patient[1]) in window_title]
    return matches[0] if len(matches) == 1 else None

def open_patient_row(row_number, total_rows, patients):
    if not wait_until_resumed():
        return None
    status_update(f"正在扫描第 {row_number} 行 (进度: {row_number}/{total_rows})...")
    mouse_click(*RPA_CONFIG["BTN_SWITCH_PATIENT"])
    time.sleep(0.2)
    target_y = RPA_CONFIG["PATIENT_FIRST_ROW"][1] + ((row_number - 1) * RPA_CONFIG["LINE_HEIGHT"])
    mouse_double_click(RPA_CONFIG["PATIENT_FIRST_ROW"][0], target_y)

    status_update(f"第 {row_number} 行: 等待复苏...")
    busy_check = 0.0
    while busy_check < 5.0 and not stop_requested:
        if get_pixel_color(*RPA_CONFIG["READY_PIXEL_POS"]) != RPA_CONFIG["READY_PIXEL_RGB"]:
            break
        time.sleep(0.2)
        busy_check += 0.2
    while not stop_requested:
        if get_pixel_color(*RPA_CONFIG["READY_PIXEL_POS"]) == RPA_CONFIG["READY_PIXEL_RGB"]:
            time.sleep(0.8)
            break
        time.sleep(0.3)
    if stop_requested:
        return None

    mouse_click(*RPA_CONFIG["AREA_SAFE_BLANK"])
    time.sleep(0.5)
    return find_patient_from_window(patients)

def locate_target_patient(target_patient, patients, total_rows):
    target_bed = normalize_bed(target_patient[0])
    if target_bed.isdigit():
        row_number = max(1, int(target_bed))
    else:
        row_number = 1

    visited_rows = set()
    current_patient = open_patient_row(row_number, total_rows, patients)
    visited_rows.add(row_number)
    if current_patient and normalize_bed(current_patient[0]) == target_bed:
        return current_patient

    current_bed = normalize_bed(current_patient[0]) if current_patient else ""
    if current_bed.isdigit() and target_bed.isdigit() and abs(int(target_bed) - int(current_bed)) > 3:
        correction = int(target_bed) - int(current_bed)
        row_number = max(1, min(total_rows, row_number + correction))
        if row_number not in visited_rows:
            current_patient = open_patient_row(row_number, total_rows, patients)
            visited_rows.add(row_number)
            if current_patient and normalize_bed(current_patient[0]) == target_bed:
                return current_patient

    for _ in range(12):
        if stop_requested:
            return None
        current_bed = normalize_bed(current_patient[0]) if current_patient else ""
        if current_bed.isdigit() and target_bed.isdigit():
            direction = 1 if int(target_bed) > int(current_bed) else -1
        else:
            direction = 1
        next_row = row_number + direction
        if next_row < 1 or next_row > total_rows or next_row in visited_rows:
            return None
        row_number = next_row
        current_patient = open_patient_row(row_number, total_rows, patients)
        visited_rows.add(row_number)
        if current_patient and normalize_bed(current_patient[0]) == target_bed:
            return current_patient
    return None

def render_patient_template(template, patient):
    values = {
        "name": clean_text(patient[1]),
        "gender": clean_text(patient[2]),
        "age": clean_text(patient[3]),
        "admit_date": clean_text(patient[4]),
        "complaint": clean_text(patient[5]),
        "admit_diag": clean_text(patient[6]),
        "current_diag": clean_text(patient[7]),
    }
    for key, value in values.items():
        template = template.replace("{{" + key + "}}", value)
    return template

def create_his_record(patient, template, record_time=None, record_title=None, replace_patient_fields=True):
    mouse_click(*RPA_CONFIG["AREA_PROGRESS_RECORD"])
    time.sleep(0.3)
    mouse_click(*RPA_CONFIG["BTN_NEW_RECORD"])
    time.sleep(1.2)

    if record_time:
        time_str = record_time.strftime("%y-%m-%d %H:%M")
        status_update(f"覆写时间: {time_str}")
        mouse_click(*RPA_CONFIG["TPL_OPTION"])
        time.sleep(0.5)
        press_key(0x09)
        time.sleep(0.1)
        press_key(0x09)
        time.sleep(0.3)
        paste_text(time_str)
        time.sleep(0.3)

        if record_title:
            press_key(0x09)
            time.sleep(0.2)
            paste_text(record_title)
            time.sleep(0.3)
            press_key(0x09)
        else:
            press_key(0x09)
            time.sleep(0.1)
            press_key(0x09)
        time.sleep(0.3)
        press_key(0x0D)
        time.sleep(1.2)
        press_key(0x0D)
        time.sleep(0.2)
    else:
        mouse_double_click(*RPA_CONFIG["TPL_OPTION"])
        time.sleep(1.2)

    final_text = render_patient_template(template, patient) if replace_patient_fields else template
    paste_text(final_text)
    time.sleep(0.8)
    user32.keybd_event(0x11, 0, 0, 0)
    press_key(0x53)
    user32.keybd_event(0x11, 0, 2, 0)
    time.sleep(2.5)

def finish_automation(message):
    global is_running_auto, is_paused_auto, stop_requested
    is_running_auto = False
    is_paused_auto = False
    stop_requested = False
    def finish_ui():
        root.deiconify()
        status_label.config(text="--- 任务结束 ---", fg="green")
        messagebox.showinfo("任务完成", message)
    root.after(0, finish_ui)

def run_automation_worker(target, args):
    try:
        target(*args)
    except Exception as exc:
        finish_automation(f"执行过程中发生异常，任务已停止。\n{exc}")

def start_automation_flow(patients, target_patients, target_time_obj, template):
    target_beds = {normalize_bed(patient[0]) for patient in target_patients}
    processed_beds = set()
    record_index = 0
    ordered_targets = sorted(target_patients, key=patient_sort_key)
    for target_patient in ordered_targets:
        if stop_requested or processed_beds == target_beds:
            break
        patient = locate_target_patient(target_patient, patients, len(patients))
        if stop_requested:
            break
        if not patient:
            status_update(f"【跳过】{target_patient[0]}床未定位成功")
            continue
        bed = normalize_bed(patient[0])
        if bed not in target_beds or bed in processed_beds:
            status_update(f"【跳过】{patient[0]}床 {patient[1]}，校验床位不在目标范围")
            continue
        if not wait_until_resumed():
            break
        record_time = target_time_obj + timedelta(minutes=record_index) if target_time_obj else None
        create_his_record(patient, template, record_time)
        processed_beds.add(bed)
        record_index += 1
        status_update(f"【成功】已归档: {patient[1]}")

    missing = sorted(target_beds - processed_beds, key=lambda bed: int(bed) if bed.isdigit() else 9999)
    summary = f"已处理 {len(processed_beds)} 位患者。"
    if missing:
        summary += "\n未找到或未处理床位: " + ", ".join(missing)
    if stop_requested:
        summary += "\n任务已按 F9 安全终止。"
    finish_automation(summary)

def start_supplement_flow(record_times, templates, record_title):
    time.sleep(0.5)
    created_count = 0
    for record_index, record_time in enumerate(record_times):
        if not wait_until_resumed():
            break
        content = templates[record_index % len(templates)]
        status_update(f"当前 HIS 患者: 第 {record_index + 1}/{len(record_times)} 条")
        create_his_record(None, content, record_time, record_title, replace_patient_fields=False)
        created_count += 1
        if stop_requested:
            break

    summary = f"当前 HIS 患者：已创建 {created_count} 条补充病史。"
    if stop_requested:
        summary += "\n任务已按 F9 安全终止。"
    finish_automation(summary)

# ==================== 4. UI 交互层 ====================
def toggle_pause_auto():
    global is_paused_auto
    if not is_running_auto:
        return
    is_paused_auto = not is_paused_auto
    status_update("【已暂停】再次按 F8 继续" if is_paused_auto else "【继续运行】", "orange" if is_paused_auto else "green")

def stop_auto():
    global stop_requested, is_paused_auto
    if not is_running_auto:
        root.deiconify()
        return
    stop_requested = True
    is_paused_auto = False
    status_update("【安全终止】当前原子动作完成后停止...", "red")

def global_hotkey_loop():
    pause_ok = user32.RegisterHotKey(None, HOTKEY_PAUSE_ID, MOD_NOREPEAT, VK_F8)
    stop_ok = user32.RegisterHotKey(None, HOTKEY_STOP_ID, MOD_NOREPEAT, VK_F9)
    if not pause_ok or not stop_ok:
        root.after(0, lambda: status_update("警告：F8/F9 全局快捷键注册失败。", "red"))
    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
        if msg.message == WM_HOTKEY and msg.wParam == HOTKEY_PAUSE_ID:
            root.after(0, toggle_pause_auto)
        elif msg.message == WM_HOTKEY and msg.wParam == HOTKEY_STOP_ID:
            root.after(0, stop_auto)

def start_global_hotkeys():
    threading.Thread(target=global_hotkey_loop, daemon=True).start()

def toggle_time_entry():
    if time_var.get():
        time_entry.config(state=tk.NORMAL)
    else:
        time_entry.config(state=tk.DISABLED)

def toggle_scope_entry():
    state = tk.DISABLED if scope_var.get() == SCOPE_ALL else tk.NORMAL
    for entry in scope_entries:
        entry.config(state=state)

def run_thread():
    global is_running_auto, is_paused_auto, stop_requested
    try:
        if is_running_auto:
            messagebox.showwarning("任务运行中", "请先完成或终止当前任务。")
            return
        if not locked_tpl_name:
            messagebox.showwarning("拦截", "请先锁定模板！")
            return
        patients = load_all_patients()
        target_patients = select_scope_patients(patients, scope_var.get(), scope_text_var.get())
        if not patients or not target_patients:
            messagebox.showwarning("无患者", "本地患者库或所选患者范围为空。")
            return
        conn = sqlite3.connect(DB_FILE)
        row = conn.execute("SELECT content FROM templates WHERE name=?", (locked_tpl_name,)).fetchone()
        conn.close()
        if not row:
            messagebox.showwarning("模板缺失", "锁定的模板不存在。")
            return
        target_time_obj = None
        if time_var.get():
            time_str = time_entry.get().strip()
            try:
                target_time_obj = datetime.strptime(time_str, "%y-%m-%d %H:%M")
            except ValueError:
                messagebox.showerror("格式阻断", "时间格式存在瑕疵！\n请严格遵循: YY-MM-DD HH:MM\n(示例: 26-04-26 10:00)")
                return
        is_running_auto = True
        is_paused_auto = False
        stop_requested = False
        status_update("程序将自动隐藏，3 秒后接管 HIS...")
        root.update()
        time.sleep(3)
        root.iconify()
        threading.Thread(
            target=run_automation_worker,
            args=(start_automation_flow, (patients, target_patients, target_time_obj, row[0])),
            daemon=True
        ).start()
    except ValueError:
        messagebox.showerror("参数错误", "患者范围格式无效，请使用例如 1-8,13-17。")

def on_tpl_select(event):
    global locked_tpl_name
    selected = tpl_listbox.curselection()
    if selected:
        locked_tpl_name = tpl_listbox.get(selected)
        if lock_var.get():
            status_label.config(text=f"状态: 挂载 <{locked_tpl_name}>", fg="blue")

def toggle_lock():
    global locked_tpl_name
    selected = tpl_listbox.curselection()
    if lock_var.get():
        if not selected:
            lock_var.set(False)
        else:
            locked_tpl_name = tpl_listbox.get(selected)
            status_label.config(text=f"状态: 挂载 <{locked_tpl_name}>", fg="blue")
    else:
        status_label.config(text="状态: 挂载解除", fg="gray")

# 数据库 CRUD 
def save_patient():
    data = (p_bed.get(), p_name.get(), p_gender.get(), p_age.get(), p_admit.get(), p_comp.get(), p_adiag.get(), p_cdiag.get())
    if not data[0] or not data[1]: return
    conn = sqlite3.connect(DB_FILE)
    conn.execute("REPLACE INTO patients VALUES (?,?,?,?,?,?,?,?)", data); conn.commit(); conn.close()
    refresh_all_data()

def delete_patient():
    selected = mgr_tree.selection()
    if not selected: return
    bed_num = mgr_tree.item(selected)['values'][0]
    conn = sqlite3.connect(DB_FILE); conn.execute("DELETE FROM patients WHERE bed=?", (bed_num,)); conn.commit(); conn.close()
    refresh_all_data()

def on_patient_select(event):
    selected = mgr_tree.selection()
    if not selected: return
    vals = mgr_tree.item(selected)['values']
    entries = [p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag]
    for i, entry in enumerate(entries):
        entry.delete(0, tk.END); entry.insert(0, str(vals[i]) if vals[i] != 'None' else "")

def save_template():
    t_name = tpl_name_entry.get(); t_content = tpl_content_text.get("1.0", tk.END).strip()
    if not t_name or not t_content: return
    conn = sqlite3.connect(DB_FILE); conn.execute("REPLACE INTO templates VALUES (?,?)", (t_name, t_content)); conn.commit(); conn.close()
    refresh_all_data()

def delete_template():
    selected = mgr_tpl_listbox.curselection()
    if not selected: return
    conn = sqlite3.connect(DB_FILE); conn.execute("DELETE FROM templates WHERE name=?", (mgr_tpl_listbox.get(selected),)); conn.commit(); conn.close()
    refresh_all_data()

def on_template_select(event):
    selected = mgr_tpl_listbox.curselection()
    if not selected: return
    conn = sqlite3.connect(DB_FILE); content = conn.execute("SELECT content FROM templates WHERE name=?", (mgr_tpl_listbox.get(selected),)).fetchone()[0]; conn.close()
    tpl_name_entry.delete(0, tk.END); tpl_name_entry.insert(0, mgr_tpl_listbox.get(selected))
    tpl_content_text.delete("1.0", tk.END); tpl_content_text.insert(tk.END, content)

def toggle_supplement_title():
    supp_title_entry.config(state=tk.DISABLED if supp_blank_title_var.get() else tk.NORMAL)

def run_supplement_thread():
    global is_running_auto, is_paused_auto, stop_requested
    if is_running_auto:
        messagebox.showwarning("任务运行中", "请先完成或终止当前任务。")
        return
    try:
        start_date = datetime.strptime(supp_start_date_entry.get().strip(), "%Y-%m-%d").date()
        start_clock = datetime.strptime(supp_time_entry.get().strip(), "%H:%M").time()
        start_time = datetime.combine(start_date, start_clock)
        record_times = build_supplement_times(start_time, supp_interval_var.get(), datetime.now())
    except ValueError as exc:
        messagebox.showerror("参数错误", str(exc))
        return
    if not record_times:
        messagebox.showwarning("时间范围无效", "起始时间晚于当前系统时间，没有可创建的病史。")
        return

    template_names = [var.get().strip() for var in supp_template_vars]
    if not template_names[0]:
        messagebox.showwarning("模板缺失", "模板1必须选择。")
        return
    if bool(template_names[1]) != bool(template_names[2]):
        messagebox.showwarning("模板组合不完整", "请只选择模板1，或同时选择模板1、2、3。")
        return
    selected_names = template_names if template_names[1] else template_names[:1]
    if len(selected_names) == 3 and len(set(selected_names)) != 3:
        messagebox.showwarning("模板重复", "循环使用的3个模板必须互不相同。")
        return
    conn = sqlite3.connect(DB_FILE)
    templates = []
    for template_name in selected_names:
        row = conn.execute("SELECT content FROM templates WHERE name=?", (template_name,)).fetchone()
        if row:
            templates.append(row[0])
    conn.close()
    if len(templates) != len(selected_names):
        messagebox.showwarning("模板缺失", "所选模板已被删除，请重新选择。")
        return

    record_title = "" if supp_blank_title_var.get() else supp_title_entry.get().strip()
    if not supp_blank_title_var.get() and not record_title:
        messagebox.showwarning("病史名称为空", "请输入病史名称，或勾选“病史名称留空”。")
        return

    if not messagebox.askyesno(
        "确认补病史",
        f"将为当前 HIS 患者创建 {len(record_times)} 条病史，使用 {len(templates)} 个模板。\n"
        f"起始 {start_time:%Y-%m-%d %H:%M}，间隔 {supp_interval_var.get()}。\n\n确认开始吗？"
    ):
        return

    is_running_auto = True
    is_paused_auto = False
    stop_requested = False
    status_update("程序将自动隐藏，3 秒后接管 HIS...")
    root.update()
    time.sleep(3)
    root.iconify()
    threading.Thread(
        target=run_automation_worker,
        args=(start_supplement_flow, (record_times, templates, record_title)),
        daemon=True
    ).start()

def save_word_template():
    content = word_template_text.get("1.0", tk.END).strip()
    if not content:
        messagebox.showwarning("拦截", "会诊单模板不能为空。")
        return
    conn = sqlite3.connect(DB_FILE)
    conn.execute("REPLACE INTO word_templates VALUES (?,?)", (WORD_TEMPLATE_NAME, content))
    conn.commit()
    conn.close()
    word_status_label.config(text="会诊单模板已保存。", fg="blue")

def load_word_template():
    conn = sqlite3.connect(DB_FILE)
    row = conn.execute("SELECT content FROM word_templates WHERE name=?", (WORD_TEMPLATE_NAME,)).fetchone()
    conn.close()
    word_template_text.delete("1.0", tk.END)
    word_template_text.insert(tk.END, row[0] if row else DEFAULT_WORD_TEMPLATE)

def patient_sort_key(row):
    bed = str(row[0])
    return (int(bed) if bed.isdigit() else 9999, bed)

def clean_text(value):
    if value is None or value == "None":
        return ""
    return str(value)

def fetch_word_patients(start_bed, count):
    conn = sqlite3.connect(DB_FILE)
    patients = conn.execute("SELECT * FROM patients").fetchall()
    conn.close()
    patients = sorted(patients, key=patient_sort_key)
    if start_bed:
        patients = [p for p in patients if patient_sort_key(p) >= patient_sort_key((start_bed,))]
    if count:
        patients = patients[:count]
    return patients

def render_word_page(template, patient):
    values = {
        "bed": clean_text(patient[0]),
        "name": clean_text(patient[1]),
        "gender": clean_text(patient[2]),
        "age": clean_text(patient[3]),
        "department": WORD_DEPARTMENT,
        "diagnosis": clean_text(patient[7]),
        "complaint": clean_text(patient[5]),
        "admit_date": WORD_FIXED_ADMIT_DATE,
    }
    for key, value in values.items():
        template = template.replace("{{" + key + "}}", value)
    return template.strip()

def copy_word_consult_text(event=None):
    try:
        start_bed = word_start_entry.get().strip()
        count_text = word_count_entry.get().strip()
        count = int(count_text) if count_text else None
        if count is not None and count < 1:
            messagebox.showerror("参数异常", "生成人数必须为空或 ≥1。")
            return
    except ValueError:
        messagebox.showerror("参数错误", "生成人数请输入有效数字，或留空。")
        return

    template = word_template_text.get("1.0", tk.END).strip()
    if not template:
        messagebox.showwarning("拦截", "会诊单模板不能为空。")
        return

    patients = fetch_word_patients(start_bed, count)
    if not patients:
        messagebox.showwarning("无数据", "未找到符合床号条件的患者。")
        return

    pages = [render_word_page(template, patient) for patient in patients]
    pyperclip.copy(WORD_PAGE_BREAK.join(pages))
    word_status_label.config(text=f"已复制 {len(pages)} 页会诊单。请切换到 Word 97-2003 后粘贴。", fg="green")

def setup_ui():
    global tpl_listbox, status_label, lock_var, root, mgr_tree, mgr_tpl_listbox, p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag, tpl_name_entry, tpl_content_text
    global word_start_entry, word_count_entry, word_template_text, word_status_label
    global time_var, time_entry, scope_var, scope_text_var, scope_entries
    global supp_template_vars, supp_template_combos, supp_start_date_entry, supp_time_entry, supp_interval_var
    global supp_title_entry, supp_blank_title_var
    
    root = tk.Tk(); root.title(f"极速精神科工作站 V{APP_VERSION}"); root.geometry("900x650")
    nb = ttk.Notebook(root); nb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    tab1 = ttk.Frame(nb); nb.add(tab1, text="🚀 极速引擎")
    cf = tk.Frame(tab1); cf.pack(fill=tk.BOTH, expand=True, padx=40, pady=20)
    tpl_listbox = tk.Listbox(cf, height=5, font=("Arial", 11)); tpl_listbox.pack(fill=tk.X); tpl_listbox.bind('<<ListboxSelect>>', on_tpl_select)
    lock_var = tk.BooleanVar(); tk.Checkbutton(cf, text="🔒 锁定模板", variable=lock_var, command=toggle_lock).pack(pady=5)
    
    # 时间劫持面板
    time_fm = tk.Frame(cf); time_fm.pack(pady=5)
    time_var = tk.BooleanVar()
    tk.Checkbutton(time_fm, text="启用历史时间覆写 (自动+1分步进)", variable=time_var, command=toggle_time_entry).pack(side=tk.LEFT)
    time_entry = tk.Entry(time_fm, width=16, justify='center')
    time_entry.insert(0, "26-04-26 10:00")
    time_entry.pack(side=tk.LEFT, padx=5)
    time_entry.config(state=tk.DISABLED) 
    
    scope_var = tk.StringVar(value=SCOPE_ALL)
    scope_text_var = tk.StringVar()
    scope_entries = []
    scope_fm = tk.LabelFrame(cf, text="患者范围"); scope_fm.pack(fill=tk.X, pady=8)
    for mode in (SCOPE_ALL, SCOPE_INCLUDE, SCOPE_EXCLUDE):
        tk.Radiobutton(scope_fm, text=mode, value=mode, variable=scope_var, command=toggle_scope_entry).pack(side=tk.LEFT, padx=8)
    scope_entry = tk.Entry(scope_fm, width=24, textvariable=scope_text_var, state=tk.DISABLED)
    scope_entry.pack(side=tk.LEFT, padx=8); scope_entries.append(scope_entry)
    tk.Label(scope_fm, text="例如 1-8,13-17", fg="gray").pack(side=tk.LEFT)
    bf = tk.Frame(cf); bf.pack(pady=10)
    tk.Button(bf, text="▶ 启动执行流", bg="#dff0d8", command=run_thread, width=22, height=2).pack(side=tk.LEFT, padx=6)
    tk.Button(bf, text="■ 安全终止 (F9)", bg="#f2dede", command=stop_auto, width=16, height=2).pack(side=tk.LEFT, padx=6)
    status_label = tk.Label(cf, text="就绪。F8 暂停/继续，F9 安全终止。", fg="green", font=("Arial", 10)); status_label.pack(pady=5)
    
    tab2 = ttk.Frame(nb); nb.add(tab2, text="⚙️ 患者管理"); p_left = tk.Frame(tab2); p_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    cols = ("bed", "name", "gender", "age", "admit_date", "complaint"); mgr_tree = ttk.Treeview(p_left, columns=cols, show="headings"); [mgr_tree.heading(c, text=t) or mgr_tree.column(c, width=60) for c, t in zip(cols, ["床号", "姓名", "性别", "年龄", "入院日", "主诉"])]; mgr_tree.pack(fill=tk.BOTH, expand=True); mgr_tree.bind('<<TreeviewSelect>>', on_patient_select)
    p_right = tk.Frame(tab2, width=280); p_right.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10); fds = [("床号:", "p_bed"), ("姓名:", "p_name"), ("性别:", "p_gender"), ("年龄:", "p_age"), ("入院日期:", "p_admit"), ("主诉摘要:", "p_comp"), ("入院诊断:", "p_adiag"), ("目前诊断:", "p_cdiag")]; entries = []
    for i, (l, v) in enumerate(fds): tk.Label(p_right, text=l).grid(row=i, column=0, sticky=tk.E, pady=3); e = tk.Entry(p_right, width=22); e.grid(row=i, column=1, pady=3); entries.append(e)
    p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag = entries; tk.Button(p_right, text="保存", command=save_patient, bg="#dff0d8").grid(row=10, column=0, columnspan=2, sticky=tk.EW, pady=10); tk.Button(p_right, text="删除", command=delete_patient, bg="#f2dede").grid(row=11, column=0, columnspan=2, sticky=tk.EW)
    tab3 = ttk.Frame(nb); nb.add(tab3, text="📝 模板管理"); t_left = tk.Frame(tab3, width=180); t_left.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5); mgr_tpl_listbox = tk.Listbox(t_left); mgr_tpl_listbox.pack(fill=tk.BOTH, expand=True); mgr_tpl_listbox.bind('<<ListboxSelect>>', on_template_select); tk.Button(t_left, text="删除", command=delete_template, bg="#f2dede").pack(fill=tk.X, pady=5)
    t_right = tk.Frame(tab3); t_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5); tpl_name_entry = tk.Entry(t_right); tpl_name_entry.pack(fill=tk.X, pady=2); tpl_content_text = tk.Text(t_right); tpl_content_text.pack(fill=tk.BOTH, expand=True, pady=5); tk.Button(t_right, text="保存", command=save_template, bg="#dff0d8", height=2).pack(fill=tk.X)
    tab4 = ttk.Frame(nb); nb.add(tab4, text="📄 会诊单复制")
    w_top = tk.Frame(tab4); w_top.pack(fill=tk.X, padx=12, pady=8)
    tk.Label(w_top, text="起始床号:").pack(side=tk.LEFT)
    word_start_entry = tk.Entry(w_top, width=8, justify='center'); word_start_entry.pack(side=tk.LEFT, padx=5)
    tk.Label(w_top, text="生成人数:").pack(side=tk.LEFT, padx=(15, 0))
    word_count_entry = tk.Entry(w_top, width=8, justify='center'); word_count_entry.pack(side=tk.LEFT, padx=5)
    tk.Button(w_top, text="生成并复制 (F1)", command=copy_word_consult_text, bg="#dff0d8").pack(side=tk.LEFT, padx=12)
    tk.Button(w_top, text="保存模板", command=save_word_template).pack(side=tk.LEFT)
    word_status_label = tk.Label(tab4, text="使用 current_diag；科室固定浦二；入院日期固定 2022.9.20。", fg="gray")
    word_status_label.pack(fill=tk.X, padx=12)
    word_template_text = tk.Text(tab4, font=("SimSun", 10), wrap=tk.WORD)
    word_template_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
    load_word_template()

    tab5 = ttk.Frame(nb); nb.add(tab5, text="⏱ 补病史")
    supp_top = tk.Frame(tab5); supp_top.pack(fill=tk.X, padx=10, pady=6)
    tk.Label(supp_top, text="起始日期:").pack(side=tk.LEFT, padx=(8, 0))
    supp_start_date_entry = tk.Entry(supp_top, width=11, justify="center"); supp_start_date_entry.insert(0, datetime.now().strftime("%Y-%m-%d")); supp_start_date_entry.pack(side=tk.LEFT, padx=3)
    supp_time_entry = tk.Entry(supp_top, width=6, justify="center"); supp_time_entry.insert(0, "09:00"); supp_time_entry.pack(side=tk.LEFT, padx=3)
    supp_interval_var = tk.StringVar(value="7天")
    ttk.Combobox(supp_top, width=7, textvariable=supp_interval_var, values=("7天", "14天", "1个月", "2个月", "3个月"), state="readonly").pack(side=tk.LEFT, padx=6)

    template_frame = tk.LabelFrame(tab5, text="病史模板"); template_frame.pack(fill=tk.X, padx=10, pady=8)
    supp_template_vars = [tk.StringVar() for _ in range(3)]
    supp_template_combos = []
    for index, variable in enumerate(supp_template_vars):
        tk.Label(template_frame, text=f"模板{index + 1}:").grid(row=index, column=0, sticky=tk.E, padx=6, pady=5)
        combo = ttk.Combobox(template_frame, width=34, textvariable=variable, state="readonly")
        combo.grid(row=index, column=1, sticky=tk.W, padx=6, pady=5)
        supp_template_combos.append(combo)

    title_frame = tk.LabelFrame(tab5, text="病史名称"); title_frame.pack(fill=tk.X, padx=10, pady=8)
    supp_title_entry = tk.Entry(title_frame, width=42)
    supp_title_entry.insert(0, DEFAULT_SUPPLEMENT_TITLE)
    supp_title_entry.pack(side=tk.LEFT, padx=8, pady=8)
    supp_blank_title_var = tk.BooleanVar(value=False)
    tk.Checkbutton(title_frame, text="病史名称留空", variable=supp_blank_title_var, command=toggle_supplement_title).pack(side=tk.LEFT, padx=8)

    supp_buttons = tk.Frame(tab5); supp_buttons.pack(fill=tk.X, padx=10, pady=6)
    tk.Button(supp_buttons, text="▶ 开始补病史", bg="#dff0d8", command=run_supplement_thread, width=22, height=2).pack(side=tk.RIGHT)

    root.bind_all("<F1>", copy_word_consult_text)
    refresh_all_data()
    start_global_hotkeys()
    root.mainloop()

if __name__ == "__main__":
    init_db(); setup_ui()
