import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pyperclip
import ctypes
import ctypes.wintypes
import threading
import time
from datetime import datetime, timedelta

# [核心防御] 强制系统 DPI 感知
try:
    ctypes.windll.user32.SetProcessDPIAware()
except AttributeError:
    pass

# ==================== RPA 物理锚点配置 ====================
RPA_CONFIG = {
    "BTN_SWITCH_PATIENT": (32, 96),      
    "PATIENT_FIRST_ROW": (86, 215),      
    "LINE_HEIGHT": 18,                   
    
    "READY_PIXEL_POS": (28, 95),         
    "READY_PIXEL_RGB": (245, 245, 245),  
    
    "AREA_SAFE_BLANK": (500, 642),       
    "AREA_PROGRESS_RECORD": (60, 267),   
    "TPL_OPTION": (825, 370),            
}

DB_FILE = "his_data.db"
is_running_auto = False
locked_tpl_name = ""

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
    for row in cursor.fetchall():
        tpl_listbox.insert(tk.END, row[0])
        mgr_tpl_listbox.insert(tk.END, row[0])
    conn.close()

# ==================== 3. 核心全自动逻辑引擎 (V5.6) ====================
def start_automation_flow(start_row, loop_count, target_time_obj):
    global is_running_auto, locked_tpl_name
    is_running_auto = True
    start_index = start_row - 1 
    
    for i in range(loop_count):
        if not is_running_auto: break 
        
        current_processing_row = start_row + i
        status_update(f"正在扫描第 {current_processing_row} 行 (进度: {i+1}/{loop_count})...")
        
        # --- 动作 1: 唤醒列表 ---
        mouse_click(*RPA_CONFIG["BTN_SWITCH_PATIENT"])
        time.sleep(0.2) 
        
        # --- 动作 2: 精确寻址切换 ---
        target_y = RPA_CONFIG["PATIENT_FIRST_ROW"][1] + ((start_index + i) * RPA_CONFIG["LINE_HEIGHT"])
        mouse_double_click(RPA_CONFIG["PATIENT_FIRST_ROW"][0], target_y)
        
        # --- 动作 3: 无极闭环轮询 ---
        status_update(f"第 {current_processing_row} 行: 等待复苏...")
        busy_check = 0.0
        while busy_check < 5.0:
            if not is_running_auto: break
            if get_pixel_color(*RPA_CONFIG["READY_PIXEL_POS"]) != RPA_CONFIG["READY_PIXEL_RGB"]: break
            time.sleep(0.2)
            busy_check += 0.2
            
        while is_running_auto:
            if get_pixel_color(*RPA_CONFIG["READY_PIXEL_POS"]) == RPA_CONFIG["READY_PIXEL_RGB"]:
                time.sleep(0.8) 
                break
            time.sleep(0.3) 

        if not is_running_auto: break 

        # --- 动作 3.5: 破障点击 (安全区) ---
        mouse_click(*RPA_CONFIG["AREA_SAFE_BLANK"])
        time.sleep(0.5)

        # --- 动作 4: 标题校验 ---
        hwnd = user32.GetForegroundWindow()
        length = user32.GetWindowTextLengthW(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buff, length + 1)
        window_title = buff.value

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM patients")
        patients = cursor.fetchall()
        target_p = None
        for p in patients:
            if str(p[1]) in window_title:
                target_p = p
                break
        
        if not target_p:
            conn.close()
            status_update(f"【跳过】行 {current_processing_row} 未建档")
            continue

        # --- 动作 5: 唤醒二级菜单 ---
        mouse_double_click(*RPA_CONFIG["AREA_PROGRESS_RECORD"])
        time.sleep(1.2) 
        
        # ==================== 核心分流控制 ====================
        if target_time_obj:
            # 【路径 B: 时间劫持模式】
            current_loop_time = target_time_obj + timedelta(minutes=i)
            time_str = current_loop_time.strftime("%y-%m-%d %H:%M")
            status_update(f"覆写时间: {time_str}")
            
            # 单击模板弹出属性框
            mouse_click(*RPA_CONFIG["TPL_OPTION"])
            time.sleep(0.5) 
            
            # Tab x2 定位时间框
            user32.keybd_event(0x09, 0, 0, 0); user32.keybd_event(0x09, 0, 2, 0)
            time.sleep(0.1)
            user32.keybd_event(0x09, 0, 0, 0); user32.keybd_event(0x09, 0, 2, 0)
            time.sleep(0.3)
            
            # 注入新时间
            pyperclip.copy(time_str)
            time.sleep(0.2)
            user32.keybd_event(0x11, 0, 0, 0); user32.keybd_event(0x56, 0, 0, 0)
            user32.keybd_event(0x56, 0, 2, 0); user32.keybd_event(0x11, 0, 2, 0)
            time.sleep(0.3)
            
            # Tab x2 跳至确定按钮
            user32.keybd_event(0x09, 0, 0, 0); user32.keybd_event(0x09, 0, 2, 0)
            time.sleep(0.1)
            user32.keybd_event(0x09, 0, 0, 0); user32.keybd_event(0x09, 0, 2, 0)
            time.sleep(0.3)
            
            # Enter 确认生成
            user32.keybd_event(0x0D, 0, 0, 0); user32.keybd_event(0x0D, 0, 2, 0)
            time.sleep(1.2) # 与路径 A 保持一致的编辑器加载时间
            
        else:
            # 【路径 A: 默认常规模式】
            mouse_double_click(*RPA_CONFIG["TPL_OPTION"])
            time.sleep(1.2) # 与路径 B 保持一致的编辑器加载时间
        # ========================================================

        # --- 动作 7: 数据提取与注入正文 ---
        cursor.execute("SELECT content FROM templates WHERE name=?", (locked_tpl_name,))
        tpl_data = cursor.fetchone()
        conn.close()
        
        if not tpl_data: continue
        final_text = tpl_data[0].replace("{{name}}", str(target_p[1])).replace("{{gender}}", str(target_p[2])).replace("{{age}}", str(target_p[3])).replace("{{admit_date}}", str(target_p[4])).replace("{{complaint}}", str(target_p[5])).replace("{{admit_diag}}", str(target_p[6])).replace("{{current_diag}}", str(target_p[7]))
        
        pyperclip.copy(final_text)
        time.sleep(0.4)
        user32.keybd_event(0x11, 0, 0, 0) # Ctrl
        user32.keybd_event(0x56, 0, 0, 0) # V
        user32.keybd_event(0x56, 0, 2, 0)
        user32.keybd_event(0x11, 0, 2, 0)
        time.sleep(0.8) 
        
        # --- 动作 8: 归档提交 (Ctrl + S) ---
        user32.keybd_event(0x11, 0, 0, 0) # Ctrl
        user32.keybd_event(0x53, 0, 0, 0) # S
        user32.keybd_event(0x53, 0, 2, 0) 
        user32.keybd_event(0x11, 0, 2, 0) 
        
        status_update(f"【成功】已归档: {target_p[1]}")
        time.sleep(2.5) 

    is_running_auto = False
    status_update("--- 任务结束 ---")
    root.deiconify() 
    messagebox.showinfo("任务完成", "流水线已处理完毕。")

# ==================== 4. UI 交互层 ====================
def status_update(msg):
    status_label.config(text=msg)

def stop_auto():
    global is_running_auto
    is_running_auto = False
    status_update("【紧急干预】正在强行终止...")
    root.deiconify()

def toggle_time_entry():
    if time_var.get():
        time_entry.config(state=tk.NORMAL)
    else:
        time_entry.config(state=tk.DISABLED)

def run_thread():
    try:
        start_row = int(start_entry.get())
        count = int(loop_entry.get())
        if start_row < 1 or count < 1:
            messagebox.showerror("参数异常", "请输入 ≥1 的正整数。")
            return
        if not locked_tpl_name:
            messagebox.showwarning("拦截", "请先锁定模板！")
            return
        
        target_time_obj = None
        if time_var.get():
            time_str = time_entry.get().strip()
            try:
                # 严格按照 YY-MM-DD HH:MM 格式解析
                target_time_obj = datetime.strptime(time_str, "%y-%m-%d %H:%M")
            except ValueError:
                messagebox.showerror("格式阻断", "时间格式存在瑕疵！\n请严格遵循: YY-MM-DD HH:MM\n(示例: 26-04-26 10:00)")
                return

        status_update("程序将自动隐藏，3 秒后接管 HIS...")
        root.update()
        time.sleep(3)
        root.iconify() 
        threading.Thread(target=start_automation_flow, args=(start_row, count, target_time_obj), daemon=True).start()
    except ValueError:
        messagebox.showerror("参数错误", "请输入有效的数字。")

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

# 数据库 CRUD 保留
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

def setup_ui():
    global tpl_listbox, status_label, loop_entry, start_entry, lock_var, root, mgr_tree, mgr_tpl_listbox, p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag, tpl_name_entry, tpl_content_text
    global time_var, time_entry
    
    root = tk.Tk(); root.title("极速精神科工作站 V5.6 (自适应双规模式)"); root.geometry("850x600")
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
    
    lf = tk.Frame(cf); lf.pack(pady=5)
    tk.Label(lf, text="起始向下扫描行:").pack(side=tk.LEFT)
    start_entry = tk.Entry(lf, width=8, justify='center'); start_entry.insert(0, "1"); start_entry.pack(side=tk.LEFT, padx=5)
    tk.Label(lf, text="连续扫描总行数:").pack(side=tk.LEFT, padx=(15, 0))
    loop_entry = tk.Entry(lf, width=8, justify='center'); loop_entry.insert(0, "4"); loop_entry.pack(side=tk.LEFT, padx=5)
    bf = tk.Frame(cf); bf.pack(pady=10); tk.Button(bf, text="▶ 启动执行流", bg="#dff0d8", command=run_thread, width=25, height=2).pack(side=tk.LEFT, padx=10); tk.Button(bf, text="🛑 急停", bg="#f2dede", command=stop_auto, width=15, height=2).pack(side=tk.LEFT, padx=10)
    status_label = tk.Label(cf, text="就绪。按需勾选时间覆写机制", fg="green", font=("Arial", 10)); status_label.pack(pady=5)
    
    tab2 = ttk.Frame(nb); nb.add(tab2, text="⚙️ 患者管理"); p_left = tk.Frame(tab2); p_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    cols = ("bed", "name", "gender", "age", "admit_date", "complaint"); mgr_tree = ttk.Treeview(p_left, columns=cols, show="headings"); [mgr_tree.heading(c, text=t) or mgr_tree.column(c, width=60) for c, t in zip(cols, ["床号", "姓名", "性别", "年龄", "入院日", "主诉"])]; mgr_tree.pack(fill=tk.BOTH, expand=True); mgr_tree.bind('<<TreeviewSelect>>', on_patient_select)
    p_right = tk.Frame(tab2, width=280); p_right.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10); fds = [("床号:", "p_bed"), ("姓名:", "p_name"), ("性别:", "p_gender"), ("年龄:", "p_age"), ("入院日期:", "p_admit"), ("主诉摘要:", "p_comp"), ("入院诊断:", "p_adiag"), ("目前诊断:", "p_cdiag")]; entries = []
    for i, (l, v) in enumerate(fds): tk.Label(p_right, text=l).grid(row=i, column=0, sticky=tk.E, pady=3); e = tk.Entry(p_right, width=22); e.grid(row=i, column=1, pady=3); entries.append(e)
    p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag = entries; tk.Button(p_right, text="保存", command=save_patient, bg="#dff0d8").grid(row=10, column=0, columnspan=2, sticky=tk.EW, pady=10); tk.Button(p_right, text="删除", command=delete_patient, bg="#f2dede").grid(row=11, column=0, columnspan=2, sticky=tk.EW)
    tab3 = ttk.Frame(nb); nb.add(tab3, text="📝 模板管理"); t_left = tk.Frame(tab3, width=180); t_left.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5); mgr_tpl_listbox = tk.Listbox(t_left); mgr_tpl_listbox.pack(fill=tk.BOTH, expand=True); mgr_tpl_listbox.bind('<<ListboxSelect>>', on_template_select); tk.Button(t_left, text="删除", command=delete_template, bg="#f2dede").pack(fill=tk.X, pady=5)
    t_right = tk.Frame(tab3); t_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5); tpl_name_entry = tk.Entry(t_right); tpl_name_entry.pack(fill=tk.X, pady=2); tpl_content_text = tk.Text(t_right); tpl_content_text.pack(fill=tk.BOTH, expand=True, pady=5); tk.Button(t_right, text="保存", command=save_template, bg="#dff0d8", height=2).pack(fill=tk.X)
    refresh_all_data(); root.mainloop()

if __name__ == "__main__":
    init_db(); setup_ui()
