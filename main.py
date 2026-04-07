import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pyperclip
import ctypes
import ctypes.wintypes
import threading
import time

# ==================== RPA 物理锚点与轮询配置 ====================
RPA_CONFIG = {
    "BTN_SWITCH_PATIENT": (31, 84),      # 1. 切换病人按钮 (无延迟)
    "PATIENT_FIRST_ROW": (86, 217),      # 2. 病人列表首行
    "LINE_HEIGHT": 18,                   # 3. 列表行高像素
    
    # --- 像素轮询核心配置 ---
    "READY_PIXEL_POS": (50, 120),        # 状态特征点坐标
    "READY_PIXEL_RGB": (245, 245, 245),  # 就绪状态的 RGB 值
    "MAX_WAIT_FREEZE": 12.0,             # 极限超时熔断时间(秒)
    
    "AREA_PROGRESS_RECORD": (64, 267),   # 4. 病程记录焦点区 (延时1s)
    "BTN_NEW_RECORD": (213, 82),         # 5. 新建病史按钮 (无延迟)
    "TPL_OPTION": (836, 369),            # 6. 二级菜单病史模板 (延时1s)
}

DB_FILE = "his_data.db"
is_running_auto = False
locked_tpl_name = ""

# ==================== 1. 底层硬件与显存模拟引擎 ====================
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

def mouse_click(x, y):
    user32.SetCursorPos(x, y)
    user32.mouse_event(2, 0, 0, 0, 0) # 左键按下
    user32.mouse_event(4, 0, 0, 0, 0) # 左键抬起

def mouse_double_click(x, y):
    mouse_click(x, y)
    time.sleep(0.1)
    mouse_click(x, y)

def get_pixel_color(x, y):
    """提取屏幕指定坐标的精准 RGB 值"""
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

# ==================== 3. 闭环极速自动化流水线 ====================
def start_automation_flow(loop_count):
    global is_running_auto, locked_tpl_name
    is_running_auto = True
    
    for i in range(loop_count):
        if not is_running_auto: break 
        
        status_update(f"正在处理第 {i+1}/{loop_count} 个患者...")
        
        # --- 动作 1: 唤醒列表 (无延迟) ---
        mouse_click(*RPA_CONFIG["BTN_SWITCH_PATIENT"])
        time.sleep(0.2) # 极微小缓冲，防止系统未响应点击
        
        # --- 动作 2: 步进寻址双击 ---
        target_y = RPA_CONFIG["PATIENT_FIRST_ROW"][1] + (i * RPA_CONFIG["LINE_HEIGHT"])
        mouse_double_click(RPA_CONFIG["PATIENT_FIRST_ROW"][0], target_y)
        
        # --- 动作 3: 像素轮询对抗系统假死 ---
        status_update(f"轮询侦测中 (极限 12s)... [进度: {i+1}/{loop_count}]")
        time.sleep(0.5) # [防御机制] 给予系统 0.5s 让按钮变灰，防止误读旧的白色
        
        polled_time = 0.5
        is_ready = False
        while polled_time < RPA_CONFIG["MAX_WAIT_FREEZE"]:
            if not is_running_auto: break
            
            current_rgb = get_pixel_color(*RPA_CONFIG["READY_PIXEL_POS"])
            if current_rgb == RPA_CONFIG["READY_PIXEL_RGB"]:
                is_ready = True
                status_update(f"就绪！仅耗时 {polled_time:.1f}s 提前切入！")
                time.sleep(0.2) # 确认就绪后的视觉过渡
                break
                
            time.sleep(0.1)
            polled_time += 0.1
            
        if not is_ready and is_running_auto:
            status_update("超时 12s 触发熔断，视为系统异常或无数据，跳过")
            continue
            
        if not is_running_auto: break # 二次防阻断检测

        # --- 动作 4: 屏幕读取与匹配 (策略 A: 熔断跳过) ---
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
            status_update(f"【跳过】未在数据库匹配到当前窗口患者")
            continue

        # --- 动作 5: 极速自动化点击与注入 ---
        # 焦点重置 (延时 1s)
        mouse_click(*RPA_CONFIG["AREA_PROGRESS_RECORD"])
        time.sleep(1.0)
        
        # 触发新建 (无延迟)
        mouse_click(*RPA_CONFIG["BTN_NEW_RECORD"])
        time.sleep(0.3) # 给菜单弹出一个渲染微秒
        
        # 挂载模板 (延时 1s)
        mouse_double_click(*RPA_CONFIG["TPL_OPTION"])
        time.sleep(1.0)
        
        # 提取数据并组装
        cursor.execute("SELECT content FROM templates WHERE name=?", (locked_tpl_name,))
        tpl_data = cursor.fetchone()
        conn.close()
        
        if not tpl_data: continue
        final_text = tpl_data[0]
        final_text = final_text.replace("{{name}}", str(target_p[1]))
        final_text = final_text.replace("{{gender}}", str(target_p[2]))
        final_text = final_text.replace("{{age}}", str(target_p[3]))
        final_text = final_text.replace("{{admit_date}}", str(target_p[4]))
        final_text = final_text.replace("{{complaint}}", str(target_p[5]))
        final_text = final_text.replace("{{admit_diag}}", str(target_p[6]))
        final_text = final_text.replace("{{current_diag}}", str(target_p[7]))
        
        # 渲染与注入
        pyperclip.copy(final_text)
        time.sleep(0.3)
        user32.keybd_event(0x11, 0, 0, 0) # Ctrl
        user32.keybd_event(0x56, 0, 0, 0) # V
        user32.keybd_event(0x56, 0, 2, 0)
        user32.keybd_event(0x11, 0, 2, 0)
        
        time.sleep(0.5) # 文本打入缓冲
        
        # 归档提交 (F3)
        user32.keybd_event(0x72, 0, 0, 0) 
        user32.keybd_event(0x72, 0, 2, 0)
        
        status_update(f"【归档成功】患者: {target_p[1]}")
        time.sleep(2) # 闭环等待 HIS 保存写入

    is_running_auto = False
    status_update("--- 极速流水线执行结束，请进行人工复核 ---")
    messagebox.showinfo("流水线终止", "任务已处理完毕。")

# ==================== 4. 界面交互与 CRUD ====================
def status_update(msg):
    status_label.config(text=msg)

def stop_auto():
    global is_running_auto
    is_running_auto = False
    status_update("【阻断响应】正在安全终止进程...")

def run_thread():
    try:
        count = int(loop_entry.get())
        if not locked_tpl_name:
            messagebox.showwarning("拦截", "请先在工作台锁定所需套用的模板！")
            return
        status_update("流水线将于 3 秒后启动，请切至 HIS 全屏...")
        root.update()
        time.sleep(3)
        threading.Thread(target=start_automation_flow, args=(count,), daemon=True).start()
    except ValueError:
        messagebox.showerror("参数异常", "请输入整数循环次数。")

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

def save_patient():
    data = (p_bed.get(), p_name.get(), p_gender.get(), p_age.get(), p_admit.get(), p_comp.get(), p_adiag.get(), p_cdiag.get())
    if not data[0] or not data[1]: return
    conn = sqlite3.connect(DB_FILE)
    conn.execute("REPLACE INTO patients VALUES (?,?,?,?,?,?,?,?)", data)
    conn.commit()
    conn.close()
    refresh_all_data()

def delete_patient():
    selected = mgr_tree.selection()
    if not selected: return
    bed_num = mgr_tree.item(selected)['values'][0]
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM patients WHERE bed=?", (bed_num,))
    conn.commit()
    conn.close()
    refresh_all_data()

def on_patient_select(event):
    selected = mgr_tree.selection()
    if not selected: return
    vals = mgr_tree.item(selected)['values']
    entries = [p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag]
    for i, entry in enumerate(entries):
        entry.delete(0, tk.END)
        entry.insert(0, str(vals[i]) if vals[i] != 'None' else "")

def save_template():
    t_name = tpl_name_entry.get()
    t_content = tpl_content_text.get("1.0", tk.END).strip()
    if not t_name or not t_content: return
    conn = sqlite3.connect(DB_FILE)
    conn.execute("REPLACE INTO templates VALUES (?,?)", (t_name, t_content))
    conn.commit()
    conn.close()
    refresh_all_data()

def delete_template():
    selected = mgr_tpl_listbox.curselection()
    if not selected: return
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM templates WHERE name=?", (mgr_tpl_listbox.get(selected),))
    conn.commit()
    conn.close()
    refresh_all_data()

def on_template_select(event):
    selected = mgr_tpl_listbox.curselection()
    if not selected: return
    conn = sqlite3.connect(DB_FILE)
    content = conn.execute("SELECT content FROM templates WHERE name=?", (mgr_tpl_listbox.get(selected),)).fetchone()[0]
    conn.close()
    tpl_name_entry.delete(0, tk.END)
    tpl_name_entry.insert(0, mgr_tpl_listbox.get(selected))
    tpl_content_text.delete("1.0", tk.END)
    tpl_content_text.insert(tk.END, content)

def setup_ui():
    global tpl_listbox, status_label, loop_entry, lock_var, root
    global mgr_tree, mgr_tpl_listbox, p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag, tpl_name_entry, tpl_content_text

    root = tk.Tk()
    root.title("轻量级精神科极速自动化工作站 (闭环轮询版)")
    root.geometry("850x550")
    root.attributes('-topmost', True) 

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # ---------- Tab 1: 自动化总控台 ----------
    tab_work = ttk.Frame(notebook)
    notebook.add(tab_work, text="🚀 自动化总控台")
    center_frame = tk.Frame(tab_work)
    center_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=20)
    
    tk.Label(center_frame, text="1. 选择并锁定需要批量套用的模板", font=("Arial", 11, "bold")).pack(pady=5)
    tpl_listbox = tk.Listbox(center_frame, height=6, font=("Arial", 11))
    tpl_listbox.pack(fill=tk.X)
    tpl_listbox.bind('<<ListboxSelect>>', on_tpl_select)
    lock_var = tk.BooleanVar()
    tk.Checkbutton(center_frame, text="🔒 锁定此模板", variable=lock_var, command=toggle_lock).pack(pady=5)
    
    tk.Label(center_frame, text="2. 极速流水线作业配置", font=("Arial", 11, "bold")).pack(pady=10)
    loop_fm = tk.Frame(center_frame)
    loop_fm.pack()
    tk.Label(loop_fm, text="预计批量处理床位数:").pack(side=tk.LEFT)
    loop_entry = tk.Entry(loop_fm, width=10, justify='center')
    loop_entry.insert(0, "1") 
    loop_entry.pack(side=tk.LEFT, padx=5)
    
    btn_fm = tk.Frame(center_frame)
    btn_fm.pack(pady=15)
    tk.Button(btn_fm, text="▶ 启动极速引擎", bg="#dff0d8", command=run_thread, width=25, height=2).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_fm, text="🛑 强制急停", bg="#f2dede", command=stop_auto, width=15, height=2).pack(side=tk.LEFT, padx=10)
    
    status_label = tk.Label(center_frame, text="系统已就绪，采用 RGB 闭环轮询模式", fg="green", font=("Arial", 10))
    status_label.pack(pady=10)

    # ---------- Tab 2 & 3: 数据管理 (与旧版相同) ----------
    tab_pmgr = ttk.Frame(notebook)
    notebook.add(tab_pmgr, text="⚙️ 患者管理台")
    pmgr_left = tk.Frame(tab_pmgr)
    pmgr_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    columns = ("bed", "name", "gender", "age", "admit_date", "complaint")
    mgr_tree = ttk.Treeview(pmgr_left, columns=columns, show="headings", selectmode="browse")
    for col, text in zip(columns, ["床号", "姓名", "性别", "年龄", "入院日", "主诉"]):
        mgr_tree.heading(col, text=text)
        mgr_tree.column(col, width=60)
    mgr_tree.pack(fill=tk.BOTH, expand=True)
    mgr_tree.bind('<<TreeviewSelect>>', on_patient_select)

    pmgr_right = tk.Frame(tab_pmgr, width=280)
    pmgr_right.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)
    fields = [("床号:", "p_bed"), ("姓名:", "p_name"), ("性别:", "p_gender"), ("年龄:", "p_age"), ("入院日期:", "p_admit"), ("主诉摘要:", "p_comp"), ("入院诊断:", "p_adiag"), ("目前诊断:", "p_cdiag")]
    entries = []
    for i, (label, var_name) in enumerate(fields):
        tk.Label(pmgr_right, text=label).grid(row=i, column=0, sticky=tk.E, pady=3)
        entry = tk.Entry(pmgr_right, width=22)
        entry.grid(row=i, column=1, pady=3)
        entries.append(entry)
    p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag = entries
    tk.Button(pmgr_right, text="保存 / 新增", command=save_patient, bg="#dff0d8").grid(row=10, column=0, columnspan=2, sticky=tk.EW, pady=10)
    tk.Button(pmgr_right, text="删除选中", command=delete_patient, bg="#f2dede").grid(row=11, column=0, columnspan=2, sticky=tk.EW)

    tab_tmgr = ttk.Frame(notebook)
    notebook.add(tab_tmgr, text="📝 模板管理台")
    tmgr_left = tk.Frame(tab_tmgr, width=180)
    tmgr_left.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
    mgr_tpl_listbox = tk.Listbox(tmgr_left)
    mgr_tpl_listbox.pack(fill=tk.BOTH, expand=True)
    mgr_tpl_listbox.bind('<<ListboxSelect>>', on_template_select)
    tk.Button(tmgr_left, text="删除选中模板", command=delete_template, bg="#f2dede").pack(fill=tk.X, pady=5)

    tmgr_right = tk.Frame(tab_tmgr)
    tmgr_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    tpl_name_entry = tk.Entry(tmgr_right)
    tpl_name_entry.pack(fill=tk.X, pady=2)
    tpl_content_text = tk.Text(tmgr_right)
    tpl_content_text.pack(fill=tk.BOTH, expand=True, pady=5)
    tk.Button(tmgr_right, text="保存 / 新增", command=save_template, bg="#dff0d8", height=2).pack(fill=tk.X)

    refresh_all_data()
    root.mainloop()

if __name__ == "__main__":
    init_db()
    setup_ui()
