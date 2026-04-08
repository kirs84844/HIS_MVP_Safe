import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pyperclip
import ctypes
import ctypes.wintypes
import threading
import time

# [核心防御] 强制系统 DPI 感知，确保测出的坐标是绝对物理像素，无视 Windows 缩放
try:
    ctypes.windll.user32.SetProcessDPIAware()
except AttributeError:
    pass

# ==================== RPA 物理锚点与轮询配置 (实地测绘数据) ====================
RPA_CONFIG = {
    "BTN_SWITCH_PATIENT": (32, 96),      # 1. 切换病人按钮 
    "PATIENT_FIRST_ROW": (86, 215),      # 2. 病人列表首行
    "LINE_HEIGHT": 18,                   # 3. 列表行高像素
    
    # --- 像素轮询核心配置 ---
    "READY_PIXEL_POS": (28, 95),         # 用于侦测假死的特征点
    "READY_PIXEL_RGB": (245, 245, 245),  # 就绪变白时的 RGB
    "MAX_WAIT_FREEZE": 12.0,             # 极限超时熔断时间(秒)
    
    # --- UI 破障与交互配置 ---
    "AREA_SAFE_BLANK": (500, 642),       # 新增：纯空白安全区，用于消除列表遮挡
    "AREA_PROGRESS_RECORD": (60, 267),   # 病程记录区 (双击唤醒菜单)
    "TPL_OPTION": (825, 370),            # 二级菜单具体模板 
}

DB_FILE = "his_data.db"
is_running_auto = False
locked_tpl_name = ""

# ==================== 1. 底层硬件与显存模拟引擎 ====================
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

# ==================== 3. 闭环极速自动化流水线 ====================
def start_automation_flow(loop_count):
    global is_running_auto, locked_tpl_name
    is_running_auto = True
    
    for i in range(loop_count):
        if not is_running_auto: break 
        
        status_update(f"正在扫描第 {i+1}/{loop_count} 行列表...")
        
        # --- 动作 1: 唤醒列表 (无延迟) ---
        mouse_click(*RPA_CONFIG["BTN_SWITCH_PATIENT"])
        time.sleep(0.2) 
        
        # --- 动作 2: 步进寻址双击 ---
        target_y = RPA_CONFIG["PATIENT_FIRST_ROW"][1] + (i * RPA_CONFIG["LINE_HEIGHT"])
        mouse_double_click(RPA_CONFIG["PATIENT_FIRST_ROW"][0], target_y)
        
        # --- 动作 3: 【双段像素轮询】 ---
        status_update(f"侦测系统挂起状态... [进度: {i+1}/{loop_count}]")
        
        # 3.1: 侦测开始假死 (变灰)
        wait_busy = 0.0
        while wait_busy < 3.0:
            if not is_running_auto: break
            if get_pixel_color(*RPA_CONFIG["READY_PIXEL_POS"]) != RPA_CONFIG["READY_PIXEL_RGB"]:
                break 
            time.sleep(0.1)
            wait_busy += 0.1
            
        # 3.2: 侦测假死结束 (变回就绪白)
        polled_time = 0.0
        is_ready = False
        while polled_time < RPA_CONFIG["MAX_WAIT_FREEZE"]:
            if not is_running_auto: break
            if get_pixel_color(*RPA_CONFIG["READY_PIXEL_POS"]) == RPA_CONFIG["READY_PIXEL_RGB"]:
                is_ready = True
                status_update(f"界面就绪！强制缓冲 0.5s 等待标题文本刷出...")
                time.sleep(0.5) 
                break
            time.sleep(0.1)
            polled_time += 0.1
            
        if not is_ready and is_running_auto:
            status_update("超时 12s 未变白，触发熔断，切下一行")
            continue
            
        if not is_running_auto: break 

        # --- 动作 3.5: 防遮挡/破障点击 ---
        # 击碎残留的列表UI，并预夺焦点
        mouse_click(*RPA_CONFIG["AREA_SAFE_BLANK"])
        time.sleep(0.5)

        # --- 动作 4: 屏幕读取与匹配 ---
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
            status_update(f"【跳过】当前床位未在数据库中建档")
            continue

        # --- 动作 5: 极速触发展开菜单 ---
        mouse_double_click(*RPA_CONFIG["AREA_PROGRESS_RECORD"])
        time.sleep(1.0)
        
        # --- 动作 6: 挂载模板 ---
        mouse_double_click(*RPA_CONFIG["TPL_OPTION"])
        time.sleep(1.0)
        
        # --- 动作 7: 提取数据并组装 ---
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
        
        # 渲染与注入剪贴板
        pyperclip.copy(final_text)
        time.sleep(0.3)
        user32.keybd_event(0x11, 0, 0, 0) # Ctrl
        user32.keybd_event(0x56, 0, 0, 0) # V
        user32.keybd_event(0x56, 0, 2, 0)
        user32.keybd_event(0x11, 0, 2, 0)
        
        time.sleep(0.5) 
        
        # --- 动作 8: 归档提交 (F3) ---
        user32.keybd_event(0x72, 0, 0, 0) 
        user32.keybd_event(0x72, 0, 2, 0)
        
        status_update(f"【归档成功】患者: {target_p[1]}")
        time.sleep(2) 

    is_running_auto = False
    status_update("--- 极速流水线执行结束 ---")
    
    # 任务完成，自动从任务栏恢复辅助程序界面
    root.deiconify() 
    messagebox.showinfo("流水线终止", "分配的行数扫描与写入已全部执行完毕。")

# ==================== 4. 界面交互与 CRUD ====================
def status_update(msg):
    status_label.config(text=msg)

def stop_auto():
    global is_running_auto
    is_running_auto = False
    status_update("【阻断响应】正在安全终止进程...")
    root.deiconify() # 若中途停止，立刻恢复界面

def run_thread():
    try:
        count = int(loop_entry.get())
        if not locked_tpl_name:
            messagebox.showwarning("拦截", "请先在工作台锁定所需套用的模板！")
            return
        status_update("程序将自动隐藏，3 秒后接管 HIS...")
        root.update()
        time.sleep(3)
        
        # 核心：启动 RPA 前，自动将本程序最小化至任务栏，彻底消除遮挡
        root.iconify() 
        threading.Thread(target=start_automation_flow, args=(count,), daemon=True).start()
    except ValueError:
        messagebox.showerror("参数异常", "请输入整数行数。")

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
    root.title("轻量级精神科极速自动化工作站 (V5.2)")
    root.geometry("850x550")

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
    tk.Label(loop_fm, text="向下需扫描的列表总行数 (含空床):").pack(side=tk.LEFT)
    loop_entry = tk.Entry(loop_fm, width=10, justify='center')
    loop_entry.insert(0, "4") 
    loop_entry.pack(side=tk.LEFT, padx=5)
    
    btn_fm = tk.Frame(center_frame)
    btn_fm.pack(pady=15)
    tk.Button(btn_fm, text="▶ 启动极速引擎", bg="#dff0d8", command=run_thread, width=25, height=2).pack(side=tk.LEFT, padx=10)
    
    # 调整急停策略提示：程序运行时在后台，如需停止需手动点击任务栏恢复窗口再停止
    tk.Button(btn_fm, text="🛑 强制急停", bg="#f2dede", command=stop_auto, width=15, height=2).pack(side=tk.LEFT, padx=10)
    
    status_label = tk.Label(center_frame, text="系统已就绪，就绪后程序将自动最小化防遮挡", fg="green", font=("Arial", 10))
    status_label.pack(pady=10)

    # ---------- Tab 2 & 3: 数据管理 (保持不变) ----------
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
