import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pyperclip
import ctypes
import ctypes.wintypes
import threading
import time

DB_FILE = "his_data.db"
status_label = None
lock_var = None
locked_tpl_name = ""

# ==================== 1. 数据库基础模块 ====================
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

# ==================== 2. 全自动化注入引擎 ====================
def handle_f4_injection():
    """F4 触发的全自动处理链路"""
    global locked_tpl_name
    
    if not lock_var.get() or not locked_tpl_name:
        status_label.config(text="【失败】请先在界面上选中一个模板并勾选[锁定]", fg="red")
        return

    # 1. 获取当前活动窗口(HIS)的标题
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    length = user32.GetWindowTextLengthW(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buff, length + 1)
    window_title = buff.value

    # 2. 从数据库加载所有患者并进行反向逆推匹配
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM patients")
    patients = cursor.fetchall()
    
    # 按照名字长度降序排序，防止"李海"匹配到"李海元"
    patients.sort(key=lambda x: len(str(x[1])), reverse=True)
    
    target_patient = None
    for p in patients:
        if str(p[1]) in window_title: # 如果数据库中的名字出现在了 HIS 的窗口标题里
            target_patient = p
            break
            
    if not target_patient:
        conn.close()
        status_label.config(text=f"【忽略】当前窗口标题未匹配到库中患者: {window_title[:15]}...", fg="gray")
        return

    # 3. 提取被锁定的模板并渲染
    cursor.execute("SELECT content FROM templates WHERE name=?", (locked_tpl_name,))
    tpl_data = cursor.fetchone()
    conn.close()
    
    if not tpl_data: return
    
    final_text = tpl_data[0]
    final_text = final_text.replace("{{name}}", str(target_patient[1]))
    final_text = final_text.replace("{{gender}}", str(target_patient[2]))
    final_text = final_text.replace("{{age}}", str(target_patient[3]))
    final_text = final_text.replace("{{admit_date}}", str(target_patient[4]))
    final_text = final_text.replace("{{complaint}}", str(target_patient[5]))
    final_text = final_text.replace("{{admit_diag}}", str(target_patient[6]))
    final_text = final_text.replace("{{current_diag}}", str(target_patient[7]))

    # 4. 模拟粘贴输出
    pyperclip.copy(final_text)
    time.sleep(0.1) # 等待系统剪贴板就绪
    
    user32.keybd_event(0x11, 0, 0, 0) # Ctrl
    user32.keybd_event(0x56, 0, 0, 0) # V
    user32.keybd_event(0x56, 0, 0x0002, 0)       
    user32.keybd_event(0x11, 0, 0x0002, 0) 
    
    status_label.config(text=f"【成功】已为 [{target_patient[1]}] 自动注入 <{locked_tpl_name}>", fg="green")

def hotkey_loop():
    user32 = ctypes.windll.user32
    user32.RegisterHotKey(None, 1, 0, 0x73) # 注册 F4
    msg = ctypes.wintypes.MSG()
    while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
        if msg.message == 0x0312: 
            handle_f4_injection()
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageA(ctypes.byref(msg))

# ==================== 3. 界面逻辑 ====================
def on_tpl_select(event):
    global locked_tpl_name
    selected = tpl_listbox.curselection()
    if selected:
        locked_tpl_name = tpl_listbox.get(selected)
        if lock_var.get():
            status_label.config(text=f"状态: 已锁定模板 <{locked_tpl_name}>，请直接前往 HIS 按 F4", fg="blue")

def toggle_lock():
    global locked_tpl_name
    selected = tpl_listbox.curselection()
    if lock_var.get():
        if not selected:
            lock_var.set(False)
            messagebox.showwarning("提示", "请先点击选中一个模板，再勾选锁定。")
        else:
            locked_tpl_name = tpl_listbox.get(selected)
            status_label.config(text=f"状态: 默认模板已锁定为 <{locked_tpl_name}>", fg="blue")
    else:
        status_label.config(text="状态: 锁定已解除，F4 注入已暂停", fg="gray")

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

# 增删改查函数 (与上版相同，略去细枝末节，保证直接运行)
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

# ==================== 4. 主 UI 构建 ====================
def setup_ui():
    global tpl_listbox, status_label, mgr_tree, mgr_tpl_listbox, lock_var
    global p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag, tpl_name_entry, tpl_content_text

    root = tk.Tk()
    root.title("轻量级精神科工作站 - 极速直连版 v3.0")
    root.geometry("850x550")

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # ---------- Tab 1: 极速驾驶舱 ----------
    tab_work = ttk.Frame(notebook)
    notebook.add(tab_work, text="🚀 极速工作台")
    
    center_frame = tk.Frame(tab_work)
    center_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=20)
    
    tk.Label(center_frame, text="工作流已极简化：无需再手动选择患者。", font=("Arial", 12, "bold"), fg="blue").pack(pady=10)
    tk.Label(center_frame, text="操作步骤：\n1. 在下方选中今天要写的病史模板，并勾选锁定。\n2. 直接去 HIS 系统点开病人的病历界面。\n3. 在输入框按下键盘 F4，系统自动识别当前窗口病人并输入。").pack(pady=5)
    
    tpl_listbox = tk.Listbox(center_frame, height=8, font=("Arial", 11))
    tpl_listbox.pack(fill=tk.X, pady=10)
    tpl_listbox.bind('<<ListboxSelect>>', on_tpl_select)
    
    lock_var = tk.BooleanVar()
    tk.Checkbutton(center_frame, text="🔒 锁定选中的模板为【默认提取模板】", variable=lock_var, command=toggle_lock, font=("Arial", 11, "bold")).pack(pady=5)
    
    status_label = tk.Label(center_frame, text="状态: 等待锁定模板...", fg="gray", font=("Arial", 10))
    status_label.pack(pady=10)

    # ---------- Tab 2: 患者管理 ----------
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
    tk.Button(pmgr_right, text="删除选中患者", command=delete_patient, bg="#f2dede").grid(row=11, column=0, columnspan=2, sticky=tk.EW)

    # ---------- Tab 3: 模板管理 ----------
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
    tk.Button(tmgr_right, text="保存 / 新增模板", command=save_template, bg="#dff0d8", height=2).pack(fill=tk.X)

    refresh_all_data()
    threading.Thread(target=hotkey_loop, daemon=True).start()
    root.mainloop()

if __name__ == "__main__":
    init_db()
    setup_ui()
