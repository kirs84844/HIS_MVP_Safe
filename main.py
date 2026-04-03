import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pyperclip
import ctypes
import ctypes.wintypes
import threading
import time

DB_FILE = "his_data.db"
ready_text_to_inject = ""
status_label = None

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
    
    # 初始化演示模板
    cursor.execute("SELECT COUNT(*) FROM templates")
    if cursor.fetchone()[0] == 0:
        demo_tpl = """团队艺术治疗后小结
姓名：{{name}}   性别：{{gender}}   年龄：{{age}}   入院日期：{{admit_date}}
主诉：{{complaint}}
入院诊断：{{admit_diag}}
目前诊断：{{current_diag}}
目前情况：神清，仪态整，接触可，对答切题，思维连贯，未见明显幻错觉及感知觉综合障碍，情感淡漠，意志力减退，智能可，自知力无。
团队艺术治疗疗效：患者经过了接受式、参与式、再创造式以及即兴演奏式四期的团队艺术治疗，当事人能透过创作释放不安的情绪，达到了治疗目标。  
总体效果：良好。今起停医嘱：团队艺术治疗 qd。"""
        cursor.execute("INSERT INTO templates VALUES ('艺术治疗小结', ?)", (demo_tpl,))
    conn.commit()
    conn.close()

# ==================== 2. 核心注入与快捷键模块 ====================
def generate_and_ready():
    global ready_text_to_inject
    selected_tpl_idx = tpl_listbox.curselection()
    selected_item = tree.selection()
    
    if not selected_tpl_idx or not selected_item:
        messagebox.showwarning("提示", "请确保在左侧选中了患者，在右侧选中了模板。")
        return
        
    tpl_name = tpl_listbox.get(selected_tpl_idx)
    patient_data = tree.item(selected_item)['values']
    
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT content FROM templates WHERE name=?", (tpl_name,))
    tpl_content = cursor.fetchone()[0]
    conn.close()

    # 变量替换引擎
    final_text = tpl_content
    final_text = final_text.replace("{{name}}", str(patient_data[1]))
    final_text = final_text.replace("{{gender}}", str(patient_data[2]))
    final_text = final_text.replace("{{age}}", str(patient_data[3]))
    final_text = final_text.replace("{{admit_date}}", str(patient_data[4]))
    final_text = final_text.replace("{{complaint}}", str(patient_data[5]))
    final_text = final_text.replace("{{admit_diag}}", str(patient_data[6]))
    final_text = final_text.replace("{{current_diag}}", str(patient_data[7]))

    ready_text_to_inject = final_text
    status_label.config(text=f"状态: 已锁定【{patient_data[1]}】，请在 HIS 中按 F4 注入", fg="green")
    pyperclip.copy(final_text)

def inject_text():
    global ready_text_to_inject
    if not ready_text_to_inject:
        return
    pyperclip.copy(ready_text_to_inject)
    time.sleep(0.1)
    
    user32 = ctypes.windll.user32
    user32.keybd_event(0x11, 0, 0, 0) # Ctrl
    user32.keybd_event(0x56, 0, 0, 0) # V
    user32.keybd_event(0x56, 0, 0x0002, 0)       
    user32.keybd_event(0x11, 0, 0x0002, 0) 
    
    status_label.config(text="状态: 注入完成，等待下一次操作", fg="blue")
    ready_text_to_inject = ""

def hotkey_loop():
    user32 = ctypes.windll.user32
    user32.RegisterHotKey(None, 1, 0, 0x73) # F4
    msg = ctypes.wintypes.MSG()
    while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
        if msg.message == 0x0312: 
            inject_text()
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageA(ctypes.byref(msg))

# ==================== 3. 界面与交互模块 ====================
def refresh_all_data():
    """刷新所有列表数据"""
    # 刷新工作台病人树
    for row in tree.get_children(): tree.delete(row)
    # 刷新管理台病人树
    for row in mgr_tree.get_children(): mgr_tree.delete(row)
    
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM patients")
    for row in cursor.fetchall():
        tree.insert("", "end", values=row)
        mgr_tree.insert("", "end", values=row)
        
    # 刷新模板列表
    tpl_listbox.delete(0, tk.END)
    mgr_tpl_listbox.delete(0, tk.END)
    cursor.execute("SELECT name FROM templates")
    for row in cursor.fetchall():
        tpl_listbox.insert(tk.END, row[0])
        mgr_tpl_listbox.insert(tk.END, row[0])
    conn.close()

# --- 患者管理逻辑 ---
def save_patient():
    data = (p_bed.get(), p_name.get(), p_gender.get(), p_age.get(), 
            p_admit.get(), p_comp.get(), p_adiag.get(), p_cdiag.get())
    if not data[0] or not data[1]:
        messagebox.showwarning("错误", "床号和姓名不能为空！")
        return
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("REPLACE INTO patients VALUES (?,?,?,?,?,?,?,?)", data)
    conn.commit()
    conn.close()
    refresh_all_data()
    messagebox.showinfo("成功", "患者信息已保存！")

def delete_patient():
    selected = mgr_tree.selection()
    if not selected: return
    bed_num = mgr_tree.item(selected)['values'][0]
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM patients WHERE bed=?", (bed_num,))
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

# --- 模板管理逻辑 ---
def save_template():
    t_name = tpl_name_entry.get()
    t_content = tpl_content_text.get("1.0", tk.END).strip()
    if not t_name or not t_content:
        messagebox.showwarning("错误", "模板名称和内容不能为空！")
        return
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("REPLACE INTO templates VALUES (?,?)", (t_name, t_content))
    conn.commit()
    conn.close()
    refresh_all_data()
    messagebox.showinfo("成功", "模板已保存！")

def delete_template():
    selected = mgr_tpl_listbox.curselection()
    if not selected: return
    t_name = mgr_tpl_listbox.get(selected)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM templates WHERE name=?", (t_name,))
    conn.commit()
    conn.close()
    tpl_name_entry.delete(0, tk.END)
    tpl_content_text.delete("1.0", tk.END)
    refresh_all_data()

def on_template_select(event):
    selected = mgr_tpl_listbox.curselection()
    if not selected: return
    t_name = mgr_tpl_listbox.get(selected)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT content FROM templates WHERE name=?", (t_name,))
    content = cursor.fetchone()[0]
    conn.close()
    
    tpl_name_entry.delete(0, tk.END)
    tpl_name_entry.insert(0, t_name)
    tpl_content_text.delete("1.0", tk.END)
    tpl_content_text.insert(tk.END, content)

# ==================== 4. 主程序运行 ====================
def setup_ui():
    global tree, tpl_listbox, status_label, mgr_tree, mgr_tpl_listbox
    global p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag
    global tpl_name_entry, tpl_content_text

    root = tk.Tk()
    root.title("轻量级精神科病史辅助录入系统 - 独立全功能版")
    root.geometry("1000x600")

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # ---------- Tab 1: 日常工作台 ----------
    tab_work = ttk.Frame(notebook)
    notebook.add(tab_work, text="▶ 日常工作台")
    
    left_frame = tk.Frame(tab_work)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    tk.Label(left_frame, text="1. 选择当前患者 (双击不可编辑，请去管理台修改)").pack(anchor=tk.W)
    
    columns = ("bed", "name", "gender", "age", "admit_date", "complaint")
    tree = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="browse")
    for col, text in zip(columns, ["床号", "姓名", "性别", "年龄", "入院日", "主诉"]):
        tree.heading(col, text=text)
        tree.column(col, width=60)
    tree.pack(fill=tk.BOTH, expand=True)

    right_frame = tk.Frame(tab_work, width=250)
    right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)
    tk.Label(right_frame, text="2. 选择病史模板").pack(anchor=tk.W)
    tpl_listbox = tk.Listbox(right_frame, height=15)
    tpl_listbox.pack(fill=tk.X)
    
    btn_ready = tk.Button(right_frame, text="3. 提取并准备 (一键就绪)", height=3, bg="#d9edf7", command=generate_and_ready)
    btn_ready.pack(fill=tk.X, pady=15)
    status_label = tk.Label(right_frame, text="状态: 待机中...", fg="gray", wraplength=200)
    status_label.pack(fill=tk.X)

    # ---------- Tab 2: 患者管理台 ----------
    tab_pmgr = ttk.Frame(notebook)
    notebook.add(tab_pmgr, text="⚙️ 患者管理台")
    
    pmgr_left = tk.Frame(tab_pmgr)
    pmgr_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    mgr_tree = ttk.Treeview(pmgr_left, columns=columns, show="headings", selectmode="browse")
    for col, text in zip(columns, ["床号", "姓名", "性别", "年龄", "入院日", "主诉"]):
        mgr_tree.heading(col, text=text)
        mgr_tree.column(col, width=60)
    mgr_tree.pack(fill=tk.BOTH, expand=True)
    mgr_tree.bind('<<TreeviewSelect>>', on_patient_select)

    pmgr_right = tk.Frame(tab_pmgr, width=300)
    pmgr_right.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)
    
    tk.Label(pmgr_right, text="患者详细信息编辑", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, pady=5)
    fields = [("床号:", "p_bed"), ("姓名:", "p_name"), ("性别:", "p_gender"), 
              ("年龄:", "p_age"), ("入院日期:", "p_admit"), ("主诉摘要:", "p_comp"), 
              ("入院诊断:", "p_adiag"), ("目前诊断:", "p_cdiag")]
    
    entries = []
    for i, (label, var_name) in enumerate(fields):
        tk.Label(pmgr_right, text=label).grid(row=i+1, column=0, sticky=tk.E, pady=2)
        entry = tk.Entry(pmgr_right, width=25)
        entry.grid(row=i+1, column=1, pady=2)
        entries.append(entry)
    
    p_bed, p_name, p_gender, p_age, p_admit, p_comp, p_adiag, p_cdiag = entries
    
    tk.Button(pmgr_right, text="保存 / 新增", command=save_patient, bg="#dff0d8").grid(row=10, column=0, columnspan=2, fill=tk.X, pady=10)
    tk.Button(pmgr_right, text="删除选中患者", command=delete_patient, bg="#f2dede").grid(row=11, column=0, columnspan=2, fill=tk.X)

    # ---------- Tab 3: 模板管理台 ----------
    tab_tmgr = ttk.Frame(notebook)
    notebook.add(tab_tmgr, text="📝 模板管理台")
    
    tmgr_left = tk.Frame(tab_tmgr, width=200)
    tmgr_left.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
    tk.Label(tmgr_left, text="已有模板列表").pack(anchor=tk.W)
    mgr_tpl_listbox = tk.Listbox(tmgr_left)
    mgr_tpl_listbox.pack(fill=tk.BOTH, expand=True)
    mgr_tpl_listbox.bind('<<ListboxSelect>>', on_template_select)
    tk.Button(tmgr_left, text="删除选中模板", command=delete_template, bg="#f2dede").pack(fill=tk.X, pady=5)

    tmgr_right = tk.Frame(tab_tmgr)
    tmgr_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    tk.Label(tmgr_right, text="模板名称:").pack(anchor=tk.W)
    tpl_name_entry = tk.Entry(tmgr_right)
    tpl_name_entry.pack(fill=tk.X, pady=2)
    
    tk.Label(tmgr_right, text="支持的变量: {{name}} {{gender}} {{age}} {{admit_date}} {{complaint}} {{admit_diag}} {{current_diag}}", fg="blue").pack(anchor=tk.W)
    tpl_content_text = tk.Text(tmgr_right)
    tpl_content_text.pack(fill=tk.BOTH, expand=True, pady=5)
    
    tk.Button(tmgr_right, text="保存 / 新增模板", command=save_template, bg="#dff0d8", height=2).pack(fill=tk.X)

    # 启动初始化
    refresh_all_data()
    threading.Thread(target=hotkey_loop, daemon=True).start()
    root.mainloop()

if __name__ == "__main__":
    init_db()
    setup_ui()
