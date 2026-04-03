import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pyperclip
import ctypes
import ctypes.wintypes
import threading
import time
import re

# ---------------- 数据库模块 ----------------
DB_FILE = "his_data.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    # 创建病人表
    cursor.execute('''CREATE TABLE IF NOT EXISTS patients (
                        bed TEXT PRIMARY KEY, name TEXT, gender TEXT, 
                        age TEXT, admit_date TEXT, complaint TEXT, 
                        admit_diag TEXT, current_diag TEXT)''')
    # 创建模板表
    cursor.execute('''CREATE TABLE IF NOT EXISTS templates (
                        name TEXT PRIMARY KEY, content TEXT)''')
    
    # 插入演示数据 (如果为空)
    cursor.execute("SELECT COUNT(*) FROM patients")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO patients VALUES ('01', '凌云', '男', '65岁', '2022.9.20', '言行紊乱，冲动毁物，总病程40年。', '精神分裂症', '精神分裂症')")
        
    cursor.execute("SELECT COUNT(*) FROM templates")
    if cursor.fetchone()[0] == 0:
        demo_tpl = """团队艺术治疗后小结
姓名：{{name}}   性别：{{gender}}   年龄：{{age}}   入院日期：{{admit_date}}
主诉：{{complaint}}
入院诊断：{{admit_diag}}
目前诊断：{{current_diag}}
目前情况：神清，仪态整，接触可，对答切题，思维连贯，未见明显幻错觉及感知觉综合障碍，情感淡漠，意志力减退，智能可，自知力无。
团队艺术治疗疗效：患者经过了接受式、参与式、再创造式以及即兴演奏式四期的团队艺术治疗，当事人能透过创作释放不安的情绪 ,将旧有的经验加以澄清。在将意念化为具体形象的过程中 ,传递出个人目前的需求与情绪 ,经过分享与讨论 ,使其人格变得统合。从团队分享讨论中，通过观察、学习、体验，认识自我、探讨自我、接纳自我，消除了自卑感，提高兴趣和增加人际交往能力，达到了治疗目标。  
总体效果：良好。今起停医嘱：团队艺术治疗 qd。"""
        cursor.execute("INSERT INTO templates VALUES ('艺术治疗小结', ?)", (demo_tpl,))
    conn.commit()
    conn.close()

# ---------------- 全局变量 ----------------
ready_text_to_inject = ""
status_label = None

# ---------------- 核心逻辑模块 ----------------
def generate_and_ready():
    global ready_text_to_inject
    
    # 获取选中的模板
    selected_tpl_idx = tpl_listbox.curselection()
    if not selected_tpl_idx:
        messagebox.showwarning("提示", "请先在右侧选择一个模板")
        return
    tpl_name = tpl_listbox.get(selected_tpl_idx)
    
    # 获取选中的病人
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("提示", "请先在左侧选择一个病人")
        return
    
    patient_data = tree.item(selected_item)['values']
    # values: [bed, name, gender, age, admit_date, complaint, admit_diag, current_diag]
    
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT content FROM templates WHERE name=?", (tpl_name,))
    tpl_content = cursor.fetchone()[0]
    conn.close()

    # 执行替换逻辑
    final_text = tpl_content
    final_text = final_text.replace("{{name}}", str(patient_data[1]))
    final_text = final_text.replace("{{gender}}", str(patient_data[2]))
    final_text = final_text.replace("{{age}}", str(patient_data[3]))
    final_text = final_text.replace("{{admit_date}}", str(patient_data[4]))
    final_text = final_text.replace("{{complaint}}", str(patient_data[5]))
    final_text = final_text.replace("{{admit_diag}}", str(patient_data[6]))
    final_text = final_text.replace("{{current_diag}}", str(patient_data[7]))

    ready_text_to_inject = final_text
    status_label.config(text=f"状态: 已锁定【{patient_data[1]}】的【{tpl_name}】，请在 HIS 中按 F4 注入", fg="green")
    pyperclip.copy(final_text) # 同时放入系统剪贴板作为双保险

# ---------------- F4 快捷键注入模块 ----------------
def inject_text():
    global ready_text_to_inject
    if not ready_text_to_inject:
        return
    
    pyperclip.copy(ready_text_to_inject)
    time.sleep(0.1)
    
    user32 = ctypes.windll.user32
    VK_CONTROL = 0x11
    VK_V = 0x56
    KEYEVENTF_KEYUP = 0x0002
    
    user32.keybd_event(VK_CONTROL, 0, 0, 0) 
    user32.keybd_event(VK_V, 0, 0, 0)       
    user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)       
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0) 
    
    status_label.config(text="状态: 注入完成，等待下一次操作", fg="blue")
    ready_text_to_inject = ""

def hotkey_loop():
    user32 = ctypes.windll.user32
    user32.RegisterHotKey(None, 1, 0, 0x73) # 0x73 is F4
    msg = ctypes.wintypes.MSG()
    while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
        if msg.message == 0x0312: 
            inject_text()
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageA(ctypes.byref(msg))

# ---------------- UI 界面模块 ----------------
def load_patients():
    for row in tree.get_children():
        tree.delete(row)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM patients")
    for row in cursor.fetchall():
        tree.insert("", "end", values=row)
    conn.close()

def load_templates():
    tpl_listbox.delete(0, tk.END)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM templates")
    for row in cursor.fetchall():
        tpl_listbox.insert(tk.END, row[0])
    conn.close()

def setup_ui():
    global tree, tpl_listbox, status_label
    root = tk.Tk()
    root.title("轻量级精神科病史辅助录入系统 (Win7 兼容版)")
    root.geometry("900x500")
    
    # 左侧：病人列表
    left_frame = tk.Frame(root)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
    tk.Label(left_frame, text="1. 选择当前患者").pack(anchor=tk.W)
    
    columns = ("bed", "name", "gender", "age", "admit_date", "complaint", "admit_diag", "current_diag")
    tree = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="browse")
    tree.heading("bed", text="床号")
    tree.heading("name", text="姓名")
    tree.heading("gender", text="性别")
    tree.column("bed", width=50)
    tree.column("name", width=80)
    tree.column("gender", width=50)
    tree.pack(fill=tk.BOTH, expand=True)
    
    # 右侧：模板与操作
    right_frame = tk.Frame(root, width=300)
    right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)
    tk.Label(right_frame, text="2. 选择要写入的模板").pack(anchor=tk.W)
    
    tpl_listbox = tk.Listbox(right_frame, height=10)
    tpl_listbox.pack(fill=tk.X)
    
    btn_ready = tk.Button(right_frame, text="3. 提取数据并准备 (一键就绪)", height=3, bg="lightblue", command=generate_and_ready)
    btn_ready.pack(fill=tk.X, pady=20)
    
    status_label = tk.Label(right_frame, text="状态: 待机中", fg="gray", wraplength=250, justify=tk.LEFT)
    status_label.pack(fill=tk.X)
    
    tk.Label(right_frame, text="*修改数据/模板请使用 SQLite 工具\n本界面为极速工作台", fg="gray").pack(side=tk.BOTTOM)

    load_patients()
    load_templates()
    
    # 启动快捷键监听线程
    threading.Thread(target=hotkey_loop, daemon=True).start()
    
    root.mainloop()

if __name__ == "__main__":
    init_db()
    setup_ui()

