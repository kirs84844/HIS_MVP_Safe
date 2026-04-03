import pyperclip
import ctypes
import ctypes.wintypes
import threading
import time

# 测试用本地数据库
patient_db = {
    "张三": "患者张三，男，35岁，精神分裂症。今日查房神志清，接触被动，存在言语性幻听，否认自杀观念。医嘱继续当前治疗方案。",
    "李四": "患者李四，女，28岁，重度抑郁发作。今日查房情绪低落，诉睡眠差，早醒。暂无消极轻生行为。密切观察情绪变化。"
}

current_patient = None

def monitor_clipboard():
    """静默监听剪贴板的后台线程"""
    global current_patient
    last_clipboard = ""
    print("【系统状态】安全版监听模块已启动...")
    print("【操作指引】请尝试复制 '张三' 或 '李四' 进行测试\n")
    
    while True:
        try:
            current_clipboard = pyperclip.paste().strip()
            if current_clipboard != last_clipboard:
                last_clipboard = current_clipboard
                if current_clipboard in patient_db:
                    current_patient = current_clipboard
                    print(f"[匹配成功] 已锁定病人资料: {current_patient}")
                    print(f"[等待注入] 请将光标移动至 HIS 输入框，并按下 F4 键！\n")
        except Exception:
            pass
        time.sleep(0.5)

def inject_text():
    """使用 Windows 原生 API 模拟 Ctrl+V 粘贴 (绕过杀软拦截)"""
    global current_patient
    if current_patient and current_patient in patient_db:
        record = patient_db[current_patient]
        print(f">>> 正在将 {current_patient} 的病历注入系统...")
        
        # 将文本写入剪贴板
        pyperclip.copy(record)
        time.sleep(0.1) # 预留 0.1 秒让系统剪贴板反应
        
        # 调用 user32.dll 模拟键盘动作
        user32 = ctypes.windll.user32
        VK_CONTROL = 0x11
        VK_V = 0x56
        KEYEVENTF_KEYUP = 0x0002
        
        user32.keybd_event(VK_CONTROL, 0, 0, 0) # 按下 Ctrl
        user32.keybd_event(VK_V, 0, 0, 0)       # 按下 V
        user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)       # 松开 V
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0) # 松开 Ctrl
        
        print(f">>> 注入完成！状态已重置，等待下一次匹配。\n")
        current_patient = None

def hotkey_loop():
    """使用 Windows 原生 API 注册全局快捷键 F4"""
    user32 = ctypes.windll.user32
    # 参数: 窗口句柄(None), 快捷键ID(1), 修饰键(0=无), 虚拟键码(0x73=F4)
    if not user32.RegisterHotKey(None, 1, 0, 0x73):
        print("【错误】F4 快捷键注册失败，可能被其他软件占用。")
        return

    # Windows 消息循环，用于持续接收快捷键被按下的系统通知
    msg = ctypes.wintypes.MSG()
    while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
        if msg.message == 0x0312: # WM_HOTKEY 消息
            inject_text()
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageA(ctypes.byref(msg))

if __name__ == "__main__":
    # 启动剪贴板监听子线程
    threading.Thread(target=monitor_clipboard, daemon=True).start()
    # 主线程进入快捷键监听循环
    hotkey_loop()
