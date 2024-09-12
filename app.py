import tkinter as tk
from tkinter import ttk, font
import win32gui
import win32con
import time

class WindowPriorityManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Window Priority Manager")
        self.root.geometry("650x450")
        self.root.attributes('-topmost', True)

        # 使用更现代的主题
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # 设置默认字体
        self.default_font = font.nametofont("TkDefaultFont")
        self.default_font.configure(family="Segoe UI", size=12)  # 增大默认字体大小

        # 主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建一个框架来包含树状视图和滚动条
        tree_frame = ttk.Frame(self.main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 树状视图
        self.tree = ttk.Treeview(tree_frame, columns=('Priority', 'Window'), show='headings', style="Custom.Treeview")
        self.tree.heading('Priority', text='Priority')
        self.tree.heading('Window', text='Window')
        self.tree.column('Priority', width=80, anchor='center')
        self.tree.column('Window', width=550, anchor='w')
        
        # 添加垂直滚动条
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        # 使用网格布局放置树状视图和滚动条
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')

        # 配置tree_frame的网格
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        # 自定义树状视图样式
        self.style.configure("Custom.Treeview", 
                        background="#f5f5f5",
                        foreground="black",
                        rowheight=30,  # 增加行高
                        fieldbackground="#f5f5f5",
                        borderwidth=1,
                        relief='solid',
                        font=('Segoe UI', 12))  # 设置树状视图字体
        self.style.map('Custom.Treeview', background=[('selected', '#e1e1e1')])

        # 自定义标题样式
        self.style.configure("Custom.Treeview.Heading",
                        background="#e1e1e1",
                        foreground="black",
                        relief="flat",
                        font=('Segoe UI', 12, 'bold'))  # 增大标题字体
        self.style.map("Custom.Treeview.Heading",
                       background=[('active', '#d1d1d1')])

        # 按钮框架
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill=tk.X, pady=10)

        # 自定义按钮样式
        self.style.configure("TButton", padding=6, relief="flat", background="#4CAF50", foreground="white", font=('Segoe UI', 12))
        self.style.map("TButton", background=[('active', '#45a049')])

        self.refresh_button = ttk.Button(self.button_frame, text="Refresh", command=self.update_window_list, style="TButton")
        self.top_button = ttk.Button(self.button_frame, text="Set Top", command=self.set_top, style="TButton")
        self.bottom_button = ttk.Button(self.button_frame, text="Set Bottom", command=self.set_bottom, style="TButton")
        self.minimize_button = ttk.Button(self.button_frame, text="Minimize", command=self.minimize_to_pin, style="TButton")

        # 使用网格布局使按钮居中
        self.button_frame.columnconfigure((0,1,2,3,4,5), weight=1)
        self.refresh_button.grid(row=0, column=1, padx=5, pady=5)
        self.top_button.grid(row=0, column=2, padx=5, pady=5)
        self.bottom_button.grid(row=0, column=3, padx=5, pady=5)
        self.minimize_button.grid(row=0, column=4, padx=5, pady=5)

        self.tree.bind('<ButtonPress-1>', self.on_press)
        self.tree.bind('<B1-Motion>', self.on_move)
        self.tree.bind('<ButtonRelease-1>', self.on_release)

        self.drag_start_item = None
        self.pin_window = None
        self.is_dragging_pin = False

        # 存储原始窗口状态
        self.original_window_states = {}

        # 捕获系统最小化和关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Unmap>", self.on_minimize)

        self.update_window_list()

        self.keep_gui_on_top()

    def keep_gui_on_top(self):
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.after(100, self.keep_gui_on_top)  # 每100毫秒检查一次

    def create_pin_window(self):
        self.pin_window = tk.Toplevel(self.root)
        self.pin_window.overrideredirect(True)
        self.pin_window.attributes('-topmost', True)
        self.pin_window.geometry("30x30+{}+{}".format(self.root.winfo_x(), self.root.winfo_y()))
        
        # 美化图钉按钮
        pin_button = tk.Button(self.pin_window, text="📌", bg="#4CAF50", fg="white", relief="flat", 
                               activebackground="#45a049", activeforeground="white", bd=0)
        pin_button.pack(fill=tk.BOTH, expand=True)

        self.pin_window.bind('<Button-1>', self.start_move_pin)
        self.pin_window.bind('<B1-Motion>', self.on_move_pin)
        self.pin_window.bind('<ButtonRelease-1>', self.stop_move_pin)

        self.is_dragging_pin = False
        self.drag_start_time = 0

        # 定期检查并确保图钉窗口置顶
        self.keep_pin_on_top()

    def keep_pin_on_top(self):
        if self.pin_window:
            self.pin_window.attributes('-topmost', True)
            self.pin_window.lift()
        # 每100毫秒检查一次
        self.root.after(100, self.keep_pin_on_top)

    def minimize_to_pin(self):
        self.root.withdraw()
        if not self.pin_window:
            self.create_pin_window()
        else:
            self.pin_window.deiconify()
        self.pin_window.attributes('-topmost', True)
        self.pin_window.lift()

    def restore_from_pin(self):
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        self.root.lift()
        if self.pin_window:
            self.pin_window.withdraw()

    def start_move_pin(self, event):
        self.is_dragging_pin = False
        self.x = event.x
        self.y = event.y
        self.drag_start_time = time.time()

    def on_move_pin(self, event):
        if not self.is_dragging_pin:
            if abs(event.x - self.x) > 5 or abs(event.y - self.y) > 5:
                self.is_dragging_pin = True
        if self.is_dragging_pin:
            deltax = event.x - self.x
            deltay = event.y - self.y
            x = self.pin_window.winfo_x() + deltax
            y = self.pin_window.winfo_y() + deltay
            self.pin_window.geometry(f"+{x}+{y}")

    def stop_move_pin(self, event):
        if not self.is_dragging_pin:
            if time.time() - self.drag_start_time < 0.2:
                self.restore_from_pin()
        self.is_dragging_pin = False

    def update_window_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        windows = self.get_window_list()
        for priority, (hwnd, title) in enumerate(windows, start=1):
            self.tree.insert('', 'end', values=(priority, title), tags=(str(hwnd),))
            # 存储原始窗口状态
            if hwnd not in self.original_window_states:
                self.original_window_states[hwnd] = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        self.root.lift()  # 确保GUI在最顶层

    def get_window_list(self):
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                windows.append((hwnd, win32gui.GetWindowText(hwnd)))
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        return windows

    def on_press(self, event):
        self.drag_start_item = self.tree.identify_row(event.y)

    def on_move(self, event):
        if self.drag_start_item:
            drag_to = self.tree.identify_row(event.y)
            if drag_to and drag_to != self.drag_start_item:
                self.tree.move(self.drag_start_item, self.tree.parent(drag_to), self.tree.index(drag_to))

    def on_release(self, event):
        if self.drag_start_item:
            self.update_priorities()
            self.drag_start_item = None

    def update_priorities(self):
        for index, item in enumerate(self.tree.get_children(), start=1):
            hwnd = int(self.tree.item(item, 'tags')[0])
            self.tree.set(item, 'Priority', index)
            
            if index == 1:
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            else:
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, 
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0, 
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        self.root.lift()  # 确保GUI在最顶层

    def set_top(self):
        selected = self.tree.selection()
        if selected:
            hwnd = int(self.tree.item(selected[0], 'tags')[0])
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            self.root.lift()  # 确保GUI在最顶层
            self.update_window_list()

    def set_bottom(self):
        selected = self.tree.selection()
        if selected:
            hwnd = int(self.tree.item(selected[0], 'tags')[0])
            win32gui.SetWindowPos(hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0, 
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            self.root.lift()  # 确保GUI在最顶层
            self.update_window_list()

    def on_minimize(self, event):
        if self.root.state() == 'iconic':
            self.minimize_to_pin()

    def on_close(self):
        self.restore_original_states()
        if self.pin_window:
            self.pin_window.destroy()
        self.root.destroy()

    def restore_original_states(self):
        for hwnd, original_state in self.original_window_states.items():
            try:
                current_state = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                new_state = (current_state & ~win32con.WS_EX_TOPMOST) | (original_state & win32con.WS_EX_TOPMOST)
                win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, new_state)
                
                # 重新设置窗口的Z顺序
                if not (original_state & win32con.WS_EX_TOPMOST):
                    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, 
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                else:
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            except:
                # 窗口可能已经关闭，忽略错误
                pass

# 主程序
if __name__ == "__main__":
    root = tk.Tk()
    app = WindowPriorityManager(root)
    root.mainloop()
