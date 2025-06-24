import tkinter as tk
from tkinter import ttk, messagebox
import keyboard
import pyautogui
from PIL import Image, ImageGrab, ImageTk
import ctypes
import sys
import pystray
import threading

class ColorPickerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("高级取色器")
        self.hotkey = 'alt+c'
        self.scale_factor = self.get_scale_factor()
        self.picking = False
        self.tray_icon = None
        self.magnifier_photo = None

        # 初始化并隐藏主窗口
        self.setup_ui()
        self.root.withdraw()

        # 设置窗口图标
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass

        # 托盘 & 热键
        self.setup_tray_icon()
        self.register_hotkey()
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

    def get_scale_factor(self):
        """获取 DPI 缩放因子"""
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            user32 = ctypes.windll.user32
            gdi32 = ctypes.windll.gdi32
            dc = user32.GetDC(0)
            dpi = gdi32.GetDeviceCaps(dc, 88)
            user32.ReleaseDC(0, dc)
            return dpi / 96.0
        except:
            return 1.0

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧颜色预览区域
        preview_frame = ttk.Frame(main_frame)
        preview_frame.grid(row=0, column=0, padx=5, pady=5)
        self.color_canvas = tk.Canvas(preview_frame, width=60, height=60, bg='white')
        self.color_canvas.pack()

        # 右侧控制按钮
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(control_frame, text="取色", command=self.pick_color).pack(pady=5)

        # 热键设置
        hotkey_frame = ttk.Frame(control_frame)
        hotkey_frame.pack(pady=5)
        ttk.Label(hotkey_frame, text="热键:").pack(side=tk.LEFT)
        self.hotkey_entry = ttk.Entry(hotkey_frame, width=10)
        self.hotkey_entry.insert(0, self.hotkey)
        self.hotkey_entry.pack(side=tk.LEFT, padx=2)
        ttk.Button(hotkey_frame, text="设置", width=4, command=self.update_hotkey).pack(side=tk.LEFT)

        # 颜色值显示
        color_info_frame = ttk.Frame(main_frame)
        color_info_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky=tk.EW)
        self.hex_entry = ttk.Entry(color_info_frame, width=20, font=('Arial', 10))
        self.hex_entry.pack(pady=2, fill=tk.X)
        self.rgb_entry = ttk.Entry(color_info_frame, width=20, font=('Arial', 10))
        self.rgb_entry.pack(pady=2, fill=tk.X)

        self.root.minsize(280, 180)
        self.root.resizable(False, False)

    def generate_icon_image(self):
        img = Image.new('RGB', (64, 64), (0, 122, 204))
        return img

    def setup_tray_icon(self):
        menu = (
            pystray.MenuItem("显示窗口", lambda icon, item: self.show_window()),
            pystray.MenuItem("退出", lambda icon, item: self.on_close())
        )
        icon_image = self.generate_icon_image()
        self.tray_icon = pystray.Icon("color_picker", icon_image, "取色器", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_window(self):
        self.root.after(0, self.root.deiconify)
        self.root.attributes('-topmost', 1)
        self.root.attributes('-topmost', 0)

    def hide_window(self):
        self.root.withdraw()

    def register_hotkey(self):
        try:
            keyboard.add_hotkey(self.hotkey, self.pick_color)
        except Exception as e:
            messagebox.showerror("错误", f"快捷键注册失败: {e}")

    def update_hotkey(self):
        new = self.hotkey_entry.get().strip()
        if not new:
            return
        try:
            keyboard.remove_hotkey(self.hotkey)
            keyboard.add_hotkey(new, self.pick_color)
            self.hotkey = new
        except Exception as e:
            messagebox.showerror("错误", f"设置热键失败: {e}")

    def pick_color(self):
        """启动取色交互：隐藏主窗口，全屏预览+放大镜"""
        if self.picking:
            return
        self.picking = True
        self.root.withdraw()

        # 全屏透明捕捉窗
        self.preview_window = tk.Toplevel(self.root)
        self.preview_window.attributes('-alpha', 0.01)
        self.preview_window.attributes('-fullscreen', True)
        self.preview_window.attributes('-topmost', True)
        self.preview_window.config(cursor='crosshair')
        self.preview_window.bind('<Motion>', self.live_preview)
        self.preview_window.bind('<Button-1>', self.select_color)
        self.preview_window.bind('<Escape>', self.cancel_pick)

        # 放大镜
        self.mag_window = tk.Toplevel(self.root)
        self.mag_window.overrideredirect(True)
        self.mag_window.attributes('-topmost', True)
        self.mag_win_size = 150
        self.zoom_factor = 10  # 放大倍数
        self.mag_label = tk.Label(self.mag_window)
        self.mag_label.pack()

    def live_preview(self, event=None):
        x, y = pyautogui.position()
        # 更新放大镜
        half = int(self.mag_win_size / self.zoom_factor / 2)
        left = x - half
        top = y - half
        right = x + half
        bottom = y + half
        img = ImageGrab.grab(bbox=(left, top, right, bottom))
        img = img.resize((self.mag_win_size, self.mag_win_size), Image.NEAREST)
        # 画十字
        draw = img.copy()
        pixels = draw.load()
        cx, cy = self.mag_win_size // 2, self.mag_win_size // 2
        for i in range(self.mag_win_size):
            pixels[cx, i] = (255, 0, 0)
            pixels[i, cy] = (255, 0, 0)
        self.magnifier_photo = ImageTk.PhotoImage(draw)
        self.mag_label.config(image=self.magnifier_photo)
        # 放大镜跟随鼠标
        self.mag_window.geometry(f"{self.mag_win_size}x{self.mag_win_size}+{x+20}+{y+20}")

        # 更新实时色块显示
        r, g, b = self.get_pixel(x, y)
        hx = f'#{r:02X}{g:02X}{b:02X}'
        self.color_canvas.config(bg=hx)
        self.hex_entry.delete(0, tk.END)
        self.hex_entry.insert(0, hx)
        self.rgb_entry.delete(0, tk.END)
        self.rgb_entry.insert(0, f"RGB: {r} {g} {b}")

    def get_pixel(self, x, y):
        sx, sy = int(x * self.scale_factor), int(y * self.scale_factor)
        return ImageGrab.grab().load()[sx, sy]

    def select_color(self, event=None):
        x, y = pyautogui.position()
        r, g, b = self.get_pixel(x, y)
        hx = f'#{r:02X}{g:02X}{b:02X}'
        self.update_color_displays(r, g, b, hx)
        self.finish_pick()
        self.root.clipboard_clear()
        self.root.clipboard_append(hx)

    def update_color_displays(self, r, g, b, hx):
        self.color_canvas.config(bg=hx)
        self.hex_entry.delete(0, tk.END)
        self.hex_entry.insert(0, hx)
        self.rgb_entry.delete(0, tk.END)
        self.rgb_entry.insert(0, f"RGB: {r} {g} {b}")

    def cancel_pick(self, event=None):
        self.finish_pick()

    def finish_pick(self):
        self.picking = False
        if hasattr(self, 'preview_window'):
            self.preview_window.destroy()
        if hasattr(self, 'mag_window'):
            self.mag_window.destroy()
        # 取色完成后主界面保持隐藏，除非手动打开

    def on_close(self):
        try:
            self.tray_icon.stop()
        except:
            pass
        keyboard.remove_all_hotkeys()
        self.root.destroy()
        sys.exit(0)

if __name__ == '__main__':
    root = tk.Tk()
    app = ColorPickerApp(root)
    root.mainloop()
