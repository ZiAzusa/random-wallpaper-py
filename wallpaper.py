# -*- mode: python ; coding: utf-8 -*-
import os, shutil, psutil, win32gui, win32con, win32api, requests
from tkinter import ttk
from tkinter import *
from tkinter.messagebox import *
from threading import Timer
from PIL import Image
from pystray import MenuItem, Icon

cancel = False
pause = False
minute = 15
sort = "pc"
ct = win32api.GetConsoleTitle()
hd = win32gui.FindWindow(0, ct) 
win32gui.ShowWindow(hd, 0)

window = Tk()
window.iconphoto(True, PhotoImage(file="./data/icon.ico"))
window.withdraw()

def main():
    global cancel, minute
    if cancel == True:
        cancel = True
        os._exit(0)
    if not pause:
        change_img()
    Timer(minute * 60, main).start()

def change_img(icon=False):
    if icon:
        icon.notify("壁纸手动更换成功~", "提示")
    api = "https://imgapi.nahida.xin/random?type=text&sort=" + sort
    path = "./data/pic/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE",
    }
    try:
        os.makedirs(path, exist_ok=True)
        url = str(requests.get(api, headers=headers).content, "utf-8")
        imgPath = path + os.path.basename(url)
        if os.path.exists(imgPath):
            if os.path.getsize(imgPath) <= 4096:
                os.remove(imgPath)
            else:
                set_wallpaper(os.path.abspath(imgPath))
                return
        res = requests.get(url)
        with open(imgPath, "wb") as f:
            f.write(res.content)
        if os.path.getsize(imgPath) <= 4096:
                os.remove(imgPath)
                showerror("错误", "图片下载失败")
        set_wallpaper(os.path.abspath(imgPath))
    except:
        showerror("错误", "更换壁纸失败")

def set_pause(icon=False):
    global pause
    if pause:
        change_img()
        icon.notify("已继续播放随机壁纸~", "提示")
    else:
        icon.notify("已暂停播放随机壁纸~", "提示")
    pause = not pause

def edit_minute(icon=False):
    global cancel, minute
    def on_submit():
        global minute
        try:
            minute = int(entry.get())
            if minute < 1:
                minute = 15
                showwarning("警告", "数据不合法，自动修改为默认值（15分钟）")
        except:
            minute = 15
            showwarning("警告", "数据不合法，自动修改为默认值（15分钟）")
        inputer.quit()
        inputer.destroy()
        showinfo("成功", "配置完成！程序将创建一个托盘图标，您可通过该图标切换壁纸或退出程序")
    def on_close():
        global cancel
        inputer.quit()
        inputer.destroy()
        showerror("错误", "您已手动退出配置流程")
        if not icon:
            cancel = True
            os.remove("./data/pid")
            os._exit(0)
    inputer = Tk()
    inputer.title("请输入壁纸切换周期（分钟）")
    screenWidth = inputer.winfo_screenwidth()
    screenHeight = inputer.winfo_screenheight()
    width = 360
    height = 60
    left = (screenWidth - width) / 2
    top = (screenHeight - height) / 2
    inputer.geometry("%dx%d+%d+%d" % (width, height, left, top))
    inputer.resizable(False, False)
    entry = Entry(inputer, width=40)
    entry.pack(side=LEFT)
    button = ttk.Button(inputer, text="提交", command=on_submit)
    button.pack(side=RIGHT)
    inputer.protocol("WM_DELETE_WINDOW", on_close)
    inputer.mainloop()

def clean_cache(icon=False):
    if askquestion("提示", "确定删除全部的壁纸缓存吗？此操作不可恢复"):
        shutil.rmtree("./data/pic")
        showinfo("成功", "已成功删除全部壁纸缓存")

def on_exit(icon=False):
    global cancel
    cancel = True
    icon.stop()
    os.remove("./data/pid")
    os._exit(0)

# style: 2拉伸, 0居中, 6适应, 10填充, 0平铺
def set_wallpaper(imgPath, style=10):
    if style == 0:
        tile = "1"
    else:
        tile = "0"
    regKey = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER, "Control Panel\\Desktop", 0, win32con.KEY_SET_VALUE)
    win32api.RegSetValueEx(regKey, "WallpaperStyle", 0, win32con.REG_SZ, str(style))
    win32api.RegSetValueEx(regKey, "TileWallpaper", 0, win32con.REG_SZ, tile)
    win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, imgPath, win32con.SPIF_SENDWININICHANGE)
    win32api.RegCloseKey(regKey)

if __name__ == "__main__":
    os.makedirs("./data", exist_ok=True)
    try:
        with open("./data/pid", "r") as f:
            if int(f.read()) in psutil.pids():
                cancel = True
                showerror("错误", "程序已在运行，请勿重复运行本程序")
                os._exit(0)
            raise Exception("Success")
    except:
        with open("./data/pid", "w") as f:
            f.write(str(os.getpid()))
    try:
        with open("./sort.txt", "r") as f:
            sort = f.read()
    except:
        pass
    edit_minute()
    main()
    menu = (
        MenuItem(text="更换壁纸", action=change_img),
        MenuItem(text="暂停/继续", action=set_pause),
        MenuItem(text="修改周期", action=edit_minute),
        MenuItem(text="删除缓存", action=clean_cache),
        MenuItem(text="退出", action=on_exit),
    )
    image = Image.open("./data/icon.ico")
    icon = Icon("name", image, "Nahida.Xin随机壁纸", menu)
    icon.run()
