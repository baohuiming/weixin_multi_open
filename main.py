# -*- encoding: utf-8 -*-
"""
@File    :   main.py
@Time    :   2024/01/16 14:06:11
@Author  :   Loopade
@Version :   1.0
@Desc    :   微信多开工具
"""
import os
import win32api
import win32con
import win32gui
import psutil
import time
import json
import shutil


def is_weixin_running():
    # 判断微信进程是否存在
    pl = psutil.pids()
    for pid in pl:
        if psutil.Process(pid).name() == "WeChat.exe":
            print("weixin is running:", pid)
            return pid
    else:
        print("weixin is not running")
        return False


def is_weixin_logined():
    # 判断微信是否登录
    pl = psutil.pids()
    for pid in pl:
        if psutil.Process(pid).name() == "WeChatUtility.exe":
            print("微信已登录：", pid)
            return True
    else:
        print("微信未登录")
        return False


# 读取注册表找到微信的安装路径
def get_weixin_install_path():
    try:
        # 注册表打开
        # RegOpenKey(key, subKey , reserved , sam)
        # key: HKEY_CLASSES_ROOT HKEY_CURRENT_USER HEKY_LOCAL_MACHINE HKEY_USERS HKEY_CURRENT_CONFIG
        # subkey: 要打开的子项
        # reserved: 必须为0
        # sam: 对打开的子项进行的操作,包括win32con.KEY_ALL_ACCESS、win32con.KEY_READ、win32con.KEY_WRITE等
        key = win32api.RegOpenKey(
            win32con.HKEY_CURRENT_USER,
            "SOFTWARE\Tencent\WeChat",
            0,
            win32con.KEY_ALL_ACCESS,
        )
        # 这里的key表示键值，后面是具体的键名，读取出来是个tuple
        value = win32api.RegQueryValueEx(key, "InstallPath")[0]
        # 用完之后记得关闭
        win32api.RegCloseKey(key)
        # 微信的路径
        value += "\\" + "WeChat.exe"
        return value
    except Exception as ex:
        print("error:", ex)


def get_weixin_files_path():
    try:
        # %AppData%\Tencent\WeChat\All Users\config\3ebffe94.ini
        ini_path = (
            os.getenv("APPDATA") + "\\Tencent\\WeChat\\All Users\\config\\3ebffe94.ini"
        )
        with open(ini_path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as ex:
        print("error:", ex)


def kill_weixin():
    os.system("taskkill /f /im wechat.exe")


def get_user_name(weixin_files_path, data_path):
    # weixin_files_path\WeChat Files\wxid_x2d2ocpyzyn722\config\AccInfo.dat
    with open(data_path, "r", errors="ignore") as f:
        data = f.read()
        return data.split(weixin_files_path + "\\WeChat Files\\")[1].split(
            "\\config\\"
        )[0]


def start_weixin():
    # refer to https://github.com/anhkgg/SuperWeChatPC
    # 隐藏式启动anhkgg.exe
    os.system("start /b anhkgg.exe")


def reset_window_pos(i: int, n: int):
    targetTitle = "微信"
    hWndList = []
    win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)
    for hwnd in hWndList[::-1]:
        # clsname = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        if title == targetTitle:
            # 获取目标窗口尺寸
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            h = bottom - top
            w = right - left
            # 获取屏幕尺寸
            sh = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            sw = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            # 并排放置在屏幕中央
            total_w = int(w * n)
            total_h = h

            padding_left = int((sw - total_w) / 2)
            padding_top = int((sh - total_h) / 2)
            # 第i个
            target_left = padding_left + int(w * i)
            target_top = padding_top
            # 设置窗口位置
            win32gui.SetWindowPos(
                hwnd,
                win32con.HWND_TOPMOST,  # 置顶
                target_left,
                target_top,
                w,
                h,
                win32con.SWP_NOSIZE | win32con.SWP_NOMOVE,
            )
            # 移动窗口
            win32gui.MoveWindow(hwnd, target_left, target_top, w, h, True)
            # 点击窗口
            btn_x = int(target_left + w / 2)
            btn_y = int(target_top + h * 0.74)
            win32api.SetCursorPos([btn_x, btn_y])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            return


def setup():
    # 微信的安装路径
    weixin_install_path = get_weixin_install_path()
    # 微信的文件路径
    weixin_files_path = get_weixin_files_path()
    config = {
        "weixin_install_path": weixin_install_path,
        "weixin_files_path": weixin_files_path,
    }
    users = []
    # 判断微信是否运行
    if is_weixin_running():
        kill_weixin()
        time.sleep(1)
    if not os.path.exists("config.json"):
        # 首次使用，需添加微信
        for i in range(2):
            print(f"请添加第{i+1}个微信并登录")
            os.startfile(weixin_install_path)
            time.sleep(10)
            while not is_weixin_logined():
                time.sleep(1)
            kill_weixin()
            # 复制文件：weixin_files_path / WeChat Files / All Users / config / config.data
            shutil.copyfile(
                os.path.join(
                    weixin_files_path,
                    "WeChat Files",
                    "All Users",
                    "config",
                    "config.data",
                ),
                f"config{i}.data",
            )
            users.append(
                {
                    "user_name": get_user_name(weixin_files_path, f"config{i}.data"),
                    "config_path": f"config{i}.data",
                }
            )
        config["users"] = users
        with open("config.json", "w") as f:
            f.write(json.dumps(config))


def run():
    with open("config.json", "r") as f:
        config = json.loads(f.read())
    for i, user in enumerate(config["users"]):
        print("正在启动：", user["user_name"])
        # 先复制文件
        shutil.copyfile(
            user["config_path"],
            os.path.join(
                config["weixin_files_path"],
                "WeChat Files",
                "All Users",
                "config",
                "config.data",
            ),
        )
        start_weixin()
        time.sleep(1)
        reset_window_pos(i, len(config["users"]))


if __name__ == "__main__":
    setup()
    run()
