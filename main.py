import cv2
import pyautogui
import time
import pyperclip
import keyboard
import os
import json
import ctypes

def read():
    """读取设置文件"""
    with open("arg.json", "r", encoding="utf-8") as f1:
        i = json.load(f1)
    return i

def printf(i_dict, text="", end="\n"):
    """格式化输出信息"""
    for key, value in i_dict.items():
        print(f"'{key}': '{value}'")
    print(text, end=end)

def time_f(num):
    """规范时间格式"""
    if num < 10: return f"0{num}"
    else: return num

def compute_pos(upper_left, lower_right, m):
    """
    根据模式计算坐标元组
    :return 坐标元组
    """
    if m == 1:
        m_pos = [
            upper_left[0] + 60,
            upper_left[1] + 70
        ]
    elif m == 2:
        m_pos = [
            upper_left[0] + 20,
            upper_left[1] + 15
        ]
    else:
        m_pos = [
            int((upper_left[0] + lower_right[0]) / 2),
            int((upper_left[1] + lower_right[1]) / 2) + 10
        ]
    for i in [0, 1]: m_pos[i] = int(m_pos[i] * screen_scale)
    return m_pos[0], m_pos[1]

def get_windows_screen_scale():
    try:
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        dpi = user32.GetDpiForSystem()
        scale = round(dpi / 96.0, 2)
        return scale
    except Exception as e:
        print(f"ERROR: {e}")
        return 0.00

def get_click_pos(terminal, mode):
    """
    计算点击坐标
    :param terminal: 模板图像
    :param mode: 坐标计算模式
    :return: 坐标元组
    """
    localtime = time.localtime()
    filename = f"{localtime.tm_year % 100}{time_f(localtime.tm_mon)}{time_f(localtime.tm_mday)}_" \
               f"{time_f(localtime.tm_hour)}{time_f(localtime.tm_min)}{time_f(localtime.tm_sec)}.png"
    # 保存屏幕截图
    pyautogui.screenshot().resize(
        (
            int(int(info["width"]) / screen_scale),
            int(int(info["height"]) / screen_scale)
        )
    ).save(fr"screenshots\{filename}")
    print(f"Screenshot    {filename}")
    # 载入图像
    image = cv2.imread(fr"screenshots\{filename}")
    # 载入模板图像
    image_terminal = cv2.imread(fr"terminal\{terminal}")
    # 读取模板的宽度和高度
    height, width, channel = image_terminal.shape
    # 进行匹配
    result = cv2.matchTemplate(image, image_terminal, cv2.TM_SQDIFF)
    # 解析出区域左上角、右下角、点击位置的坐标
    upper_left = cv2.minMaxLoc(result)[2]
    lower_right = (upper_left[0] + width, upper_left[1] + height)
    pos = compute_pos(upper_left, lower_right, mode)
    print(f"Clicked       {pos}")
    return pos

def auto_click_1(avg):
    """点击方式 I  点击后粘贴文本"""
    pyautogui.click(avg[0], avg[1], button="left")
    time.sleep(float(info["auto_click_1"]))
    pyperclip.copy(info["name"])
    pyautogui.hotkey("ctrl", "v")

def auto_click_2(avg):
    """点击方式 II  仅点击坐标"""
    pyautogui.click(avg[0], avg[1], button="left")

def edit_info():
    """编辑模式 """
    time.sleep(1)
    print("编辑模式启动...")
    while True:
        key = input("key: ")
        if key == "break": break
        value = input("value: ")
        if value == "cancel": continue
        if key == "del":
            del info[value]
            continue
        info[key] = value
    with open("arg.json", "w", encoding="utf-8") as f2:
        json.dump(info, f2, indent=4, ensure_ascii=False)
    printf(info, "\n编辑模式关闭...")

def run():
    """主程序"""
    print("正在识别...")
    # 记录开始时间
    t1 = time.time()
    # 姓名
    click_pos = get_click_pos("01.png", 1)
    auto_click_1(click_pos)
    time.sleep(float(info["1"]))
    # 性别
    if info["sex"] == "m": click_pos = get_click_pos("02.png", 2)
    if info["sex"] == "f": click_pos = get_click_pos("02_0.png", 2)
    auto_click_2(click_pos)
    time.sleep(float(info["2"]))
    # 班级（点击下拉菜单）
    click_pos = get_click_pos("03.png", 1)
    auto_click_2(click_pos)
    time.sleep(float(info["3"]))
    # 班级（点击item）
    click_pos = get_click_pos("04.png", 2)
    auto_click_2(click_pos)
    time.sleep(float(info["4"]))
    # 向下滚动
    pyautogui.scroll(0 - int(info["scroll"]))
    time.sleep(float(info["5"]))
    # 选课
    click_pos = get_click_pos("05.png", 2)
    auto_click_2(click_pos)
    time.sleep(float(info["6"]))
    # 提交
    click_pos = get_click_pos("06.png", 3)
    auto_click_2(click_pos)
    time.sleep(float(info["7"]))
    # 确认提交
    click_pos = get_click_pos("07.png", 3)
    auto_click_2(click_pos)
    # 记录结束时间
    t2 = time.time()
    print(f"本次用时:     {t2 - t1} sec\n运行结束...")

def kill():
    """杀掉进程"""
    print("正在关闭...")
    keyboard.unhook_all()
    pid = os.getpid()
    print(f"PID: {pid}")
    time.sleep(3)
    os.system(f"taskkill /PID {pid} /T /F")

if __name__ == '__main__':
    # 读取配置数据
    info = read()
    # 物理分辨率
    screen_size = pyautogui.size()
    info["width"] = f"{screen_size.width}"
    info["height"] = f"{screen_size.height}"
    # 屏幕缩放比例
    screen_scale = get_windows_screen_scale()
    info["scale"] = f"{screen_scale}"
    with open("arg.json", "w", encoding="utf-8") as f3:
        json.dump(info, f3, indent=4, ensure_ascii=False)
    # 输出信息
    printf(info, f"\n屏幕分辨率: {screen_size}\n屏幕缩放比: {screen_scale}")
    # 快捷键
    keyboard.add_hotkey("alt+s", run)
    keyboard.add_hotkey("alt+C", kill)
    keyboard.add_hotkey("alt+i", edit_info)
    print('''
食用方法:
    Alt + S - 启动识别
    Alt + C - 结束程序
注意事项:
    1. 请尽量不要在非表单页面使用该代码
    2. 请将要识别的内容完整呈现在屏幕（至少将姓名、性别、班级三个信息呈现在页面上）
    3. 程序执行时，请勿使用键盘、鼠标等干扰程序执行
    4. 如果你使用的是多显示器，请将表单窗口置于主显示器后再开始识别
    5. 如果出现识别不准确的情况（例：修改显示器分辨率、缩放，切换主副显示器等），请重启应用
    6. 如果识别总时间过长，可以尝试降低主显示器分辨率，重启应用后重试
''')
    keyboard.wait()
