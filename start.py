import sys
import time
import pyautogui
import pyperclip
import platform
import xlrd
import _thread
from pynput.keyboard import Controller, Key, Listener

in_process = 1
# 监听按压
def on_press(key):
    print("已经按下:", format(key))


# 监听释放
def on_release(key):
    print("已经释放:", format(key))
    if key == Key.esc:
        print('stop')
        global in_process
        in_process = 0
        # 停止监听
        exit()

# 开始监听
def start_listen():
    with Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()

def get_import_data():
    data = xlrd.open_workbook('data.xlsx')
    return data.sheets()[0]
    
def send_message(user,msg):
    pyperclip.copy(user)
    hot_key('command','f')
    hot_key('command','v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyperclip.copy(msg)
    hot_key('command','v')
    pyautogui.press('enter')
    
def hot_key(key1,key2):
    if platform.system().lower() == 'windows':
        pyautogui.hotkey(key1,key2)
        # darwin is macos
    elif platform.system().lower() == 'darwin': 
        mac_hot_key(key1,key2)
        
def mac_hot_key(key1,key2):
    pyautogui.keyDown(key1)
    pyautogui.keyDown(key2)
    time.sleep(1)
    pyautogui.keyUp(key2)
    pyautogui.keyUp(key1)
    
def main():
    # start_listen()
    try:
        _thread.start_new_thread(start_listen, ())
    except:
        print ("Error: unable to start thread")
    table = get_import_data()
    nrows = table.nrows
    print('一共'+ str(nrows) + '条数据')
    print('请在5秒内打开企业微信窗口...')
    time.sleep(5)
    for n in range(1,nrows):
        print(in_process)
        if(in_process == 0):
            exit()
        row_data = table.row_values(n, start_colx=0, end_colx=None)
        send_message(row_data[0], row_data[1])
if __name__ == '__main__':
    main()