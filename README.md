### 自动给企业微信联系人发消息

### 使用前提
+ python3环境

### 依赖的库
+ pyautogui（模拟鼠标键盘操作）
+ pyperclip（复制剪切板）
+ xlrd（读取excel文件）
+ _thread（线程）
+ pynput（监听键盘操作）

### 库的安装方法
```bash
pip3 install pyautogui
pip3 install pyperclip
pip3 install xlrd==1.2.0
pip3 install pynput
```
#### 使用方法
##### 1.配置data.xlsx
【姓名】列为要发送的人
【消息】列卫要发送的内容

##### 执行python3 start.py
在脚本开始执行的5秒内切换到企业微信，并保持企业微信运行在前台