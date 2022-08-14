from __future__ import annotations

import os
import openpyxl as xl
# from openpyxl import Workbook, load_workbook
import pyperclip as clp
import pyautogui
import time


wb = xl.load_workbook('data/email_list.xlsx')
ws = wb.active
lst = []
for r in ws.rows:
    if r[0].value is None:
        continue
    lst.append([])
    for c in r:
        lst[-1].append(c.value)
    print(lst[-1])
lst.pop(0)

dirname = os.path.dirname(os.path.abspath(__file__))    
write_btn_rel_path = 'data\\write_mail.jpg'
send_btn_rel_path = 'data\\btn_send.jpg'
write_btn_path = os.path.join(dirname, write_btn_rel_path)
send_btn_path = os.path.join(dirname, send_btn_rel_path)
print(write_btn_path)
print(send_btn_path)
for i in lst:
    write_btn_location = pyautogui.locateOnScreen(write_btn_path, confidence = 0.8)
    print(write_btn_location)
    write_btn_center = pyautogui.center(write_btn_location)
    pyautogui.click(write_btn_center[0], write_btn_center[1])
    time.sleep(1)
    clp.copy(i[-1])
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    clp.copy(f"{i[1]} 급여명세서 발송 테스트.")
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.hotkey('tab')
    clp.copy(
    f"""
    {i[1]}님, 수고하셨습니다.
    """)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    
    try:
        print(send_btn_path)
    except:
        pass
    send_btn_location = pyautogui.locateOnScreen(send_btn_path, confidence = 0.8)
    print(send_btn_location)
    send_btn_center = pyautogui.center(send_btn_location)
    pyautogui.click(send_btn_center[0], send_btn_center[1])
    time.sleep(3)