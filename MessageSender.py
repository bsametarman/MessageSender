# Written by bsametarman
# 28.08.2021

"""

    This program was made to send same message to persons from WhatsApp.
    Person numbers are taken from excel file.

    You have to log in WhatsApp web before running this program.
    "search.png" and "search2.png" pictures have to be same folder with MessageSender.py

"""

import webbrowser
from numpy import NAN
import pandas as pd
import pyautogui as pg
import time

def findPosition():
    if pg.locateOnScreen('./search.png'):
        location = pg.center(pg.locateOnScreen('./search.png'))
        x = location.x
        y = location.y
        return x, y
        
    elif pg.locateOnScreen('./search2.png'):
        location = pg.center(pg.locateOnScreen('./search2.png'))
        x = location.x
        y = location.y
        return x, y

def sendMessage(x, y):
    pg.click(x, y)
    pg.typewrite(str(number))
    time.sleep(2)
    pg.click(x, y + 140)
    time.sleep(1)
    pg.typewrite(message)
    pg.press('enter')
    time.sleep(1)
    pg.click(x, y)
    pg.hotkey('ctrl', 'a')
    pg.press('backspace')

# Your excel file path (Excel file should have 'Numbers' column)
print("Enter your excel file path...\n")
path = input()
if path == '':
    path = './numbers.xlsx'

# Gets the all rows from excel file
numbersRow = pd.read_excel(path, usecols=["Numbers"])

# Gets the number of rows
leng = pd.read_excel(path)
leng = len(leng.index)

print("Write the message that you want to send...\n")
message = input()

webbrowser.open("https://web.whatsapp.com")
time.sleep(10)
positions = findPosition()

for x in range(leng):
    number = int(numbersRow.values[x])
    sendMessage(positions[0], positions[1])