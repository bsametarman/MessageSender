# Written by bsametarman
# 28.08.2021

"""

    This program was made to send same message to persons from WhatsApp.
    Person numbers are taken from excel file.

    You have to log in WhatsApp web before running this program.
    "search.png" and "search2.png" pictures have to be same folder with MessageSender.py

"""

import webbrowser
import pandas as pd
import pyautogui as pg
import time

def sendMessage():
    pg.hotkey('ctrl', 'l')
    pg.typewrite("https://web.whatsapp.com/send?phone=+" + str(number) + "&text=" + message)
    pg.press('enter')
    time.sleep(5)
    pg.click(screenSize[0]/2, screenSize[1]/2)
    pg.press('enter')
    time.sleep(2)

def waitTimeCalculate(hour: int, minute: int):

    webbrowser.open("https://web.whatsapp.com")

    currentTime = time.localtime()
    if currentTime.tm_hour == 0:
        currentTime.tm_hour = 24
    currentSec = (currentTime.tm_hour * 3600) + (currentTime.tm_min * 60) + (currentTime.tm_sec)
    sendedTotalSec = (hour * 3600) + (minute * 60)

    totalSec = sendedTotalSec - currentSec
    if totalSec <= 0:
        left_time = 86400 + totalSec
    
    print(f"Messages will be sent {totalSec} Secounds later...")

    time.sleep(totalSec)


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

# Gets screen resulotion
screenSize = pg.size()

print("Write the message that you want to send...\n")
message = input()

print("Write the time that you want to send message...\n")
sendedTime = input()
sendedTime = sendedTime.split()
sendedTimeHour = int(sendedTime[0])
if sendedTimeHour == 0:
    sendedTimeHour = 24
sendedTimeMin = int(sendedTime[1])
waitTimeCalculate(sendedTimeHour, sendedTimeMin)

for x in range(leng):
    number = int(numbersRow.values[x])
    sendMessage()