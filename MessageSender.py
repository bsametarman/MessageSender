# Written by bsametarman
# 28.08.2021

"""

    This program was made to send same message to persons from WhatsApp Web.
    Person numbers are taken from excel file.

    You have to login WhatsApp web before running this program.

"""

import webbrowser
import pandas as pd
import pyautogui as pg
import time

def sendMessage(phoneNumber :str, message :str):
    pg.hotkey('ctrl', 'l')
    pg.typewrite("https://web.whatsapp.com/send?phone=+" + phoneNumber + "&text=" + message)
    pg.press('enter')
    time.sleep(5)
    pg.press('enter')
    time.sleep(2)

def waitTimeCalculate(hour: int, minute: int):
    currentTime = time.localtime()
    if currentTime.tm_hour == 0:
        currentTime.tm_hour = 24
    currentSec = (currentTime.tm_hour * 3600) + (currentTime.tm_min * 60) + (currentTime.tm_sec)
    sendedTotalSec = (hour * 3600) + (minute * 60)

    totalSec = sendedTotalSec - currentSec
    if totalSec <= 0:
        totalSec = 86400 + totalSec
    
    print(f"Messages will be sent {totalSec} seconds later....")
    print(f"Please, do not close this program. Your browser will open in {totalSec} second...")

    time.sleep(totalSec)

    webbrowser.open("https://web.whatsapp.com")
    time.sleep(15)

def program():
        # Your excel file path (Excel file should have 'Numbers' column)
        print("Enter your excel file path... (you can drag your file here)\n")
        path = input()
        if path == '':
            path = './numbers.xlsx'

        # Gets the all rows from excel file
        if option == '1':
            numbersRow = pd.read_excel(path, usecols=["Numbers"])
        else:
            numbersRow = pd.read_excel(path, usecols=["Numbers"])
            messagesRow = pd.read_excel(path, usecols=["Messages"])

        # Gets the number of rows
        leng = pd.read_excel(path)
        leng = len(leng.index)

        if option == '1':
            print("Write the message that you want to send...\n")
            message = input()
        else:
            print("Messages will be taken from excel file... \n")

        print("Write the time that you want to send message... (Exmp: 12 36)\n")
        sendedTime = input()
        sendedTime = sendedTime.split()
        sendedTimeHour = int(sendedTime[0])

        if sendedTimeHour == 0:
            sendedTimeHour = 24
        sendedTimeMin = int(sendedTime[1])

        if sendedTimeHour > 24 or sendedTimeMin > 60:
            raise Exception("Time format is wrong, try again...")
        
        waitTimeCalculate(sendedTimeHour, sendedTimeMin)

        if option == '1':
            for x in range(leng):
                phoneNumber = int(numbersRow.values[x])
                sendMessage(str(phoneNumber), message)
        else:
            for x in range(leng):
                phoneNumber = int(numbersRow.values[x])
                message = str(messagesRow.values[x][0])
                sendMessage(str(phoneNumber), message)

def main():
    print("---Before using this program you have to be login whatsapp web--- \n")
    global option
    option = input("Choose a option from below... \n \n [1] Send same message \n [2] Send different messages \n")
    
    if option == '1':
        program()
    elif option == '2':
        program()
    else:
        raise Exception("Please choose 1 or 2")

if __name__ == "__main__":
    main()