# %%
from pathlib import Path
import shutil
import glob
import pandas as pd
import os
import numpy as np
import math
import openpyxl
import datetime
import time
from openpyxl.styles import Font
import pyttsx3
from python_calamine import CalamineWorkbook
import xlrd

#GUI
from tkinter import *
import tkinter as tk
from tkinter import ttk

#Fixing Blur UI
from ctypes import windll

# %%
piData = ""
process5Data = ""
dateToday = ""
dateTodayCsvFormat = ""
csvData = ""
compiledData = ""

processSNValue = ""

isMasterPumpOrRunning = False

isProcess5Reading = ""

isVT5Readed = False

isPiFinished = ""

seconds = 0
readCount = 0

# %%
def ReadPIMachine():
    global piData
    global process5Data
    global dateToday
    global dateTodayCsvFormat

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    piDirectory = (r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1')
    # piDirectory = (r'C:\Users\c.raniel\Documents')
    os.chdir(piDirectory)
    
    piData = pd.read_csv('log000_FC1.csv', encoding='latin1', skiprows=1)
    # piData = pd.read_csv('log000_FC1_1.csv', encoding='latin1', skiprows=1)
    piData.columns = [
        "DATE",
        "TIME",
        "MODEL CODE",
        "S/N",
        "PASS/NG",
        "¶",
        "VOLTAGE MAX (V)",
        "WATTAGE MAX (W)",
        "CLOSED PRESSURE_MAX (kPa)",
        "VOLTAGE Middle (V)",
        "WATTAGE Middle (W)",
        "AMPERAGE Middle (A)",
        "CLOSED PRESSURE Middle (kPa)",
        "dB(A) 1",
        "dB(A) 2",
        "dB(A) 3",
        "VOLTAGE MIN (V)",
        "WATTAGE MIN (W)",
        "CLOSED PRESSURE MIN (kPa)",
        "Middle züÊ",
        "Max züÊ",
        "INSULATION PASS/NG",
        "WITHSTAND VOLTAGE PASS/GO",
        "mL/min",
        "DM01293",
        "VOLTAGE MAX PASS/NG",
        "WATTAGE MAX PASS/NG",
        "CLOSED PRESSURE MAX PASS/NG",
        "Middle inhale Air volume",
        "MAX inhale Air volume",
        "VOLTAGE Middle PASS/NG",
        "WATTAGE Middle PASS/NG",
        "AMPERAGE Middle PASS/NG",
        "CLOSED PRESSURE Middle PASS/NG",
        "VOLTAGE MIN PASS/NG",
        "WATTAGE MIN PASS/NG",
        "CLOSED PRESSURE MIN PASS/NG",
        "Ø°¸TESTÊPASS/NG",
        "INSPECTED Q'TY",
        "PASSED Q'TY",
        "AMPERAGE MAX (A)",
        "PRESSURE MAX@(kPa)",
        "PRESSURE Middle (kPa)",
        "PRESSURE MIN (kPa)",
        "Min LEAK PRESSURE (kPa)",
        "Min LEAK TIME (sec)",
        "CLOSED VOLTAGE MAX (V)",
        "CLOSED AMPERAGE MAX (A)",
        "NG Q'TY",
        "CLOSED WATTERGE MAX (W)",
        "CLOSED VOLTAGE Middle (V)",
        "CLOSED AMPERAGE Middle (A)",
        "CLOSED WATTERGE Middle (W)",
        "AMPERAGE MIN (A)",
        "CLOSED VOLTAGE MIN (V)",
        "CLOSED AMPERAGE MIN (A)",
        "CLOSED WATTERGE MIN (W)",
        "DM01800",
        "S/N  SWAP",
        "ÄÞ×²ÊÞd³ügªèl",
    ]
    process5Directory = (r'\\192.168.2.10\csv\csv\VT5')
    os.chdir(process5Directory)

    try:
        process5Data = pd.read_csv(f'log000_5_{dateTodayCsvFormat}.csv', encoding='latin1')
    except:
        print


# %%
def GetDateToday():
    global dateToday
    global dateTodayCsvFormat

    dateToday = datetime.datetime.today()
    dateTodayCsvFormat = dateToday.strftime('%y%m%d')
    dateToday = dateToday.strftime('%Y/%m/%d')

# %%
def CompileCsv():
    global piData
    global csvData
    global compiledData

    global processSNValue

    global isMasterPumpOrRunning

    piData = piData[(piData["DATE"].isin([dateToday]))]
    piData = piData.tail(1)

    
    
    csvData = {
        "DATE": [piData["DATE"].values[0]],
        "TIME": [piData["TIME"].values[0]],
        "MODEL CODE": [piData["MODEL CODE"].values[0]],
        "PROCESS S/N": processSNValue,
        "S/N": [piData["S/N"].values[0]],
        "PASS/NG": [piData["PASS/NG"].values[0]],
        "VOLTAGE MAX (V)": [piData["VOLTAGE MAX (V)"].values[0]],
        "WATTAGE MAX (W)": [piData["WATTAGE MAX (W)"].values[0]],
        "CLOSED PRESSURE_MAX (kPa)": [piData["CLOSED PRESSURE_MAX (kPa)"].values[0]],
        "VOLTAGE Middle (V)": [piData["VOLTAGE Middle (V)"].values[0]],
        "WATTAGE Middle (W)": [piData["WATTAGE Middle (W)"].values[0]],
        "AMPERAGE Middle (A)": [piData["AMPERAGE Middle (A)"].values[0]],
        "CLOSED PRESSURE Middle (kPa)": [piData["CLOSED PRESSURE Middle (kPa)"].values[0]],
        "dB(A) 1": [piData["dB(A) 1"].values[0]],
        "dB(A) 2": [piData["dB(A) 2"].values[0]],
        "dB(A) 3": [piData["dB(A) 3"].values[0]],
        "VOLTAGE MIN (V)": [piData["VOLTAGE MIN (V)"].values[0]],
        "WATTAGE MIN (W)": [piData["WATTAGE MIN (W)"].values[0]],
        "CLOSED PRESSURE MIN (kPa)": [piData["CLOSED PRESSURE MIN (kPa)"].values[0]],
        "CHECKING": "-"
    }
    compiledData = pd.DataFrame(csvData)
    if isMasterPumpOrRunning:
        if compiledData["S/N"].values[0] == 201002523:
            print("Master")
            compiledData["PROCESS S/N"] = "MASTER PUMP"
        elif compiledData["S/N"].values[0] == 230600841:
            print("Master")
            compiledData["PROCESS S/N"] = "MASTER PUMP"
        else:
            print("Running")
            compiledData["PROCESS S/N"] = "RUNNING"
    else:
        compiledData["PROCESS S/N"] = process5Data["Process 5 S/N"].tail(1).values[0]
    

# %%
def WriteCsv(excelData):
    fileDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs')
    # fileDirectory = (r'C:\Users\c.raniel\Documents')
    os.chdir(fileDirectory)
    print(os.getcwd())

    if os.path.exists(f"{fileDirectory}/PICompiled.csv"):
            print("Overiting The Existing File")
            #Read Excel File
            existedExcel = pd.read_csv(f"PICompiled.csv")
            newValue = pd.concat([existedExcel, pd.DataFrame(excelData, index=[0])], axis = 0, ignore_index = True)
            wireFrame = newValue
            wireFrame.to_csv(f"PICompiled.csv", index = False)
            
    
    else:
        print("Creating New File")
        #Create Excel File
        newValue = pd.concat([excelData], axis = 0, ignore_index = True)
        wireFrame = newValue
        wireFrame.to_csv(f"PICompiled.csv", index = False)

# %%
piOrigFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1.csv')
# piOrigFile = os.path.getmtime(r'C:\Users\hpi.python\Documents\log000_FC1_1.csv')
GetDateToday()

#Reading The Time Of VT5 CSV
try:
    process5OrigFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5_{dateTodayCsvFormat}.csv')
    isVT5Readed = True
except:
    print("No VT5 Data Readed")
    isVT5Readed = False

while True:
    try:
        try:
            piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1.csv')
        except:
            print
        # piCurrentFile = os.path.getmtime(r'C:\Users\hpi.python\Documents\log000_FC1_1.csv')

        try:
            process5CurrentFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5_{dateTodayCsvFormat}.csv')
        except:
            print

        # if piCurrentFile != piOrigFile:
        #     isMasterPumpOrRunning = True
            
        #     print("Data Created")
        #     ReadPIMachine()
        #     GetDateToday()
        #     CompileCsv()

        #     time.sleep(5)
        #     WriteCsv(compiledData)
        #     piOrigFile = piCurrentFile
        #     if isVT5Readed:
        #         process5OrigFile = process5CurrentFile
        #     # time.sleep(1)

        if piCurrentFile != piOrigFile:
            isProcess5Reading = False
            while seconds < 50:
                seconds += 1

                try:
                    process5CurrentFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5_{dateTodayCsvFormat}.csv')
                except:
                    print

                if process5CurrentFile != process5OrigFile:
                    isMasterPumpOrRunning = False

                    print("Data Created")
                    ReadPIMachine()
                    GetDateToday()
                    CompileCsv()

                    time.sleep(5)
                    WriteCsv(compiledData)
                    isProcess5Reading = True        
                    time.sleep(5)
                    piOrigFile = piCurrentFile
                    process5OrigFile = process5CurrentFile
                    break
                    
                time.sleep(1)

            seconds = 0

            
            if isProcess5Reading == False:
                isMasterPumpOrRunning = True
                
                print("Data Created")
                ReadPIMachine()
                GetDateToday()
                CompileCsv()

                time.sleep(5)
                WriteCsv(compiledData)
                time.sleep(5)
                piOrigFile = piCurrentFile
        
        # elif isVT5Readed and process5CurrentFile != process5OrigFile:
        #     # time.sleep(3)
        #     while process5CurrentFile != process5OrigFile:
        #         print("Waiting For PI Data")
        #         try:
        #             piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1.csv')
        #             # piCurrentFile = os.path.getmtime(r'C:\Users\hpi.python\Documents\log000_FC1_1.csv')
        #         except:
        #             print("Retrying")
        #         if piCurrentFile != piOrigFile:
        #             isMasterPumpOrRunning = False

        #             print("Data Created")
        #             ReadPIMachine()
        #             GetDateToday()
        #             CompileCsv()

        #             time.sleep(5)
        #             WriteCsv(compiledData)
        #             isProcess5Reading = True        
        #             piOrigFile = piCurrentFile
        #             process5OrigFile = process5CurrentFile
        #         time.sleep(.5)

        #Reading The Time Of VT5 CSV
        if not isVT5Readed:
            try:
                process5OrigFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5_{dateTodayCsvFormat}.csv')
                isVT5Readed = True
            except:
                print("No VT5 Data Readed")
                isVT5Readed = False
    except:
        print("Failure in system reading, retrying in 1 second")
    print("Reading PI Machine And VT5 Data")

    readCount += 1
    if readCount >= 10:
        os.system('cls')
        readCount = 0
    time.sleep(1)
        


