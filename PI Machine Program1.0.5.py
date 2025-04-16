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
tempPiData = ""
piData = ""
process5Data = ""
process5DataLastValue = ""
dateToday = ""
dateTodayCsvFormat = ""
csvData = ""
compiledData = ""

processSNValue = ""

isMasterPump = False
isRunning = False
isNG = False
isGood = False
isNGPressure = False

isProcess5Reading = ""

isVT5Readed = False

isPiFinished = ""

seconds = 0
readCount = 0

isPiDataReaded = ""

piDataFilteredGood = ""
serialNumberList = ""

isCheckedSerialValue = False
isProcess5DataReaded = False

isPiCompiledReaded = False
piCompiledLastTimeValue = ""
piMachineLastTimeValue = ""

# %%
def ReadPIMachine():
    global piData
    global dateToday
    global dateTodayCsvFormat

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    piDirectory = (r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1')
    # piDirectory = (r'C:\Users\c.raniel\Documents')
    os.chdir(piDirectory)
    
    piData = pd.read_csv('log000_FC1Trial.csv', encoding='latin1', skiprows=1)
    # piData = pd.read_csv('log000_FC1Trial.csv', encoding='latin1', skiprows=1)
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

# %%
def ReadPiLastTimeValue():
    global piCompiledLastTimeValue

    piDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs')
    os.chdir(piDirectory)
    
    piCompiledLastTimeValue = pd.read_csv('PICompiled6.csv', encoding='latin1')
    piCompiledLastTimeValue = piCompiledLastTimeValue["TIME"].tail(1).values

# %%
def ReadProcess5():
    global process5Data

    process5Directory = (r'\\192.168.2.10\csv\csv\VT5')
    os.chdir(process5Directory)

    process5Data = pd.read_csv(f'log000_5.csv', encoding='latin1')
    

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

    global isMasterPump
    global isRunning
    global isNG
    global isGood

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
    if isMasterPump:
        print("Master Pump")
        compiledData["PROCESS S/N"] = "MASTER PUMP"    
    elif isRunning:
        print("Running")
        compiledData["PROCESS S/N"] = "RUNNING"
    elif isNG:
        print("NG")
        compiledData["PROCESS S/N"] = process5Data["Process 5 S/N"].tail(1).values[0]
    elif isGood:
        print("Good")
        compiledData["PROCESS S/N"] = process5Data["Process 5 S/N"].tail(1).values[0]
    elif isNGPressure:
        csvData = {
        "DATE": ["NG PRESSURE"],
        "TIME": ["NG PRESSURE"],
        "MODEL CODE": ["NG PRESSURE"],
        "PROCESS S/N": processSNValue,
        "S/N": ["NG PRESSURE"],
        "PASS/NG": ["NG PRESSURE"],
        "VOLTAGE MAX (V)": ["NG PRESSURE"],
        "WATTAGE MAX (W)": ["NG PRESSURE"],
        "CLOSED PRESSURE_MAX (kPa)": ["NG PRESSURE"],
        "VOLTAGE Middle (V)": ["NG PRESSURE"],
        "WATTAGE Middle (W)": ["NG PRESSURE"],
        "AMPERAGE Middle (A)": ["NG PRESSURE"],
        "CLOSED PRESSURE Middle (kPa)": ["NG PRESSURE"],
        "dB(A) 1": ["NG PRESSURE"],
        "dB(A) 2": ["NG PRESSURE"],
        "dB(A) 3": ["NG PRESSURE"],
        "VOLTAGE MIN (V)": ["NG PRESSURE"],
        "WATTAGE MIN (W)": ["NG PRESSURE"],
        "CLOSED PRESSURE MIN (kPa)": ["NG PRESSURE"],
        "CHECKING": "-"
        }
        compiledData = pd.DataFrame(csvData)
        compiledData["PROCESS S/N"] = process5Data["Process 5 S/N"].tail(1).values[0]
    

# %%
def WriteCsv(excelData):
    fileDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs')
    # fileDirectory = (r'C:\Users\c.raniel\Documents')
    os.chdir(fileDirectory)
    print(os.getcwd())

    if os.path.exists(f"{fileDirectory}/PICompiled6.csv"):
            print("Overiting The Existing File")
            #Read Excel File
            existedExcel = pd.read_csv(f"PICompiled6.csv")
            newValue = pd.concat([existedExcel, pd.DataFrame(excelData, index=[0])], axis = 0, ignore_index = True)
            wireFrame = newValue
            wireFrame.to_csv(f"PICompiled6.csv", index = False)
            
    
    else:
        print("Creating New File")
        #Create Excel File
        newValue = pd.concat([excelData], axis = 0, ignore_index = True)
        wireFrame = newValue
        wireFrame.to_csv(f"PICompiled6.csv", index = False)

# %%
piOrigFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1Trial.csv')
process5OrigFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5.csv')
GetDateToday()

while True:
    try:
        #Checking Changes In PI File
        try:
            piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1Trial.csv')
        except:
            print
        #Checking Changes In Process 5 File
        try:
            process5CurrentFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5.csv')
        except:
            print

        if piCurrentFile != piOrigFile:
            print("Changes Detected")

            #Reset Variable Values
            isMasterPump = False
            isRunning = False
            isNG = False
            isGood = False
            isNGPressure = False
            isPiDataReaded = False
            isProcess5Reading = False
            isVT5Readed = False
            checkingSerialValue = False
            isProcess5DataReaded = False
            isPiCompiledReaded = False
            #__________________________________

            #Force Reading Last Value Of PI Data
            while not checkingSerialValue:
                try:
                    ReadPIMachine()
                    tempPiData = piData.tail(1)
                    checkingSerialValue = True
                except:
                    print
            #____________________________________

            #Checking If Master Pump
            if "M" in tempPiData["MODEL CODE"].values[0]:
                #Force Reading PiCompiled, If Not Exist Proceed
                while not isPiCompiledReaded:
                    try:
                        ReadPiLastTimeValue()
                        isPiCompiledReaded = True
                    except FileNotFoundError:
                        isPiCompiledReaded = True
                        print("Proceed")
                    except:
                        print("Repeat")

                if piCompiledLastTimeValue == tempPiData["TIME"].values:
                    print
                    while not isPiDataReaded:
                        try:
                            piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1Trial.csv')
                            isPiDataReaded = True
                        except:
                            print

                    piOrigFile = piCurrentFile
                else:
                    isMasterPump = True
                    ReadPIMachine()
                    GetDateToday()
                    CompileCsv()
                    WriteCsv(compiledData)
                    time.sleep(10)

                    while not isPiDataReaded:
                        try:
                            piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1Trial.csv')
                            isPiDataReaded = True
                        except:
                            print

                    piOrigFile = piCurrentFile

                    isMasterPump = False
                    isRunning = False
                    isNG = False
                    isGood = False
                    isPiDataReaded = False
                    isProcess5Reading = False
            #__________________________________________________________________________________________________

            #Checking If NG/Running
            elif tempPiData["PASS/NG"].values == 0:
                print("Checking If NG/Running")
                ReadPIMachine()
                piDataFilteredGood = piData[(piData["PASS/NG"].isin([1]))]

                serialNumberList = piDataFilteredGood["S/N"].values

                for a in serialNumberList[:-1]:
                    #Checking S/N If Same Value Exists = Running
                    if tempPiData["S/N"].values == a:
                        print("Same Value Exists")
                        isRunning = True

                        ReadPIMachine()
                        GetDateToday()
                        CompileCsv()
                        WriteCsv(compiledData)
                        time.sleep(10)

                        while not isPiDataReaded:
                            try:
                                piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1Trial.csv')
                                isPiDataReaded = True
                            except:
                                print

                        piOrigFile = piCurrentFile
                        # process5OrigFile = process5CurrentFile

                        isMasterPump = False
                        isRunning = False
                        isNG = False
                        isGood = False
                        isPiDataReaded = False
                        isProcess5Reading = False
                        break

                if not isRunning:
                    isNG = True
                    while not isVT5Readed:
                        try:
                            process5OrigFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5.csv')
                        except:
                            print
                        try:
                            if process5CurrentFile != process5OrigFile:
                                time.sleep(6)
                                
                                while not isProcess5DataReaded:
                                    try:
                                        ReadProcess5()
                                        isProcess5DataReaded = True
                                    except:
                                        print

                                ReadPIMachine()
                                GetDateToday()
                                CompileCsv()
                                WriteCsv(compiledData)
                                time.sleep(10)

                                while not isPiDataReaded:
                                    try:
                                        piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1Trial.csv')
                                        isPiDataReaded = True
                                    except:
                                        print
                                while not isVT5Readed:
                                    try:
                                        process5CurrentFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5.csv')
                                        isVT5Readed = True
                                    except:
                                        print

                                piOrigFile = piCurrentFile
                                process5OrigFile = process5CurrentFile

                                isMasterPump = False
                                isRunning = False
                                isNG = False
                                isGood = False
                                isPiDataReaded = False
                                isProcess5Reading = False
                        except:
                            print
                            
                        print("Waiting For Process 5 Data")
                        time.sleep(1)
            #_____________________________________________________________________________________________________________________
            
            # Checking If Good/Running
            elif tempPiData["PASS/NG"].values == 1:
                print("Checking If Good/Running")
                print(isRunning)
                ReadPIMachine()
                piDataFilteredGood = piData[(piData["PASS/NG"].isin([1]))]

                serialNumberList = piDataFilteredGood["S/N"].values

                for a in serialNumberList[:-1]:
                    #Checking S/N If Same Value Exists = Running
                    if tempPiData["S/N"].values == a:
                        print("Same Value Exists")
                        isRunning = True

                        ReadPIMachine()
                        GetDateToday()
                        CompileCsv()
                        WriteCsv(compiledData)
                        time.sleep(10)

                        while not isPiDataReaded:
                            try:
                                piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1Trial.csv')
                                isPiDataReaded = True
                            except:
                                print

                        piOrigFile = piCurrentFile
                        # process5OrigFile = process5CurrentFile

                        isMasterPump = False
                        isRunning = False
                        isNG = False
                        isGood = False
                        isPiDataReaded = False
                        isProcess5Reading = False
                        break
                #If Not Running = Process Good OR NG Pressure
                if not isRunning:
                    while not isVT5Readed:
                        try:
                            process5CurrentFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5.csv')
                        except:
                            print

                        try:
                            if process5CurrentFile != process5OrigFile:
                                time.sleep(6)

                                while not isProcess5DataReaded:
                                    try:
                                        ReadProcess5()
                                        isProcess5DataReaded = True
                                    except:
                                        print

                                #Checking If NG PRESSURE - NO Data Write In FC1 LOG
                                process5DataLastValue = process5Data.tail(1)
                                if process5DataLastValue["Process 5 NG Cause"].values == "NG PRESSURE":
                                    isNGPressure = True
                                else:
                                    isGood = True
                                
                                ReadPIMachine()
                                GetDateToday()
                                CompileCsv()
                                WriteCsv(compiledData)
                                time.sleep(10)

                                while not isPiDataReaded:
                                    try:
                                        piCurrentFile = os.path.getmtime(r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1Trial.csv')
                                        isPiDataReaded = True
                                    except:
                                        print
                                while not isVT5Readed:
                                    try:
                                        process5CurrentFile = os.path.getmtime(fr'\\192.168.2.10\csv\csv\VT5\log000_5.csv')
                                        isVT5Readed = True
                                    except:
                                        print

                                piOrigFile = piCurrentFile
                                process5OrigFile = process5CurrentFile

                                isMasterPump = False
                                isRunning = False
                                isNG = False
                                isGood = False
                                isPiDataReaded = False
                                isProcess5Reading = False
                        except:
                            print
                            
                        print("Waiting For Process 5 Data")
                        time.sleep(1)
                

    except:
        print("Failure in system reading, retrying in 1 second")
    print("Reading PI Machine And VT5 Data")

    #Clearing Cmd Logs When Reaches 10 Lines
    readCount += 1
    if readCount >= 10:
        os.system('cls')
        readCount = 0
    time.sleep(1)
    #_______________________________________
        


