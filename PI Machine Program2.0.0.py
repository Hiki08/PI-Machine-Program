# %%
from Imports import *
import DateAndTimeManager

# %%
isProgramRunning = True
isCompiling = True

isDebugMode = False

#Csv Data
allPIData = ""
piData = ""
tempPiData = ""
process5Data = ""
tempProcess5Data = ""

#Read Checker
isPiDataReaded = False
isProcess5DataReaded = False

#Date
dateToday = "2024/09/20"

#Output
csvData = ""
compiledData = ""

#Row To Read
piRow = 0
process5Row = 0

isCanProceed = True

#Process Type
isMasterPump = False
isTrial = False
isTrialRunning = False
isRunning = False
isNG = False
isGood = False
isNGPressure = False

readCount = 0

# %%
def settingFC1Directory():
    global isDebugMode

    if isDebugMode:
        return r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\Debug\log000_FC1.csv'
    else:
        return r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\log000_FC1.csv'

# %%
def settingProcess5Directory():
    global isDebugMode

    if isDebugMode:
        return fr'\\192.168.2.10\csv\csv\VT5\Debug\log000_5.csv'
    else:
        return fr'\\192.168.2.10\csv\csv\VT5\log000_5.csv'

# %%
def ReadPIMachine():
    global isDebugMode

    global dateToday

    global allPIData
    global piData

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    if isDebugMode:
        piDirectory = (r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1\Debug')
    else:
        piDirectory = (r'\\192.168.2.19\general\INSPECTION-MACHINE\FC1')
        
    os.chdir(piDirectory)
    
    allPIData = pd.read_csv('log000_FC1.csv', encoding='latin1', skiprows=1)
    
    allPIData.columns = [
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

    piData = allPIData[(allPIData["DATE"].isin([dateToday]))]

# %%
def ReadProcess5():
    global isDebugMode

    global dateToday

    global process5Data

    if isDebugMode:
        process5Directory = (r'\\192.168.2.10\csv\csv\VT5\Debug')
    else:
        process5Directory = (r'\\192.168.2.10\csv\csv\VT5')

    os.chdir(process5Directory)

    process5Data = pd.read_csv(f'log000_5.csv', encoding='latin1')
    process5Data = process5Data[(process5Data["DATE"].isin([dateToday]))]
    

# %%
def checkIfCanProceed():
    global isCompiling

    global allPIData
    global piData
    global tempPiData
    global process5Data
    global tempProcess5Data

    global piRow
    global process5Row

    global isCanProceed

    global isMasterPump
    global isTrial
    global isTrialRunning
    global isRunning
    global isNG
    global isGood
    global isNGPressure

    #Private Variables
    isPiDataBlank = True
    isProcess5DataBlank = True

    #Reseting Value
    isCanProceed = False

    isMasterPump = False
    isTrial = False
    isTrialRunning = False
    isRunning = False
    isNG = False
    isGood = False
    isNGPressure = False

    try:
        tempPiData = piData.iloc[[piRow], :]
        isPiDataBlank = False
        print("Pi Not Blank")
    except:
        isPiDataBlank = True
        isCanProceed = False
        isCompiling = False
        print("Pi Data Blank")
    try:
        tempProcess5Data = process5Data.iloc[[process5Row], :]
        tempProcess5Data["Process 5 NG Cause"]
        if tempProcess5Data["Process 5 NG Cause"].values[0].replace(' ', '') == "NGPRESSURE":
            isNGPressure = True
            isCanProceed = True
            print("NG Pressure")
    except:
        pass

    if not isPiDataBlank and not isNGPressure:
        if "M" in tempPiData["MODEL CODE"].values[0]:
            isMasterPump = True
            isCanProceed = True
            print("Master Pump")

        #Checking If Trial/TrialRunning
        elif len(str(tempPiData["S/N"].values[0])) < 9:
            print("Checking If Trial/TrialRunning")

            piDataFilteredGood = allPIData[(allPIData["PASS/NG"].isin([1])) & (allPIData["MODEL CODE"].isin([tempPiData["MODEL CODE"].values[0]]))]
            serialNumberList = piDataFilteredGood["S/N"].values

            sameSerialList = []
            for a in serialNumberList:
            #Checking S/N If Same Value Exists = Running
                if tempPiData["S/N"].values == a:
                    sameSerialList.append(a)
                if len(sameSerialList) > 5:
                    print("Trial Running")
                    isTrialRunning = True
                    isCanProceed = True
            print(f"Same Value Length is {len(sameSerialList)}")
            if not isTrialRunning:
                print("Trial")

                isTrial = True
                isCanProceed = True

        #Checking If NG/Running
        elif tempPiData["PASS/NG"].values == 0:
            print("Checking If NG/Running")

            piDataFilteredGood = allPIData[(allPIData["PASS/NG"].isin([1])) & (allPIData["MODEL CODE"].isin([tempPiData["MODEL CODE"].values[0]]))]
            serialNumberList = piDataFilteredGood["S/N"].values

            sameSerialList = []
            for a in serialNumberList:
            #Checking S/N If Same Value Exists = Running
                if tempPiData["S/N"].values == a:
                    sameSerialList.append(a)
                if len(sameSerialList) > 2:
                    print("Running")
                    isRunning = True
                    isCanProceed = True

            if not isRunning:
                print("NG")

                try:
                    tempProcess5Data = process5Data.iloc[[process5Row], :]
                    print("Process 5 Not Blank")
                    isNG = True
                    isCanProceed = True
                except:
                    isCanProceed = False
                    isCompiling = False
                    print("Process 5 Data Blank")

        #Checking If Good/Running
        elif tempPiData["PASS/NG"].values == 1:
            print("Checking If Good/Running")

            piDataFilteredGood = allPIData[(allPIData["PASS/NG"].isin([1])) & (allPIData["MODEL CODE"].isin([tempPiData["MODEL CODE"].values[0]]))]
            #Not Including the tempPiData
            # piDataFilteredGood = piDataFilteredGood[(piDataFilteredGood["TIME"].isin([tempPiData["TIME"].values[0]]))]
            serialNumberList = piDataFilteredGood["S/N"].values


            sameSerialList = []
            for a in serialNumberList:
            #Checking S/N If Same Value Exists = Running
                if tempPiData["S/N"].values == a:
                    sameSerialList.append(a)
                if len(sameSerialList) > 2:
                    print("Running")
                    isRunning = True
                    isCanProceed = True
            
            if not isRunning:
                print("Good")

                try:
                    tempProcess5Data = process5Data.iloc[[process5Row], :]
                    print("Process 5 Not Blank")
                    isGood = True
                    isCanProceed = True
                except:
                    isCanProceed = False
                    isCompiling = False
                    print("Process 5 Data Blank")

        else:
            isCanProceed = False
            isCompiling = False
            print("Reached The Last Row")

# %%
def compileCsv():
    global piData
    global tempPiData
    global process5Data
    global tempProcess5Data

    global piRow
    global process5Row

    global csvData
    global compiledData

    global isMasterPump
    global isTrial
    global isTrialRunning
    global isRunning
    global isNG
    global isGood
    global isNGPressure

    csvData = {
        "DATE": [tempPiData["DATE"].values[0]],
        "TIME": [tempPiData["TIME"].values[0]],
        "MODEL CODE": [tempPiData["MODEL CODE"].values[0]],
        "PROCESS S/N": "",
        "S/N": [tempPiData["S/N"].values[0]],
        "PASS/NG": [tempPiData["PASS/NG"].values[0]],
        "VOLTAGE MAX (V)": [tempPiData["VOLTAGE MAX (V)"].values[0]],
        "WATTAGE MAX (W)": [tempPiData["WATTAGE MAX (W)"].values[0]],
        "CLOSED PRESSURE_MAX (kPa)": [tempPiData["CLOSED PRESSURE_MAX (kPa)"].values[0]],
        "VOLTAGE Middle (V)": [tempPiData["VOLTAGE Middle (V)"].values[0]],
        "WATTAGE Middle (W)": [tempPiData["WATTAGE Middle (W)"].values[0]],
        "AMPERAGE Middle (A)": [tempPiData["AMPERAGE Middle (A)"].values[0]],
        "CLOSED PRESSURE Middle (kPa)": [tempPiData["CLOSED PRESSURE Middle (kPa)"].values[0]],
        "dB(A) 1": [tempPiData["dB(A) 1"].values[0]],
        "dB(A) 2": [tempPiData["dB(A) 2"].values[0]],
        "dB(A) 3": [tempPiData["dB(A) 3"].values[0]],
        "VOLTAGE MIN (V)": [tempPiData["VOLTAGE MIN (V)"].values[0]],
        "WATTAGE MIN (W)": [tempPiData["WATTAGE MIN (W)"].values[0]],
        "CLOSED PRESSURE MIN (kPa)": [tempPiData["CLOSED PRESSURE MIN (kPa)"].values[0]],
        "CHECKING": "-"
    }
    csvData = pd.DataFrame(csvData)

    print(piRow)
    
    if isMasterPump:
        print("Master Pump")
        csvData["PROCESS S/N"] = "MASTER PUMP"   
        piRow += 1
        compiledData = pd.concat([compiledData, csvData], ignore_index=True)
    elif isTrial:
        print("Trial")
        csvData["PROCESS S/N"] = "TRIAL"   
        piRow += 1
        compiledData = pd.concat([compiledData, csvData], ignore_index=True)
    elif isTrialRunning:
        print("Trial Running")
        csvData["PROCESS S/N"] = "TRIAL RUNNING"   
        piRow += 1
        compiledData = pd.concat([compiledData, csvData], ignore_index=True)
    elif isRunning:
        print("Running")
        csvData["PROCESS S/N"] = "RUNNING"   
        piRow += 1
        compiledData = pd.concat([compiledData, csvData], ignore_index=True)
    elif isNG:
        print("NG")
        csvData["PROCESS S/N"] = tempProcess5Data["Process 5 S/N"].values[0]
        piRow += 1
        process5Row += 1
        compiledData = pd.concat([compiledData, csvData], ignore_index=True)
    elif isGood:
        print("Good")
        csvData["PROCESS S/N"] = tempProcess5Data["Process 5 S/N"].values[0]
        piRow += 1
        process5Row += 1
        compiledData = pd.concat([compiledData, csvData], ignore_index=True)
    elif isNGPressure:
        print("NG PRESSURE")
        csvData["DATE"] = "NG PRESSURE" 
        csvData["TIME"] = "NG PRESSURE" 
        csvData["MODEL CODE"] = "NG PRESSURE" 
        csvData["PROCESS S/N"] = tempProcess5Data["Process 5 S/N"].values[0]
        csvData["S/N"] = "NG PRESSURE" 
        csvData["PASS/NG"] = "NG PRESSURE" 
        csvData["VOLTAGE MAX (V)"] = "NG PRESSURE" 
        csvData["WATTAGE MAX (W)"] = "NG PRESSURE" 
        csvData["CLOSED PRESSURE_MAX (kPa)"] = "NG PRESSURE" 
        csvData["VOLTAGE Middle (V)"] = "NG PRESSURE" 
        csvData["WATTAGE Middle (W)"] = "NG PRESSURE" 
        csvData["AMPERAGE Middle (A)"] = "NG PRESSURE" 
        csvData["CLOSED PRESSURE Middle (kPa)"] = "NG PRESSURE" 
        csvData["dB(A) 1"] = "NG PRESSURE" 
        csvData["dB(A) 2"] = "NG PRESSURE" 
        csvData["dB(A) 3"] = "NG PRESSURE" 
        csvData["VOLTAGE MIN (V)"] = "NG PRESSURE" 
        csvData["WATTAGE MIN (W)"] = "NG PRESSURE" 
        csvData["CLOSED PRESSURE MIN (kPa)"] = "NG PRESSURE" 
        csvData["CHECKING"] = "NG PRESSURE" 
        process5Row += 1
        compiledData = pd.concat([compiledData, csvData], ignore_index=True)

# %%
def WriteCsv(excelData):
    global isDebugMode

    global dateToday

    if isDebugMode:
        fileDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\Debug\PICompiled')
    else:
        fileDirectory = (r'\\192.168.2.19\ai_team\AI Program\Outputs\PICompiled')

    os.chdir(fileDirectory)
    print(os.getcwd())
    
    print("Creating New File")
    #Create Excel File
    newValue = pd.concat([excelData], axis = 0, ignore_index = True)
    wireFrame = newValue
    wireFrame.to_csv(f"PICompiled{dateToday.replace('/', '-')}.csv", index = False)

# %%
def startProgram():
    global isCompiling

    global csvData
    global compiledData 

    global isCanProceed
    
    col = [
        "DATE",
        "TIME",
        "MODEL CODE",
        "PROCESS S/N",
        "S/N",
        "PASS/NG",
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
        "CHECKING"
    ]
    compiledData = pd.DataFrame(columns=col)

    while isCompiling:
        # DateAndTimeManager.GetDateToday()
        ReadPIMachine()
        ReadProcess5()
        checkIfCanProceed()
        if isCanProceed:
            compileCsv()
    WriteCsv(compiledData)

    time.sleep(1)

# %%
#Normal Mode Or Debug Mode Prompt
isNormalOrDebug = input("1. Normal Mode, 2. Debug Mode\n")
if isNormalOrDebug == '1':
    print("Normal Mode Activated")
    isDebugMode = False
else:
    print("Debug Mode Activated")
    isDebugMode = True


DateAndTimeManager.GetDateToday()
dateToday = DateAndTimeManager.dateToday

#Checking Changes In CSV File Every Seconds

#Reading Original File
piDataOrig = os.path.getmtime(settingFC1Directory())
process5DataOrig = os.path.getmtime(settingProcess5Directory())

while isProgramRunning:
    DateAndTimeManager.GetDateToday()
    if dateToday != DateAndTimeManager.dateToday:
        DateAndTimeManager.GetDateToday()
        dateToday = DateAndTimeManager.dateToday

    #Checking Changes In PiCompiled Using Forced Reading Of Csv
    while not isPiDataReaded:
        try:
            piDataCurrent = os.path.getmtime(settingFC1Directory())
            isPiDataReaded = True
        except:
            print("Failed Reading Of PiCompiled Trying Again In 1 Seconds")
            pass
    isPiDataReaded = False

    while not isProcess5DataReaded:
        try:
            process5DataCurrent = os.path.getmtime(settingProcess5Directory())
            isProcess5DataReaded = True
        except:
            print("Failed Reading Of PiCompiled Trying Again In 1 Seconds")
            pass
    isProcess5DataReaded = False

    #If Changes Detected In File
    if piDataCurrent != piDataOrig or process5DataCurrent != process5DataOrig:
        print("Changes Detected")

        startProgram()

        #Setting The Original File Onto Current File
        piDataOrig = piDataCurrent
        process5DataOrig = process5DataCurrent

    print("Waiting For Changes In PiCompiled")
    
    #Clearing Cmd Logs When Reaches 10 Lines
    readCount += 1
    if readCount >= 10:
        os.system('cls')
        readCount = 0
    time.sleep(1)
    #_______________________________________



