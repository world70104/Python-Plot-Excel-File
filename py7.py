#---------------------------------------------------------------
#      Plot the excel data by tkinter & matplotlib & pandas
#
# Created: 2017/10/19
#---------------------------------------------------------------

# import module
from tkinter import *
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd

# tkinter create
root = Tk()

# tkinter frame size
root.geometry("%dx%d+%d+%d" % (500, 50, 0, 0))

# tkinter frame title
root.title("Device-Patient-Variable-Result Line")

# device, patient, varaibal, result index for select
deviceNum = 0
patientNum = 0
variableNum = 0
resultNum = 0

# menu iterms
choices_Device = ['Device 1', 'Device 2',  'Device 3']
choices_Patient = ['Patient 1', 'Patient 2']
choices_Variable = ['Variable 1', 'Variable 2']
choices_Result = ['Result 1', 'Result 2', 'Result 3', 'Result 4' ]

# initial value for select
varDevice = StringVar(root)
varPatient = StringVar(root)
varVariable = StringVar(root)
varResult = StringVar(root)
varDevice.set( choices_Device[ deviceNum ] )
varPatient.set( choices_Patient[ patientNum ] )
varVariable.set( choices_Variable[ variableNum ] )
varResult.set( choices_Result[ resultNum ] )

# real x,y number data in execel
allXList = []
allYList = []

excelFileName = "D://Sampledata.xlsx"

#-----------------------------------------
#      excel data parsing class
# out : allXList - real x number list
#       allYList - real y number list
#-----------------------------------------
class ExecelClass:
    # get the device data
    def getData(self):
        # load the execel file
        execelData = pd.ExcelFile( excelFileName )

        # get the sheet data
        sheetIdx = 0
        columCnt = 0
        while sheetIdx < 3:
            # get the data for one sheet
            data = execelData.parse(execelData.sheet_names[sheetIdx])
            # the number of rows
            rowCnt = len(data)
            # The column headings
            columCnt = len(data.columns)
            # get the real data
            idx = 0
            while idx < columCnt:
                realPlotData = data[data.columns[idx]]
                # get the real ploat number data only
                rIdx = 0
                # 4 - real data start index
                while rIdx < 4:
                    del realPlotData[ rIdx ]
                    rIdx += 1
                if idx % 2 == 0:
                   allXList.append( realPlotData )
                else:
                   allYList.append( realPlotData )
                idx = idx + 1
            sheetIdx += 1

# ---------------------------------------------------
#      Get the x, y data
# inp:
#       deviceNum: 0 ~ 2
#       patientNum: 0 -1
#       variableNum: 0 - 1
#       resultNum: 0 - 3
# out:
#       tuple index of x,y data
# ---------------------------------------------------
def GetXYData():
    global deviceNum, patientNum, variableNum, resultNum
    # get the index in data
    idexInData = deviceNum * 16 + patientNum * 8 + variableNum * 4 + resultNum

    return  allXList[ idexInData ], allYList[ idexInData ]

# ---------------------------------------------------
#      Plot the x, y data
#----------------------------------------------------
def plotXY ():
    # load the x,y for plot
    xx, yy = GetXYData()

    # pre plot cls
    plt.clf()

    # sub window title
    subTitle = "Device " + str( deviceNum + 1 ) + "- Patient " + str( patientNum + 1 ) + "- Variable " \
               + str( variableNum + 1 ) + "- Result " + str( resultNum + 1 )
    plt.suptitle (subTitle, fontsize=16)

    # plot x, y
    plt.plot( xx, yy, color='blue' )
    plt.ylabel("Y", fontsize=14)
    plt.xlabel("X", fontsize=14)
    plt.show()

# ---------------------------------------------------
#      Device menu event handler
#----------------------------------------------------
def OptionMenu_SelectionEvent_Device(event):
    global deviceNum
    if choices_Device[ deviceNum  ] != event:
        deviceNum = 0
        while( deviceNum < 3 ):
            if choices_Device[ deviceNum ] == event:
                break
            else:
                deviceNum += 1
        plotXY()
    pass

# ---------------------------------------------------
#      Patient menu event handler
#----------------------------------------------------
def OptionMenu_SelectionEvent_Patient(event):
    global patientNum
    if choices_Patient[ patientNum ] != event:
        patientNum = 0
        while( patientNum < 2 ):
            if choices_Patient[ patientNum ] == event:
                break
            else:
                patientNum += 1
        plotXY()
    pass

# ---------------------------------------------------
#      Variable menu event handler
#----------------------------------------------------
def OptionMenu_SelectionEvent_Variable(event):
    global variableNum
    if choices_Variable[ variableNum ] != event:
        variableNum = 0
        while( variableNum < 2 ):
            if choices_Variable[ variableNum ] == event:
                break
            else:
                variableNum += 1
        plotXY()
    pass

# ---------------------------------------------------
#      Result menu event handler
#----------------------------------------------------
def OptionMenu_SelectionEvent_Result(event):
    global resultNum
    if choices_Result[resultNum] != event:
        resultNum = 0
        while (variableNum < 4):
            if choices_Result[resultNum] == event:
                break
            else:
                resultNum += 1
        plotXY()
    pass

# Device menu make
optionDevice = OptionMenu(root, varDevice, *choices_Device, command = OptionMenu_SelectionEvent_Device )
optionDevice.pack(side='left', padx = 10, pady = 10)

# Patient menu make
optionPatient = OptionMenu(root, varPatient, *choices_Patient, command = OptionMenu_SelectionEvent_Patient )
optionPatient.pack(side='left', padx = 10, pady = 10)

# Variable menu make
optionVariable = OptionMenu(root, varVariable, *choices_Variable, command = OptionMenu_SelectionEvent_Variable )
optionVariable.pack(side='left', padx = 10, pady = 10)

# Result menu make
optionResult = OptionMenu(root, varResult, *choices_Result, command = OptionMenu_SelectionEvent_Result )
optionResult.pack(side='left', padx = 10, pady = 10)

# load the excel data
excelData = ExecelClass()
excelData.getData()
# plot
plotXY()

# event waiting
root.mainloop()

