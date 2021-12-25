#Imports
import os
from numpy import number
import pandas as pd
import datetime
import csv
from pathlib import Path
import re
import xlsxwriter
import xlrd
from xlsxwriter.workbook import Workbook
import numpy as np
import sys

##Functions-------------
def updateLine(exerciseName, chosenLine):

    #need to append to corresponding exercise line array    
    if(exerciseName == "Bench Press (Barbell)" or exerciseName == "Bench Press" or exerciseName == "Supino Deitada" 
    or exerciseName == "Incline Bench Press (Barbell)" or exerciseName == "Supino inclinado" or exerciseName == "Bench Press (Dumbbell)" ):
        bpLine.append(chosenLine + 1)
    
    elif(exerciseName == "Squat (Barbell)" or exerciseName == "Squat (Machine)" or exerciseName == "Goblet Squat" 
        or exerciseName == "Leg Press" or exerciseName == "Agachamento"):
        squatLine.append(chosenLine + 1)

    elif(exerciseName == "Rowing (Machine)" or exerciseName == "Bent Over Row (Barbell)" or exerciseName == "Remada"):
        rowLine.append(chosenLine + 1)

    elif(exerciseName == "Pull Up" or exerciseName == "Chin Up" or exerciseName == "Pull Up (Band)"):
        pullUpLine.append(chosenLine + 1)

    elif(exerciseName == "Shoulder Press (Machine)" or exerciseName =="Seated Overhead Press (Dumbbell)"):
        spLine.append(chosenLine + 1)

    elif(exerciseName == "Stiff Leg Deadlift (Barbell)" or exerciseName == "Stiff Leg Deadlift (Dumbbell)" or exerciseName == "Peso morto stiff"):
        stiffLine.append(chosenLine + 1)

    elif(exerciseName == "Hip Thrust (Barbell)"):
        hpLine.append(chosenLine + 1)

#Transforms similiar movement in valid exercise Name
def exerciseNameVerifier(exerciseName):

    if(exerciseName == "\"Bench Press (Barbell)\"" or exerciseName == "\"Bench Press\"" or exerciseName == "\"Supino Deitada\"" or exerciseName == "\"Incline Bench Press (Barbell)\""  
        or exerciseName == "\"Supino inclinado\"" or exerciseName == "\"Bench Press (Dumbbell)\"" or exerciseName == "Bench Press (Barbell)" or exerciseName == "Bench Press" or exerciseName == "Supino Deitada" or exerciseName == "Incline Bench Press (Barbell)"  
        or exerciseName == "Supino inclinado" or exerciseName == "Bench Press (Dumbbell)"):

         exerciseName = "Bench Press (Barbell)"
         return exerciseName
    
    elif(exerciseName == "\"Squat (Barbell)\"" or exerciseName == "\"Squat (Machine)\"" or exerciseName == "\"Goblet Squat\""  or exerciseName == "\"Leg Press\"" 
        or exerciseName =="\"Agachamento\"" or exerciseName == "\"Goblet Squat (Kettlebell)\"" or exerciseName == "Squat (Barbell)" or exerciseName == "Squat (Machine)" or exerciseName == "Goblet Squat"  or exerciseName == "Leg Press" 
        or exerciseName =="Agachamento" or exerciseName == "Goblet Squat (Kettlebell)" or exerciseName == "\"Globet Squat\"" or exerciseName == "Globet Squat"):

         exerciseName = "Squat (Barbell)"
         return exerciseName

    elif(exerciseName == "\"Rowing (Machine)\"" or exerciseName == "\"Bent Over Row (Barbell)\"" or exerciseName == "\"Remada\"" 
        or exerciseName == "Rowing (Machine)" or exerciseName == "Bent Over Row (Barbell)" or exerciseName == "Remada"):

        exerciseName = "Rowing (Machine)"
        return exerciseName
        
    elif(exerciseName == "\"Pull Up\"" or exerciseName == "\"Chin Up\"" or exerciseName == "\"Pull Up (Band)\"" or exerciseName == "\"Lat Pulldown (Cable)\"" or exerciseName == "\"Lat Pulldown\""
        or exerciseName == "Pull Up" or exerciseName == "Chin Up" or exerciseName == "Pull Up (Band)" or exerciseName == "Lat Pulldown (Cable)" or exerciseName == "Lat Pulldown" ):
        exerciseName = "Pull Up"
        return exerciseName

    elif(exerciseName == "\"Shoulder Press (Machine)\"" or exerciseName =="\"Seated Overhead Press (Dumbbell)\"" or exerciseName == "Shoulder Press (Machine)" or exerciseName =="Seated Overhead Press (Dumbbell)"
        or exerciseName == "\"Hammer Strength Shoulder Press\"" or exerciseName == "Hammer Strength Shoulder Press" ):
        exerciseName = "Shoulder Press (Machine)" 
        return exerciseName

    elif(exerciseName == "\"Stiff Leg Deadlift (Barbell)\"" or exerciseName == "\"Stiff Leg Deadlift (Dumbbell)\"" or exerciseName =="\"Peso morto stiff\"" or exerciseName == "\"Deadlift (Band)\""
        or exerciseName == "\"Deadlift (Barbell)\"" or exerciseName == "Stiff Leg Deadlift (Barbell)" or exerciseName == "Stiff Leg Deadlift (Dumbbell)" or exerciseName =="Peso morto stiff" or exerciseName == "Deadlift (Band)"
        or exerciseName == "Deadlift (Barbell)"):
        exerciseName = "Stiff Leg Deadlift (Barbell)"
        return exerciseName

    elif(exerciseName == "\"Hip Thrust (Barbell)\"" or exerciseName == "Hip Thrust (Barbell)" or exerciseName == "\"Glute Kickback\"" or exerciseName == "Glute Kickback"):
        exerciseName = "Hip Thrust (Barbell)"
        return exerciseName
         
    else:
        exerciseName = "invalid"
        return exerciseName
    
############################################################
##########################################################

#users hashtable
users = {"josemiguel": r"FitNotes_Export josemiguel.csv","ruimoreira": r"strong_ruimoreira.csv", 
"inesbarros": "FitNotes_Export inesbarros.csv", "catarinaferreira": r"strong_catarinaferreira.csv",
"teste": r"strong_teste.csv", "rui": r"strong_rui.csv", "rui_teste": "strong_rui_teste.csv",
"joaodias": r"strong_joaodias.csv"}

#Ask for user input manually if not introduced in command line
# if(len(sys.argv) < 2):
#     userInput = str(input("Enter the user you want to update.\n")).lower()
# else:
#     userInput = sys.argv[1]

#Insert File Name Manually
#TODO Se for Android deves alterar os ; por , antes de usares esta script
userInput = "josemiguel"

#fetches the data according to user inputted
#CSV reader
data = []
with open(users[userInput], newline='') as file:
    reader = csv.reader(file, delimiter=',', quotechar='|')
    for row in reader:
        #Store all the rows in readArr to manipulate further
        data.append(row)

# PRE-DATA DIGEST TO REMOVE HOURS INCASE OF STRONG EXPORT CSV
if('FitNotes_Export' in users[userInput]):
    print("No Pre-Processing has to be done")
else:
    #Delete unecessary Headers
    del data[0][1] 
    del data[0][4]

    for i in range(1,len(data)):

        # Remove hours from Date
        prevDate = data[i][0]
        newOnlyDate = re.split('\s.*', prevDate)[0]
        data[i][0] = newOnlyDate

        #Remove unecessary rows (Workout Name, Weight Unit)
        del data[i][1] #Workout name
        del data[i][4] #Weight Unit

    #Write the modified Info in the CSV file
    with open(users[userInput], 'w', newline='') as csvfile: 
        # creating a csv writer object 
        csvwriter = csv.writer(csvfile, quotechar = '?')   

        # writing the data rows 
        csvwriter.writerows(data)
####################################################################

# Dates for validation
today = datetime.date.today()
todayObject = datetime.datetime(today.year, today.month, today.day)
lastWeek = today - datetime.timedelta(days=7) 
lastMonthDay = today - datetime.timedelta(days=28) 
lastMonthDayObject = datetime.datetime(lastMonthDay.year, lastMonthDay.month, lastMonthDay.day)

#Data,Nome do treino,Duração,Nome do exercício,Ordem da série,Peso,Reps,Distância,Segundos,Notas,Notas do treino,RPE
#Dates array
rowIndex = []
for rowi in range(1,len(data)):

    #Regex to isolate date from other stuff
    onlyDate = data[rowi][0]

    #Gets the day and year from the DataFrame object
    day = int(re.split('-',onlyDate)[2])
    month = int(re.split('-',onlyDate)[1])
    year = int(re.split('-',onlyDate)[0])
    actualDateTimeObject = datetime.datetime(year, month, day) #row date datetime object

    if(actualDateTimeObject >= lastMonthDayObject and actualDateTimeObject <= todayObject):
        rowIndex.append(rowi) #this value doesnt take into account first line(headers)


#Algorithm to store exercise performance
#Create the xlsx sheet in case doesnt exist

#Writes data to user CSV performance review file
filename = userInput +'_perfReview_' + str(today.day) + '_' + str(today.month) + '_' + str(today.year) + '.xlsx'
workbook = xlsxwriter.Workbook(filename)
worksheet = workbook.add_worksheet("Dados_Treino")

#Column and Line initialization
chosenArr = []
chosenLine = 1
bpLine = [3]
squatLine = [3]
rowLine = [3]
pullUpLine = [3]
spLine = [3]
stiffLine = [3]
hpLine = [3]
column = 1
line = 1

#Header format
headFormat = workbook.add_format()
headFormat.set_bold()
headFormat.set_bg_color("#366092")
headFormat.set_font_color("white")

#Sub-Header format 
subHeadFormat = workbook.add_format()
subHeadFormat.set_bg_color("#95B3D7")
subHeadFormat.set_font_color("white")

#Relative Header format (with width adjust)
relHeadWidthFormat = workbook.add_format()
relHeadWidthFormat.set_bold()
relHeadWidthFormat.set_bg_color("#963634")
relHeadWidthFormat.set_font_color("white")

#Relative Header format (without width adjust)
relHeadFormat = workbook.add_format()
relHeadFormat.set_bold()
relHeadFormat.set_bg_color("#963634")
relHeadFormat.set_font_color("white")

#Sub-Relative Perf header format 
subRelHeadFormat = workbook.add_format()
subRelHeadFormat.set_bg_color("#DA9694")
subRelHeadFormat.set_font_color("white")

#Headers
headerArr= np.arange(60)
headIndex = 0
for i in range(0,7):
    worksheet.write(line+1, headerArr[headIndex+1], "Set"    ,subHeadFormat )
    worksheet.write(line+1, headerArr[headIndex+2], "Weight" ,subHeadFormat )
    worksheet.write(line+1, headerArr[headIndex+3], "Reps"   ,subHeadFormat )
    worksheet.write(line+1, headerArr[headIndex+4], "RPE"    ,subHeadFormat )
    worksheet.write(line+1, headerArr[headIndex+5], "Data"   ,subHeadFormat )
    worksheet.write(line+1, headerArr[headIndex+6], "&Weight",subRelHeadFormat )
    worksheet.write(line+1, headerArr[headIndex+7], "&Reps"  ,subRelHeadFormat )
    headIndex +=8

#Exercise Headers 
worksheet.write(line, 1 ,"Bench Press", headFormat)
for iii in range(2,6):
    worksheet.write(line, iii ,"", headFormat) #fill the headers until Relative perf
worksheet.write(line, 6 ,"Relative Performance", relHeadWidthFormat)
worksheet.write(line, 7 ,"", relHeadFormat)

worksheet.write(line, 9 , "Squat", headFormat )
for iii in range(10,14):
    worksheet.write(line, iii ,"", headFormat)
worksheet.write(line, 14 ,"Relative Performance", relHeadWidthFormat)
worksheet.write(line, 15 ,"", relHeadFormat)

worksheet.write(line, 17, "Row" , headFormat)
for iii in range(18,22):
    worksheet.write(line, iii ,"", headFormat)
worksheet.write(line, 22 ,"Relative Performance", relHeadWidthFormat)
worksheet.write(line, 23 ,"", relHeadFormat)

worksheet.write(line, 25, "Pull Up", headFormat )
for iii in range(26,30):
    worksheet.write(line, iii ,"", headFormat)
worksheet.write(line, 30 ,"Relative Performance", relHeadWidthFormat)
worksheet.write(line, 31 ,"", relHeadFormat)

worksheet.write(line, 33, "Shoulder Press", headFormat )
for iii in range(34,38):
    worksheet.write(line, iii ,"", headFormat)
worksheet.write(line, 38 ,"Relative Performance", relHeadWidthFormat)
worksheet.write(line, 39 ,"", relHeadFormat)

worksheet.write(line, 41, "Stiff Leg Deadlift", headFormat )
for iii in range(42,46):
    worksheet.write(line, iii ,"", headFormat)
worksheet.write(line, 46 ,"Relative Performance", relHeadWidthFormat)
worksheet.write(line, 47 ,"", relHeadFormat)

worksheet.write(line, 49, "Hip Thrust", headFormat )
for iii in range(50,54):
    worksheet.write(line, iii ,"", headFormat)
worksheet.write(line, 54 ,"Relative Performance", relHeadWidthFormat)
worksheet.write(line, 55 ,"", relHeadFormat)

#Get Initial exercise name and Date
exerciseName = exerciseNameVerifier(data[rowIndex[0]][1])
date = data[rowIndex[0]][0]

#Previous Date and Exercise
#TODO ERROR WHEN THE FIRST EXERCISE INSIDE OF DATE RANGE IS INVALID ONE
previousDate = date
previousExercise = exerciseName

#Store the exercises performance individually in according array
for i in range(len(rowIndex)):

    row = rowIndex[i]

    #Protection to avoid having different exerciseName even though same exercise with different nomenclature
    exerciseName = exerciseNameVerifier(data[row][1])

    if(exerciseName == "invalid"):
        #Need to make for cycle restart with i iterated
        continue

    serie = int(data[row][2])

    #Protection against invalid Data input from users
    if(data[row][3] == ''):
        data[row][3] = 1
        
    weight = float(data[row][3])

    #Protection to Pull-Up Weight being zero
    if(exerciseName == "Pull Up" and weight == 0):
        weight = 1.0

    reps = int(data[row][4])
    rpe = data[row][5]
    date = data[row][0]

    #Exercise Columns
    bpCols = [1, 2, 3, 4, 5]
    squatCols = [9,10,11,12,13]
    rowCols = [17,18,19,20,21]
    pullUpCols = [25,26,27,28,29]
    spCols = [33,34,35,36,37]
    stiffCols = [41,42,43,44,45]
    hpCols = [49,50,51,52,53]

    #Writes two empty Line to divide different date workout data
    if(previousDate != date or previousExercise != exerciseName):
        emptyLine = chosenLine+1
        worksheet.write(emptyLine, chosenArr[0], "")
        worksheet.write(emptyLine+1, chosenArr[0], "")

        updateLine(previousExercise, emptyLine+1)

    #Exercise present in actual line
    #TODO, a regex that searches for the exercise name would be convenient to avoid a lot of ORs

    if(exerciseName == "Bench Press (Barbell)" or exerciseName == "Bench Press" or exerciseName == "Supino Deitada" or exerciseName == "Incline Bench Press (Barbell)" ):
        chosenArr = bpCols
        chosenLine = bpLine[-1] #Last Recorded Line for Exercise
    
    elif(exerciseName == "Squat (Barbell)" or exerciseName == "Squat (Machine)" or exerciseName == "Goblet Squat" or exerciseName == "Leg Press"):
        chosenArr = squatCols
        chosenLine = squatLine[-1]

    elif(exerciseName == "Rowing (Machine)" or exerciseName == "Bent Over Row (Barbell)"):
        chosenArr = rowCols
        chosenLine = rowLine[-1]
        
    elif(exerciseName == "Pull Up" or exerciseName == "Chin Up"):
        chosenArr = pullUpCols
        chosenLine = pullUpLine[-1]

    elif(exerciseName == "Shoulder Press (Machine)"):
        chosenArr = spCols
        chosenLine = spLine[-1]

    elif(exerciseName == "Stiff Leg Deadlift (Barbell)" or exerciseName == "Stiff Leg Deadlift (Dumbbell)"):
        chosenArr = stiffCols   
        chosenLine = stiffLine[-1] 

    elif(exerciseName == "Hip Thrust (Barbell)"):
        chosenArr = hpCols   
        chosenLine = hpLine[-1]


    #Add the series, weight, reps and RPE to the xlsx
    worksheet.write(chosenLine, chosenArr[0], serie)
    worksheet.write(chosenLine, chosenArr[1], weight)
    worksheet.write(chosenLine, chosenArr[2], reps)
    worksheet.write(chosenLine, chosenArr[3], rpe)
    worksheet.write(chosenLine, chosenArr[4], date)

    previousDate = date #Updates previousDate
    previousExercise = exerciseName
    updateLine(exerciseName, chosenLine) #Increments line after adding data


workbook.close()

######################### PART 2 - Volume calculation and cell coloring ###############################
#######################################################################################################
#######################################################################################################

exerciseList = ["Bench Press (Barbell)","Squat (Barbell)","Rowing (Machine)","Pull Up",
"Shoulder Press (Machine)","Stiff Leg Deadlift (Barbell)","Hip Thrust (Barbell)"]

#Reading library requirements
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter

#OpenPyXl file Initialization
wb = load_workbook(filename)
sheet = wb.active

#Function to check what was the last line written
def findLastLineExercise(exerciseName):
    #Exercise Columns
    bpCols = [1, 2, 3, 4, 5]
    squatCols = [9,10,11,12,13]
    rowCols = [17,18,19,20,21]
    pullUpCols = [25,26,27,28,29]
    spCols = [33,34,35,36,37]
    stiffCols = [41,42,43,44,45]
    hpCols = [49,50,51,52,53]

    if(exerciseName == "Bench Press (Barbell)" or exerciseName == "Bench Press" or exerciseName == "Supino Deitada" ):
        chosenArr = bpCols
        chosenLine = bpLine[-1] #Last Recorded Line for Exercise
    
    elif(exerciseName == "Squat (Barbell)" or exerciseName == "Squat (Machine)" or exerciseName == "Goblet Squat"):
        chosenArr = squatCols
        chosenLine = squatLine[-1]

    elif(exerciseName == "Rowing (Machine)"):
        chosenArr = rowCols
        chosenLine = rowLine[-1]
        
    elif(exerciseName == "Pull Up"):
        chosenArr = pullUpCols
        chosenLine = pullUpLine[-1]

    elif(exerciseName == "Shoulder Press (Machine)"):
        chosenArr = spCols
        chosenLine = spLine[-1]

    elif(exerciseName == "Stiff Leg Deadlift (Barbell)"):
        chosenArr = stiffCols   
        chosenLine = stiffLine[-1] 

    elif(exerciseName == "Hip Thrust (Barbell)"):
        chosenArr = hpCols   
        chosenLine = hpLine[-1]

    return chosenLine

#Function to get the cellValue
def cellValue(sheet, rows, cols):
    return sheet.cell(row = rows, column = cols).value

#Function that returns the cell in which the first data from the exercise is
def findExerciseIndex(exerciseName, sheet):
    findIndexArr = []

    for col_num in range(1,60):
        for row_num in range(1,4):

            cellExercise = str(sheet.cell(row = row_num, column = col_num).value)

            if cellExercise in exerciseName:

                serieCellIndex_col = col_num
                findIndexArr.append(serieCellIndex_col)

                serieCellIndex_row = row_num+2
                findIndexArr.append(serieCellIndex_row)

                return findIndexArr

# This function will be responsible for iterating all the series of the session and calculating volumes
def perfComparisonWritter(serieCellIndex, sheet, exerciseName):
    perfHashTable = {}

    col_s = serieCellIndex[0]
    row_s = serieCellIndex[1]

    exerciseName = exerciseNameVerifier(exerciseName)

    #Cell background color formatting variables
    cell_format = workbook.add_format()
    cell_format.set_pattern(1)  # This is optional when using a solid fill.
    cell_format.set_bg_color('orange')

    tempVolume = 0
    tempSerie  = 0
    tempWeight = 0
    tempReps   = 0

    seshVolume = 0
    seshSerie  = 0
    seshWeight = 0
    seshReps   = 0

    lastLine = findLastLineExercise(exerciseName)
    #Could iterate to insert RIGHT border until lastLine of exercise
    for ii in range(1, lastLine):
        sheet.cell(row = ii, column = col_s + 6).border = Border(right=Side(style='thick'))

    #Adjust Column Width for Relative Performance column
    RelPerfHeaders = [7,15,23,31,39,47,55]
    for ii in RelPerfHeaders:
        sheet.column_dimensions[get_column_letter(ii)].width = 19
    
    #Adjust Column Width for Data Column
    RelPerfHeaders = [6,14,22,30,38,46,54]
    for ii in RelPerfHeaders:
        sheet.column_dimensions[get_column_letter(ii)].width = 10

    i = 4 #Every first index of the exercises start in row 4
    firstWeekDate = ""

    #Iterate all lines of exerciseName columns
    while(i < lastLine + 1):

        #Session data
        if(sheet.cell(row = i, column = col_s).value != None):

            if(i == 4):
                firstWeekDate = cellValue(sheet, i, col_s + 4)

            #Add the weights x reps to volume
            tempDate = cellValue(sheet, i, col_s + 4)
            tempSerie = cellValue(sheet, i, col_s)
            tempWeight = cellValue(sheet, i, col_s + 1)
            tempReps   = cellValue(sheet, i, col_s + 2)
            tempVolume = float(tempWeight) * int(tempReps)

            #Add series to hashTable with date
            #Structure is for weights: w_SERIE_DATE
            tempWeightKey = "w_" + exerciseName + "_" + str(tempSerie) + "_" + str(tempDate)
            tempRepKey    = "r_" + exerciseName + "_" + str(tempSerie) + "_" + str(tempDate) 
            perfHashTable[tempWeightKey] = float(tempWeight)
            perfHashTable[tempRepKey] = int(tempReps)

            seshWeight += float(tempWeight)
            seshReps += int(tempReps)
            seshSerie = int(tempSerie)
            seshVolume += float(tempVolume)

            #Algorithm to add rep difference from last week session
            #ignore first week
            if(tempDate == firstWeekDate):
                #First week wont have comparisons only storage of temp values   
                print("")

            else:
                #Need to compare to PREVIOUS week only
                #Only compare if there is a equal series number of last week
                serieVerifier = "w_" + exerciseName + "_" + str(seshSerie) + "_" + str(previousWeek)

                if serieVerifier in perfHashTable.keys():

                    #Calculate difference between weeks
                    weightDiff = tempWeight - perfHashTable["w_" + exerciseName + "_" + str(seshSerie) + "_" + str(previousWeek)]
                    repDiff = tempReps - perfHashTable["r_" + exerciseName + "_" + str(seshSerie) + "_" + str(previousWeek)]

                    sheet.cell(row = i, column = col_s + 5).value = weightDiff
                    sheet.cell(row = i, column = col_s + 6).value = repDiff

                    if(weightDiff < 0):
                        #red color
                        sheet.cell(row = i, column = col_s + 5).fill = PatternFill(start_color= Color(indexed=29), fill_type = "solid")

                    elif(weightDiff == 0):
                        #yellow color
                        sheet.cell(row = i, column = col_s + 5).fill = PatternFill(start_color= Color(indexed=26), fill_type = "solid")

                    elif(weightDiff > 0):
                        #green color
                        sheet.cell(row = i, column = col_s + 5).fill = PatternFill(start_color= Color(indexed=50), fill_type = "solid")

                    if(repDiff < 0):
                        #red color
                        sheet.cell(row = i, column = col_s + 6).fill = PatternFill(start_color= Color(indexed=29), fill_type = "solid")

                    elif(repDiff == 0):
                        #yellow color
                        sheet.cell(row = i, column = col_s + 6).fill = PatternFill(start_color= Color(indexed=26), fill_type = "solid")

                    elif(repDiff > 0):
                        #green color
                        sheet.cell(row = i, column = col_s + 6).fill = PatternFill(start_color= Color(indexed=50), fill_type = "solid")

            #Increment i to pass to the next row
            i +=1

        else:
            #Calculate Medians
            weightMedian = seshWeight / seshSerie
            repsMedian = seshReps / seshSerie

            #Write volume medians
            sheet.cell(row = i, column = col_s - 1).value = "VOLUME" #Header of the row
            sheet.cell(row = i, column = col_s - 1).fill = PatternFill(start_color=Color(indexed=47), fill_type = "solid")

            sheet.cell(row = i, column = col_s).value = seshSerie
            sheet.cell(row = i, column = col_s).fill = PatternFill(start_color= Color(indexed=47), fill_type = "solid")

            sheet.cell(row = i, column = col_s+1).value = weightMedian
            sheet.cell(row = i, column = col_s+1).fill = PatternFill(start_color= Color(indexed=47), fill_type = "solid")

            sheet.cell(row = i, column = col_s+2).value = repsMedian
            sheet.cell(row = i, column = col_s+2).fill = PatternFill(start_color=Color(indexed=47), fill_type = "solid")

            sheet.cell(row = i, column = col_s+3).value = "Total Vol" 
            sheet.cell(row = i, column = col_s+3).fill = PatternFill(start_color=Color(indexed=47), fill_type = "solid") 

            sheet.cell(row = i, column = col_s+4).value = seshVolume
            sheet.cell(row = i, column = col_s+4).fill = PatternFill(start_color=Color(indexed=47), fill_type = "solid")

            #Save this week performance as previous week
            previousWeek = tempDate

            #Reset actual session data before moving to next session
            seshWeight = 0
            seshReps   = 0
            seshSerie  = 0
            seshVolume = 0
            i +=2 #Increment i to step over the next empty row

    

#INITIALIZE THE READING AND VOLUME CALCULATION AND WRITTING

#Function to choose appropriate array according to exericse Name
for exercise in exerciseList:
    cellIndex = findExerciseIndex(exercise, sheet)
    perfComparisonWritter(cellIndex, sheet, exercise)

wb.save(filename)
wb.close()
