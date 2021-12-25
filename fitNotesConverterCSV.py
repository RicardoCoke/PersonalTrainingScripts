#  Strong structure : 
# Date,Workout Name,Exercise Name,Set Order,Weight,Weight Unit,Reps,RPE,Distance,Distance Unit,Seconds,Notes,Workout Notes,Workout Duration (14 args)
# 2020-05-01 16:35:47,"Treino 1",7 min,"Deadlift (Band)",1,0,8,0,0,"","",8

# Finotes Structure
# Date,Exercise,Category,Weight (kgs),Reps,Distance,Distance Unit,Time (8 args)
# 2020-07-23,Globet Squat,Legs,0.0,12,,,


#Needed args Date, Exercise Name,Set Order,Weight,Reps,RPE,


#Reading library requirements
import sys
import csv

def strongConverter(readArr,i,ii):
    if(readArr[i][ii] == "Exercise"):
        readArr[i][ii] = "Exercise Name"

    if(readArr[i][ii] == "Weight (kgs)"):
        readArr[i][ii] = "Weight"

    if(readArr[i][ii] == "Category"):
        readArr[i][ii] = "Set Order"   


#OpenPyXl file Initialization
#Ask for user input manually if not introduced in command line
# if(len(sys.argv) < 2):
#     filename = str(input("Enter the Fitnotes filename you want to convert.\n")).lower() + ".csv"
# else:
#     filename = sys.argv[1] + ".csv"

#Row Array
readArr = []

#INPUT THE FILE NAME HERE !!!!!!!!!!!!!!!!!!!!!!!!!!
fileName = 'FitNotes_Export inesbarros.csv'

#CSV reader
with open(fileName, newline='') as file:
    reader = csv.reader(file, delimiter=',', quotechar='|')
    for row in reader:
        #Store all the rows in readArr to manipulate further
        readArr.append(row)

#Iterate over the readArr , change first row(HEADERS) to StrongApp Headers, add Set Order after Exercise

# Convert Fitnotes CSV into same Headers as StrongApp
# i represents the first row of the CSV (first line)
for i in range(len(readArr)):
    for ii in range(len(readArr[i])):
        strongConverter(readArr,i,ii)
        


# Replace converted CSV Category Header information with the correct Set Order Expected
setOrder = 1
row = 1
while(row < len(readArr)):

    #First row case
    if(row == 1):
        readArr[row][2] = setOrder
        readArr[row][6] = "" #RPE
        previousExercise = readArr[row][1]
        row +=1
        continue #Re-start the for cycle

    #Rest of the rows
    if(readArr[row][1] == previousExercise):
        setOrder +=1
        readArr[row][2] = setOrder
        readArr[row][6] = "" #RPE
        row +=1

    else:
        setOrder = 1
        readArr[row][2] = setOrder
        readArr[row][6] = "" #RPE
        previousExercise = readArr[row][1]
        row +=1

    
#Write the modified Info in the CSV file
with open(fileName, 'w', newline='') as csvfile: 
    # creating a csv writer object 
    csvwriter = csv.writer(csvfile)   

    # writing the data rows 
    csvwriter.writerows(readArr)

