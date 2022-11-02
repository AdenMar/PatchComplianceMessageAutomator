# all data should be in sheet titled all within workbook
# data should be organized by site
# status indicators should be "on" and "off"

import openpyxl
import sys
import os
import calendar
import datetime

def main(path, saveDir):
    # path = "2022_10-28_patch_compliance.xlsx"
    wb = openpyxl.load_workbook(path[1])
    sh = wb["all"]

    filtered = []
    bySite = []
    siteListHolder = []
    tempOn = []
    tempOff = []
    siteHolder = sh.cell(2,8)

    messageP1 = "Good Morning!\n\nWe noticed that a couple of your systems appear to be out of date or offline.\n\n"


    # itterate through rows (start at 2 avoid header)
    # index bySite [n][0] is on [n][1] is off
    for row in range(2, sh.max_row+1): 
        currSite = sh.cell(row,8).value
        currLastUser = sh.cell(row,5).value
        currName = sh.cell(row,2).value
        currStatus = sh.cell(row,1).value

        # add status to on or off holders
        if (currStatus == "on"):
            tempOn.append([currName,currLastUser])
        elif (currStatus == "off"):
            tempOff.append([currName,currLastUser])
        
        if (currSite != siteHolder):
            print(currSite)
            print(f"on: {findLen(tempOn)}")
            print(f"off: {findLen(tempOff)}")

            tempOn = []
            tempOff = []
            siteHolder = currSite
            # open file ~currSite~.txt
                # analyze on and off to determine format of doc
                    # append proper intro
                    # append ood data (if needed)
                    # append off data (if needed)
                    # append exit
            # close file
            # update siteHolder -- siteHolder = currSite
            # reset on and off holders

            # bySite.append(siteListHolder) no more need probably
            # writeToFile(currSite,messageP1,saveDir) no more need for this likely

            # f = open(currSite, "a")
            


            
            

def writeToFile(fileName, contents, saveDir):
    filePath = saveDir + "\\" + fileName + ".txt"
    f = open(filePath, "a")
    f.write(contents)
    f.close()

def findLen(tempList):
    count = 0
    for i in tempList:
        count += 1
    return count
    
if __name__=="__main__":
    date = datetime.datetime.utcnow()
    dateFormatted = (str(date).split(" "))[0]

    currDir = os.path.dirname(os.path.realpath(__file__))
    newDir = "GetWell_" + dateFormatted
    dirFolderName = newDir

    os.mkdir(os.path.join(currDir,dirFolderName))
    print("Creating new dir \"" + dirFolderName + "\"")
    main(sys.argv, dirFolderName)

        
    itterator = 0

    while True:
        try:

            break
        except:
            if itterator < 10:
                itterator += 1
                dirFolderName = (newDir + "(" + str(itterator) + ")")
            else:
                print("Failsafe: too many folders with same name convention")
                break
            
        