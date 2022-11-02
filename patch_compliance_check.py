# all data should be in sheet titled all within workbook
# data should be organized by site
# status indicators should be "on" and "off"

import openpyxl
import sys
import os
import calendar
import datetime

# remove all file writing and just print to command line maybe,
# include stops between so user can go through at own pace
# maybe add ability to customize hello message from "good morning"

# def main(path, saveDir):
def main(path):
    # path = "2022_10-28_patch_compliance.xlsx"
    wb = openpyxl.load_workbook(path[1])
    sh = wb["all"]

    tempOn = []
    tempOnSize = 0
    tempOff = []
    tempOffSize = 0
    siteHolder = sh.cell(2,8)

    message = "Good Morning!\n\nWe noticed that a couple of your systems appear to be "


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
            tempOnSize += 1
        elif (currStatus == "off"):
            tempOff.append([currName,currLastUser])
            tempOffSize += 1
        
        if (currSite != siteHolder):
            print(currSite)
            print(f"on: {tempOnSize}")
            print(f"off: {tempOffSize}")

            if ((tempOn > 0)  & (tempOff > 0)):
                message = message + "out of date or offline.\n\n"
            elif ((tempOn > 0)  & (tempOff == 0)):
                message = message + "out of date.\n\n"
            elif ((tempOff > 0)  & (tempOn == 0)):
                message = message + "offline.\n\n"

            if (tempOn > 0):
                message = message + "Please let us know a good time to schedule a reboot for these systems:\n"
                for i in tempOn:
                    # write to file here
                    print(f"{tempOn[0]}, last user {tempOn[1]}")
            

            tempOn = []
            tempOnSize = 0
            tempOff = []
            tempOffSize = 0
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
    
if __name__=="__main__":
    date = datetime.datetime.utcnow()
    dateFormatted = (str(date).split(" "))[0]

    currDir = os.path.dirname(os.path.realpath(__file__))
    newDir = "GetWell_" + dateFormatted
    dirFolderName = newDir

    os.mkdir(os.path.join(currDir,dirFolderName))
    print("Creating new dir \"" + dirFolderName + "\"")
    # main(sys.argv, dirFolderName)
    main(sys.argv)

    # not working properly, runs mutliple times when main() is run
    # itterator = 0

    # while True:
    #     try:
    #         # os.mkdir(os.path.join(currDir,dirFolderName))
    #         # print("Creating new dir \"" + dirFolderName + "\"")
    #         # main(sys.argv, dirFolderName)
    #         break
    #     except:
    #         if itterator < 10:
    #             itterator += 1
    #             dirFolderName = (newDir + "(" + str(itterator) + ")")
    #         else:
    #             print("Failsafe: too many folders with same name convention")
    #             break
            
        