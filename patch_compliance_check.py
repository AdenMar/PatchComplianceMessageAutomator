# all data should be in sheet titled all within workbook
# data should be organized by site
# status indicators should be "on" and "off"

import openpyxl
import sys
import os
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
    tempOnLen = 0
    tempOff = []
    tempOffLen = 0
    siteHolder = sh.cell(2,8).value

    input(f"_____Press \"Enter\" to move onto {siteHolder}_____")

    # itterate through rows (start at 2 avoid header)
    # index bySite [n][0] is on [n][1] is off
    for row in range(2, sh.max_row+2): 
        currSite = sh.cell(row,8).value
        currLastUser = sh.cell(row,5).value
        currName = sh.cell(row,2).value
        currStatus = sh.cell(row,1).value

        # add status to on or off holders
        if (currStatus == "on"):
            tempOn.append([currName,currLastUser])
            tempOnLen += 1
        elif (currStatus == "off"):
            tempOff.append([currName,currLastUser])
            tempOffLen += 1
        
        if (currSite != siteHolder):
            print()
            message = "Good Morning!\n\nWe noticed that a couple of your systems appear to be "
            header = "\tPC Name\t\t\tLast User\n"

            # used for testing
            # print(siteHolder)
            # print(f"on: {tempOnLen}")
            # print(f"off: {tempOffLen}")

            if ((tempOnLen > 0)  & (tempOffLen > 0)):
                message = message + "out of date or offline.\n"
            elif ((tempOnLen > 0)  & (tempOffLen == 0)):
                message = message + "out of date.\n"
            elif ((tempOnLen > 0)  & (tempOffLen == 0)):
                message = message + "offline.\n"
            
            print(message)

            if (tempOnLen > 0):
                print(f"For your out of date systems, please let us know a good time to schedule a reboot for these systems:\n\n{header}")

                for i in tempOn:
                    print("\t" + i[0] + "\t\t\t" + i[1])
                print("\nIt is important to keep these systems up to date to minimize security risk and ensure your computers run as effectively as possible.\n")
            
            if (tempOffLen > 0):
                print(f"For you offline systems that are still in use, please boot these systems so we can ensure they are up to date:\n\n{header}")

                for i in tempOff:
                    print("\t" + i[0] + "\t\t\t" + i[1])
                print("\nIf no longer in use, please let us know instead so we may act accordingly.\n")
             
            print("Thank you,\nAden and Kristin\n")
            
            input(f"_____Press \"Enter\" to move onto {currSite}_____")

            # return to defaults
            tempOn = []
            tempOnLen = 0
            tempOff = []
            tempOffLen = 0
            siteHolder = currSite
            


            
            

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

    # os.mkdir(os.path.join(currDir,dirFolderName))
    # print("Creating new dir \"" + dirFolderName + "\"")
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
            
        