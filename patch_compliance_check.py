# all data should be in sheet titled all within workbook
# data should be organized by site
# status indicators should be "on" and "off" - not anymore

import openpyxl
import sys
import os
import datetime

# args to add:
# filter (on/off)
# customizable open and closer

def main(path, saveDir):
    wb = openpyxl.load_workbook(path[1])
    sh = wb["all"]
    outF = open(saveDir, "a")

    tempOn = []
    tempOnLen = 0
    tempOff = []
    tempOffLen = 0
    header = []

    for col in range(1,sh.max_column+1):
        # print(sh.cell(1,col).value)
        header.append(sh.cell(1,col).value)
    
    colSite = header.index("Site Name") + 1
    colLastUser = header.index("Last User") + 1
    colName = header.index("Name") + 1
    colStatus = header.index("Availability") + 1

    siteHolder = sh.cell(2,colSite).value

    # itterate through rows (start at 2 avoid header)
    # index bySite [n][0] is on [n][1] is off
    for row in range(2, sh.max_row+2): 
        currSite = sh.cell(row,colSite).value
        currLastUser = sh.cell(row,colLastUser).value
        currName = sh.cell(row,colName).value
        currStatus = sh.cell(row,colStatus).value
        
        if (currSite != siteHolder):
            writeAndLog("",outF)
            message = "Good Morning!\n\nWe noticed that a couple of your systems appear to be "
            header = "\tPC Name\t\tLast User"

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
            
            writeAndLog(message,outF)

            if (tempOnLen > 0):
                writeAndLog(f"For your out of date systems, please let us know a good time to schedule a reboot:\n\n{header}",outF)

                for i in tempOn:
                    writeAndLog("\t" + i[0] + "\t\t" + i[1] ,outF)
                writeAndLog("\nIt is important to keep these systems up to date to minimize security risk and ensure your computers run as effectively as possible.\n",outF)
            
            if (tempOffLen > 0):
                writeAndLog(f"For your offline systems that are still in use, please boot them so we can ensure they are up to date:\n\n{header}",outF)

                for i in tempOff:
                    writeAndLog("\t" + i[0] + "\t\t" + i[1],outF)
                writeAndLog("\nIf no longer in use, please let us know instead so we may act accordingly.\n",outF)
             
            writeAndLog("Thank you,\nAden and Kristen\n",outF)
            
            writeAndLog("___" + str(currSite) + "___",outF)

            # return to defaults
            tempOn = []
            tempOnLen = 0
            tempOff = []
            tempOffLen = 0
            siteHolder = currSite

        # add status to on or off holders
        if (currStatus == "true"):
            tempOn.append([currName,currLastUser])
            tempOnLen += 1
        elif (currStatus == "false"):
            tempOff.append([currName,currLastUser])
            tempOffLen += 1
            

def writeAndLog(content, file):
    file.write(content + "\n")
    # for debug perposes
    # print("File Updated")
    # print(content)

if __name__=="__main__":
    date = datetime.datetime.utcnow()
    # dateFormatted = (str(date).split(" "))[0]

    # currDir = os.path.dirname(os.path.realpath(__file__))
    # newDir = "GetWell_" + dateFormatted
    # dirFolderName = newDir
    newFile = "GetWell_" + str(date).replace(":","-") + ".txt"
    

    # os.mkdir(os.path.join(currDir,dirFolderName))
    # print("Creating new dir \"" + dirFolderName + "\"")
    main(sys.argv, newFile)
    # main(sys.argv)

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
            
        