"""
xl2csv - standalone background processing server
robby emslie

v. 0.1 alpha - Jan 2022
"""

import os
import shutil
import time
from datetime import datetime
import xl2csv

ROOT_DIR = "sample_data"
WATCH_DIR = os.path.join(ROOT_DIR,  "queue")
OUT_DIR = os.path.join(ROOT_DIR, "output")
PROCESSED_DIR = os.path.join(ROOT_DIR, "completed")
FAIL_DIR = os.path.join(ROOT_DIR, "failed")
LOG_FILE = os.path.join(ROOT_DIR, "log")
RUN_INTERVAL = 10   # interval in seconds
RUN_DTIME = datetime.today()

runlog = open(LOG_FILE, "w")
message = ""


def fileCheck():
    checkTime = str(datetime.today())
    queueList = []
    message = "Checking for new files at %s\n" % checkTime
    runlog.writelines(message)
    print(message)
    for root, directories, files in os.walk(WATCH_DIR):
        for name in files:
            if '.' in name:
                if name.split('.')[-1] == 'xlsx' or name.split('.')[-1] == 'xls':
                    fileDir = os.path.join(root, name)
                    queueList.append(fileDir)
                    message = "Found file %s - added to processing queue.\n" % (
                        name)
                    runlog.writelines(message)
                    print(message)
    message = "Queue listing completed.\nThe following files will be sent to parsing: %s\n" % queueList
    runlog.writelines(message)
    print(message)
    return queueList


def move_processed_file(processed_file, error=False):
    fileToMove = processed_file
    filename = fileToMove.split("/")[-1]
    if error is False:
        destinationFile = os.path.join(PROCESSED_DIR, filename)
    else:
        destinationFile = os.path.join(FAIL_DIR, filename)
    shutil.move(fileToMove, destinationFile)


def process_files(filesList):
    message = "Beginning file processing...\n"
    runlog.writelines(message)
    print(message)

    for file in filesList:
        processor = xl2csv.parse_to_csv(file, OUT_DIR)
        message = "File %s, with %s sheets, has been processed.\n" % (
            processor["wb_data"]["workbook"],
            processor["wb_data"]["number_of_sheets"])
        runlog.writelines(message)
        print(message)
        runlog.writelines("Sheet Name\tOutput File\n")
        for eachSheet in processor["sheets"]:
            of_name = eachSheet["sheet_file"].split("/")[-1]
            runlog.writelines('{}{}\n'.format(eachSheet["sheet_name"].ljust(20),
                                              of_name.ljust(
                                                  45)))

        move_processed_file(file, processor["wb_data"]["error_code"])
    return 0


def checkAndPush():
    message = "Checking for new files...\n"
    runlog.writelines(message)
    print(message)
    queue = fileCheck()

    message = "Found %s files; processing...\n" % str(len(queue))
    runlog.writelines(message)
    print(message)
    process_files(queue)


def startService():
    message = "Starting xl2csv service. Press [CTRL]-[C] to stop the process.\n"
    runlog.writelines(message)
    print(message)
    try:
        while True:
            checkAndPush()
            time.sleep(RUN_INTERVAL)
        pass
    except KeyboardInterrupt:
        message = "Service Stopped.\n"
        runlog.writelines(message)
        print(message)


def main():
    message = "xl2csv\nv. 0.1 alpha\n"
    runlog.writelines(message)
    print(message)

    powerSwitch = True

    while powerSwitch is True:
        startService()
        message = "Service Stopped...\n"
        runlog.writelines(message)
        print(message + "\n")
        a = input("Enter \'start\' and press [Enter] to restart; otherwise, \
            press [Enter] to quit? ")
        if a.upper() == "START":
            powerSwitch = True
        else:
            powerSwitch = False
    pass

    message = "Exiting xl2csv\n"
    runlog.writelines(message)
    print(message)


main()
