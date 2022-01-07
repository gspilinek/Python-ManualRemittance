# written by Gerald Spilinek
# last edited 01/07/2022
# This program was made to make the manual remittance process faster for michelle to save her time at work
# it does this by utilizing a wrapper someone made for AHK and reading in the data from an excel spreadsheet
# it then grabs the ones we need clicks onto the window it needs and then starts typing

from ahk import AHK
import pandas as pd
import time
import keyboard

# File Path global Declarations
AHKFILEPATH = "PATH TO YOUR AHK FILE"
XCELFILEPATH = 'PATH TO THE EXCEL FILE'


def typeParamaters(ahk, xcel):  # given a spreadsheet of script#, amount paid, & date filled will process the remitance

    #index subject to change after comparing to the excel file,
    scripts = xcel['script #']
    paid = xcel['amount paid']
    datefilled = xcel['date filed']


    # click onto window to start typing
    # requires pre set up of making the website a left aligned window
    ahk.mouse_move(x=0, y=0, speed=1)
    ahk.click()
    time.sleep(2)
    ahk.click()

    # typing out each manual remittance
    for x in range(len(scripts)):
        if keyboard.is_pressed('esc'):  #kill switch hold ESC
            break
        #otherwise print the things, will finish typing the one it is on before it checks again
        ahk.type(str(scripts[x]))
        ahk.key_press('Tab')
        ahk.type(str(paid[x]))
        ahk.key_press('Tab')
        ahk.type(datefilled[x])
        ahk.key_press('Enter')
        time.sleep(1)   #had to add a delay as ahk was trying to submit to the website too fast



def main():
    ahk = AHK(AHKFILEPATH)  # declaring auto hot key object
    xcel = pd.read_csv(XCELFILEPATH)  # reading in the file
    typeParamaters(ahk, xcel)


if __name__ == "__main__" :
    main()

