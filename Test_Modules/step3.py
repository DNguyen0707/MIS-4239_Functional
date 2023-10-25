# import packages
import PySimpleGUI as sg
import datetime
import shutil
from pathlib import Path
import openpyxl

def run():
    #global var
    JP5 = 0
    JP6 = 0
    JP7 = 0
    JP1 = 0
    
    #set font
    sg.set_options(font=('Arial Bold', 14))
    
    picture = [
        [sg.Image(filename = 'Z:/05. Manufacturing/60. Uncontrolled/Troubleshoot/Phat/MIS/727-4239/Step3.png')],
        [sg.Text(size=(10,1))]
    ]
    
    instructionText = [
        [sg.Text('Record the supply current of the power supply (maximum 100mA):')],
        [sg.Text(), sg.InputText(size=(20,1), key="CurrSupply")],
        
        [sg.Text('Black Probe goes to JP5-pin2. Red Probe goes to JP5-pin1. Record in 2 decimal places (23.00V-25.00V):')],
        [sg.Text(), sg.InputText(size=(20,1), key="JP5")],
        
        [sg.Text('Black Probe goes to JP6-pin2. Red Probe goes to JP6-pin1. Record in 2 decimal places (4.80V-5.20V):')],
        [sg.Text(), sg.InputText(size=(20,1), key="JP6")],
        
        [sg.Text('Black Probe goes to JP7-pin2. Red Probe goes to JP7-pin1. Record in 2 decimal places (3.20V-3.40V):')],
        [sg.Text(), sg.InputText(size=(20,1), key="JP7")],
        
        [sg.Text('Black Probe goes to JP1-pin2. Red Probe goes to JP1-pin1. Record in 2 decimal places (14.00V-16.00V):')],
        [sg.Text(), sg.InputText(size=(20,1), key="JP1")],
    ]
    
    layout = [
        [sg.Column(picture)],
        [sg.Column(instructionText)],
        [sg.Button("Next"), sg.Exit(), sg.Button ("Layout")]
    ]
    
    window = sg.Window('Voltage Test', layout, size=(1000,900), enable_close_attempted_event=True)
    
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':
            return False
        elif event == sg.WIN_CLOSE_ATTEMPTED_EVENT:
            return False
        elif event == "Layout":
            sg.popup_ok(image="Z:/05. Manufacturing/60. Uncontrolled/Troubleshoot/Phat/MIS/727-4239/Layout.png") #Change Address
        elif event == "Next":
            
            #grab number
            Curr = int(values['CurrSupply'])
            JP5 = int(values['JP5'])
            JP6 = int(values['JP6'])
            JP7 = int(values['JP7'])
            JP1 = int(values['JP1'])
            
            window.close()
            break
        
    return Curr, JP5, JP6, JP7, JP1

if __name__ == "__main__":
    print("Debug Mode")
    run()