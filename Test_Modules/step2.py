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
    
    instructionText = [
        [sg.Text('Using Digital Meter to check the resistance between Test Point.')],
        [sg.Text('Lookup for board layout to see the location of Test Points.')],
        [sg.Text(size=(10,1))],
        
        # Fill in
        [sg.Text('Black Probe goes to JP5-pin2. Red Probe goes to JP5-pin1 (minimum 5000 Ohm):')],
        [sg.InputText(size=(20,1), key="JP5")],
        
        [sg.Text('Black Probe goes to JP6-pin2. Red Probe goes to JP6-pin1 (minimum 5000 Ohm):')],
        [sg.InputText(size=(20,1), key="JP6")],
        
        [sg.Text('Black Probe goes to JP7-pin2. Red Probe goes to JP7-pin1 (minimum 5000 Ohm):')],
        [sg.InputText(size=(20,1), key="JP7")],
        
        [sg.Text('Black Probe goes to JP1-pin2. Red Probe goes to JP1-pin1 (minimum 5000 Ohm):')],
        [sg.InputText(size=(20,1), key="JP1")],
        [sg.Text(size=(10,1))],
    ]
    
    layout = [
        [sg.Column(instructionText)],
        [sg.Button("Next"), sg.Exit(), sg.Button ("Layout")]
    ]
    
    window = sg.Window('Resistance Test', layout, size=(800,450), enable_close_attempted_event=True)

    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED or event == 'Exit':
            window.close()
            return 0, 0, 0, 0
        elif event == sg.WIN_CLOSE_ATTEMPTED_EVENT:
            window.close()
            return 0, 0, 0, 0
        elif event == "Layout":
            sg.popup_ok(image="Z:/05. Manufacturing/60. Uncontrolled/Troubleshoot/Phat/MIS/727-4239/Layout.png") #Change Address
        elif event == "Next":
            
            #grab number
            JP5 = int(values['JP5'])
            JP6 = int(values['JP6'])
            JP7 = int(values['JP7'])
            JP1 = int(values['JP1'])
            
            window.close()
            break
        
    return JP5, JP6, JP7, JP1

if __name__ == "__main__":
    print("Debug Mode")
    run()