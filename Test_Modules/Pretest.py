# import packages
import PySimpleGUI as sg
import datetime
import shutil
from pathlib import Path
import openpyxl

def run():
    #set font
    sg.set_options(font=('Arial Bold', 14))
    
    picture = [
        [sg.Image(filename = 'Z:/05. Manufacturing/60. Uncontrolled/Troubleshoot/Phat/MIS/727-4239/Req.png')],
        [sg.Text(size=(10,1))]
    ]
    
    instructionText = [
        [sg.Text('Close the window after done reading')],
    ]
    
    layout = [
        [sg.Column(picture)],
        [sg.Column(instructionText)],
        [sg.Button("Next"), sg.Exit()]
    ]
    
    window = sg.Window('Pre-Test', layout, size=(950,725), enable_close_attempted_event=True)

    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED or event == 'Exit':
            return False
        elif event == sg.WIN_CLOSE_ATTEMPTED_EVENT:
            return False
        elif event == "Next":            
            window.close()
            break
        
    return True

if __name__ == "__main__":
    print("Debug Mode")
    run()