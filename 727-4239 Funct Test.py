#!/usr/bin/env python3

# import packages
import PySimpleGUI as sg
import datetime
import shutil
from pathlib import Path
import openpyxl


#import test module
import sys
import os
import dotenv
current_dir = os.path.dirname(os.path.abspath(__file__))
testModule_path = os.path.join(current_dir, "Test_Modules")
sys.path.append("Test_Modules")

dotenv_file = dotenv.find_dotenv()
dotenv.load_dotenv(dotenv_file)

#font setup
sg.set_options(font=('Arial Bold', 16))

#import variable
dcModel = os.getenv('DC_MODEL')
dcEQID = os.getenv('DC_EQID')
dcLastCal = os.getenv('DC_LAST_CAL')
dmmModel = os.getenv('DMM_MODEL')
dmmEQID = os.getenv('DMM_EQID')
dmmLastCal = os.getenv('DMM_LAST_CAL')

# result column
image_viewer_column = [
    [sg.Text("Functional Step's Detail:", )],
    [sg.Text(size=(40, 1), key="-TOUT-")],
    [sg.Image(key="-IMAGE-")],
    [sg.Text('Step 2/VCC_PRI: ', size=(22, 1)), sg.Text(key='2.1')],
    [sg.Text('Step 2/5V: ', size=(22, 1)), sg.Text(key='2.2')],
    [sg.Text('Step 2/3.3V: ', size=(22, 1)), sg.Text(key='2.3')],
    [sg.Text('Step 2/VCC_LED: ', size=(22, 1)), sg.Text(key='2.4')],
    [sg.Text('Step 3/Current: ', size=(22, 1)), sg.Text(key='3.1')],
    [sg.Text('Step 3/VCC_PRI: ', size=(22, 1)), sg.Text(key='3.2')],
    [sg.Text('Step 3/5V: ', size=(22, 1)), sg.Text(key='3.3')],
    [sg.Text('Step 3/3.3V: ', size=(22, 1)), sg.Text(key='3.4')],
    [sg.Text('Step 3/VCC_LED: ', size=(22, 1)), sg.Text(key='3.5')],
    [sg.Text('Step 4/VCC_PRI: ', size=(22, 1)), sg.Text(key='4.1')],
    [sg.Text('Step 4/5V: ', size=(22, 1)), sg.Text(key='4.2')],
    [sg.Text('Step 4/3.3V: ', size=(22, 1)), sg.Text(key='4.3')],
    [sg.Text('Step 4/VCC_LED: ', size=(22, 1)), sg.Text(key='4.4')],
    [sg.Text('Final Test Result: ', size=(22, 1)), sg.Text(key='Final')]
]

# information
functional_step_column = [
    [sg.Text('Operator Name: ', size=(22, 1)), sg.Combo(['Phat Huynh', 'Anh Phan', 'Thanh Lam', 'Dai Nguyen', 'Thao Tran', 'Trung Tran'], default_value='', key='OP')],
    [sg.Text('Work Order Number: ', size=(22, 1)), sg.InputText(size=(10, 1), key="WO")],
    [sg.Text('Serial Number: ', size=(22, 1)), sg.InputText(size=(10, 1), key="SN")],
    [sg.Text('Date: ', size=(22, 1)), sg.Text(key='DATE')],
    [sg.Text(size=(10,1))],
    
    [sg.Text('DC Power Supply')],
    [sg.Text('Model/Serial',size=(22, 1)), sg.InputText(size=(10,1), key="dcModel")],
    [sg.Text('EQID',size=(22, 1)), sg.InputText(size=(10,1), key="dcEQID")],
    [sg.Text('Last Cal',size=(22, 1)), sg.InputText(size=(10,1), key="dcLastCal")],
    [sg.Text('Next Cal',size=(22, 1)), sg.Text(key="dcNextCal")],
    [sg.Text(size=(10,1))],
    
    [sg.Text('Multimeter')],
    [sg.Text('Model/Serial',size=(22, 1)), sg.InputText(size=(10,1), key="dmmModel")],
    [sg.Text('EQID',size=(22, 1)), sg.InputText(size=(10,1), key="dmmEQID")],
    [sg.Text('Last Cal',size=(22, 1)), sg.InputText(size=(10,1), key="dmmLastCal")],
    [sg.Text('Next Cal',size=(22, 1)), sg.Text(key="dmmNextCal")],
    
]

# full layout
layout = [
    [sg.Text('727-4239 MIS CARRIER FUNCTIONAL TEST', font=('Arial Bold', 20), size=20, expand_x=True, justification='center')],
    [sg.Column(functional_step_column),
        sg.VSeperator(),
        sg.Column(image_viewer_column)],
    [sg.Text(size=(10,1))],
    [sg.Text('Press Start to begin the Test.')],
    [sg.Button("Start"), sg.Button("Board Layout"), sg.Exit()]
]


window = sg.Window("727-4239 Final Functional Test", layout, finalize=True, size=(900, 750), enable_close_attempted_event=True)

# update
window['dcModel'].update(dcModel)
window['dcEQID'].update(dcEQID)
window['dcLastCal'].update(dcLastCal)
window['dmmModel'].update(dmmModel)
window['dmmEQID'].update(dmmEQID)
window['dmmLastCal'].update(dmmLastCal)

while True:
    event, values = window.read()
    
    if(event == sg.WIN_CLOSE_ATTEMPTED_EVENT and sg.popup_yes_no("Exit Test?")):
        break
    if(event == sg.WIN_CLOSED or event == "Exit"):
        break
    if(event == "Board Layout"):
        sg.popup_ok(image="Z:/05. Manufacturing/60. Uncontrolled/Troubleshoot/Phat/MIS/727-4239/Layout.png")
    if(event == "Start"):
        
        # Check if empty
        if values['OP'] == '':
            sg.popup_no_buttons('Operator name is blank')
        elif values['WO'] == '':
            sg.popup_no_buttons('Work Order is blank')
        elif values['SN'] == '':
            sg.popup_no_buttons('Serial Number is blank')
        
        #get data
        wo = values['WO']
        serial = values['SN']
        date = datetime.date.today().strftime('%m/%d/%Y')
        window['DATE'].update(date)
        
        #check DC data
        if values['dcModel'] != dcModel or values['dcEQID'] != dcEQID or values['dcLastCal'] != dcLastCal:
            #if DC data is not the same as env, update env
            dcModel = values['dcModel']
            dcEQID = values['dcEQID']
            dcLastCal = values['dcLastCal']
            
            os.environ["DC_MODEL"] = dcModel
            dotenv.set_key(dotenv_file, "DC_MODEL", os.environ["DC_MODEL"])
            
            os.environ["DC_EQID"] = dcEQID
            dotenv.set_key(dotenv_file, "DC_EQID", os.environ["DC_EQID"])
            
            os.environ["DC_LAST_CAL"] = dcLastCal
            dotenv.set_key(dotenv_file, "DC_LAST_CAL", os.environ["DC_LAST_CAL"])
        
        #display date for dc
        dcNextCal = datetime.datetime.strptime(dcLastCal, '%m/%d/%y')
        dcNextCal = dcNextCal.replace(dcNextCal.year + 1).strftime('%m/%d/%Y')
        window['dcNextCal'].update(dcNextCal)
            
        #check DMM data
        if values['dmmModel'] != dmmModel or values['dmmEQID'] != dmmEQID or values['dmmLastCal'] != dmmLastCal:
            #if DMM data is not the same env, update env
            dmmModel = values['dmmModel']
            dmmEQID = values['dmmEQID']
            dmmLastCal = values['dmmLastCal']
            
            os.environ["DMM_MODEL"] = dmmModel
            dotenv.set_key(dotenv_file, "DMM_MODEL", os.environ["DMM_MODEL"])
            
            os.environ["DMM_EQID"] = dmmEQID
            dotenv.set_key(dotenv_file, "DMM_EQID", os.environ["DMM_EQID"])
            
            os.environ["DMM_LAST_CAL"] = dmmLastCal
            dotenv.set_key(dotenv_file, "DMM_LAST_CAL", os.environ["DMM_LAST_CAL"])
        
        #display date for dmm
        dmmNextCal = datetime.datetime.strptime(dmmLastCal, '%m/%d/%y')
        dmmNextCal = dmmNextCal.replace(dmmNextCal.year + 1).strftime('%m/%d/%Y')
        window['dmmNextCal'].update(dmmNextCal)
