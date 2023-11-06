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
        sg.popup_ok(image="Z:/05. Manufacturing/60. Uncontrolled/Troubleshoot/Phat/MIS/727-4239/Layout.png") #Change Address
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
        os.environ["DC_NEXT_CAL"] = dcNextCal
        dotenv.set_key(dotenv_file, "DC_NEXT_CAL", os.environ["DC_NEXT_CAL"])
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
        os.environ["DMM_NEXT_CAL"] = dmmNextCal
        dotenv.set_key(dotenv_file, "DMM_NEXT_CAL", os.environ["DMM_NEXT_CAL"])
        window['dmmNextCal'].update(dmmNextCal)
        
        #test start
        
        #pretest
        import Pretest
        PreTestResult = Pretest.run()

        #step 2 - resistance test
        import Step2
        ResistResult = Step2.run()
        Resist_PRI = int(ResistResult[0])
        Resist_5V = int(ResistResult[1])
        Resist_3_3V = int(ResistResult[2])
        Resist_LED = int(ResistResult[3])
        
        if Resist_PRI < 5000: #More than 5000 to pass
            # fail
            sg.popup_ok(" Test Failed")
            continue
        window["2.1"].update(Resist_PRI)
        
        if Resist_5V < 5000: #More than 5000 to pass
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["2.2"].update(Resist_5V)
        
        if Resist_3_3V < 5000: #More than 5000 to pass
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["2.3"].update(Resist_3_3V)
        
        if Resist_LED < 5000: #More than 5000 to pass
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["2.4"].update(Resist_LED)
        
        
        #step 3 - Voltage Test
        import Step3
        VoltageResult = Step3.run()
        Voltage_Curr = float(VoltageResult[0])
        Voltage_PRI = float(VoltageResult[1])
        Voltage_5V = float(VoltageResult[2])
        Voltage_3_3V = float(VoltageResult[3])
        Voltage_LED = float(VoltageResult[4])
        
        if Voltage_Curr > 100: #Less than 100 to pass
            # fail
            sg.popup_ok("Test Failed")
            continue
        window["3.1"].update(Voltage_Curr)
        
        if 23 > Voltage_PRI or Voltage_PRI > 25: #Between 23V and 100V to pass
            # fail
            sg.popup_ok("Test Failed")
            continue
        window["3.2"].update(Voltage_PRI)
        
        if 4.8 > Voltage_5V and Voltage_5V > 5.2: #Between 4.8V and 5.2V to pass
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["3.3"].update(Voltage_5V)
        
        if 3.2 > Voltage_3_3V or Voltage_3_3V > 3.4: #Between 3.2V and 3.4V to pass
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["3.4"].update(Voltage_3_3V)
        
        if 14 > Voltage_LED or Voltage_LED > 16: #Between 14V and 16V to pass
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["3.5"].update(Voltage_LED)
        
        
        #Step 4 - PoE Test
        import Step4
        PoEResult = Step4.run()
        PoE_PRI = float(PoEResult[0])
        PoE_5V = float(PoEResult[1])
        PoE_3_3V = float(PoEResult[2])
        PoE_LED = float(PoEResult[3])
        
        #check if value is right
        if PoE_PRI < 22 or PoE_PRI > 32: #Between 22V and 32V
            # fail
            sg.popup_ok("Test Failed")
            continue
        window["4.1"].update(PoE_PRI)
        
        if PoE_5V < 4.8 or PoE_5V > 5.2: #between 4.8V and 5.2V
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["4.2"].update(PoE_5V)
        
        if PoE_3_3V < 3.2 or PoE_3_3V > 3.4: #Between 3.2V and 3.4V
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["4.3"].update(PoE_3_3V)
        
        if PoE_LED < 14 or PoE_LED > 16: #between 14V and 16V
            # fail
            sg.popup_ok("Functional Test Failed")
            continue
        window["4.4"].update(PoE_LED)
        
        import PostTest
        if (PostTest.run(wo, serial, values['OP'], date, Resist_PRI, Resist_5V, Resist_3_3V, Resist_LED, Voltage_Curr, Voltage_PRI, Voltage_5V, Voltage_3_3V, Voltage_LED, PoE_PRI, PoE_5V, PoE_3_3V, PoE_LED)):
            sg.popup_ok(image="Z:/05. Manufacturing/60. Uncontrolled/Troubleshoot/Phat/MIS/727-4239/Pass.png")
            window['Final'].update("Passed")