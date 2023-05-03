import PySimpleGUI as sg
from datetime import datetime
import pathlib
from openpyxl import Workbook
import openpyxl

#create excel file and give them headers and save it to location 
file = pathlib.Path("backend_Data.xlsx")
if file.exists ():

    pass

else:
    file=Workbook()
    sheet=file.active
    sheet["A1"]= "FIRST NAME"
    sheet["B1"]= "LAST NAME"
    sheet["C1"]= "NUMBER"
    sheet["D1"]= "BIRTH DATE"
    sheet["E1"]= "ADDRESS"
file.save("backend_Data.xlsx")


# windoews theam 
sg.theme('DarkAmber')

# All the stuff inside your window.
layout = [  [sg.Text('Details', enable_events=True,
                        key='-TEXT-', font=('Arial bold', 15),
                        expand_x=True, justification='center') ],
                        #First name:
                        [sg.T("")], 
                        [sg.Text("First Name:",size =(10, 1)), sg.Input(size=(30, 1),key="-fname-" )],
                        #last name:
                        [sg.T("")],
                        [sg.Text("Last Name:",size =(10, 1)), sg.Input(size=(30, 1),key="-lname-")],
                        #phone Number:
                        [sg.T("")],
                        [sg.Text("Number:",size =(10, 1)), sg.Input(size=(30, 1),key="-pnum-")],
                        #birth date:
                        [sg.T("")],
                        [sg.Text("Birth Date:",size =(10, 1)),sg.CalendarButton("Date", 
                        close_when_date_chosen=True,  target='-date-',
                        location=(0,0), no_titlebar=False ),(sg.Input(key='-date-', size=(22,1)))],
                        #Address
                        [sg.T("")],
                        [sg.Text("Address",size=(10, 1)),sg.Multiline(key='-address-', size=(30,3))],
                        #buttons
                        [sg.T("")],
                        [sg.Button('Save',button_color = 'Green'),
                        sg.Button('Clear',button_color = 'red'),
                        sg.Button('Cancel')]
         ]

# Create the Window
window = sg.Window('Personal Details', layout,size=(370,450))

# Event Loop to process "events" and button functions
while True:
    event, values = window.read()
    
    # if user closes window or clicks cancel
    if event == sg.WIN_CLOSED or event == 'Cancel':
        sg.popup('Do You Want to Close this Window') 
        break
    
    #Save button clicked
    if event == 'Save':
        file = openpyxl.load_workbook("backend_Data.xlsx")
        sheet=file.active
        sheet.cell(column=1, row=sheet.max_row+1,value=values['-fname-'])
        sheet.cell(column=2, row=sheet.max_row,value=values['-lname-'])
        sheet.cell(column=3, row=sheet.max_row,value=values["-pnum-"])
        sheet.cell(column=4, row=sheet.max_row,value=values["-date-"])
        sheet.cell(column=5, row=sheet.max_row,value=values["-address-"])
        file.save("backend_Data.xlsx")
        sg.popup('Details Save Successfully')
    
    #Clear button clicked    
    if event == 'Clear':
        window['-fname-'].update('')
        window['-lname-'].update('')
        window['-pnum-'].update('')
        window['-date-'].update('')
        window['-address-'].update('')
        sg.popup('Values Clear Successfully')

window.close()