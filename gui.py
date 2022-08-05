import PySimpleGUI as sg
import os.path
import date_popup as dp
from data_processor import create_report
from data_processor import convert_pd
from data_processor import check_timesheet_format
from datetime import datetime
import traceback
import textwrap

def wrapper(t):
	return textwrap.wrap(t, 120)

input_column = [
	[
		sg.Text('Enter Bealine paycheck:    '),
		sg.In(size=(25,1), enable_events=True, disabled=True, key="-FILE1-"),
		sg.FileBrowse(),
	],
	[
		sg.Text('Enter Fieldglass paycheck:'),
		sg.In(size=(25,1), enable_events=True, disabled=True, key="-FILE2-"),
		sg.FileBrowse(),
	],
	[
		sg.Text('Enter Capa paycheck:       '),
		sg.In(size=(25,1), enable_events=True, disabled=True, key="-FILE3-"),
		sg.FileBrowse(),
	],
	[
		sg.Text('Enter Paychex paycheck:  '),
		sg.In(size=(25,1), enable_events=True, disabled=True, key="-FILE4-"),
		sg.FileBrowse(),
	],
	
]


output_column = [
	[
		sg.In(size=(8,1), enable_events=True, disabled=True, key="-DATE1-"),
		sg.Button('Beginning Date'),
		sg.In(size=(8,1), enable_events=True, disabled=True, key="-DATE2-"),
		sg.Button('End Date'),
	],
	[
		sg.Text('(Option) Last Name, First Name:'),
		sg.In(size=(15,1), enable_events=True, disabled=False, key="-NAMEs-")],
	[
		sg.Text('Output Folder:'),
		sg.In(size=(21,1), enable_events=True, key="-FOLDER-"),
		sg.FolderBrowse(),
	],
	[
		sg.Button("Audit!"),
		sg.Button("Clear!")
	]
]

progress_text=sg.Text('Progress: ...', size=(80, None))

progress_row= [
	[progress_text],
	[sg.ProgressBar(2, orientation='h', size=(60, 20), key='progress')]
]

layout = [
	[
		sg.Column(input_column),
		sg.VSeperator(),
		sg.Column(output_column),
	],
	[
		sg.Text("_" * 110),
	],
	[
		progress_row
	]
]

window = sg.Window("Paychex Timesheet Validation", layout)

try:
	while True:
		event, values = window.read()
		if event == sg.WIN_CLOSED:
			break
		if event == "Beginning Date":
			date1 = dp.popup_get_date()
			if(date1!=None):
				window["-DATE1-"].update(date1)
		if event == "End Date":
			date2 = dp.popup_get_date()
			if(date2!=None):
				window["-DATE2-"].update(date2)
				# print(window["-DATE2-"].get())
		if event == "Audit!":
			con=True
			error_msg=""
			date1 = window["-DATE1-"].get()
			date2 = window["-DATE2-"].get()
			if (date1=='' or date2==''):
				con=False
				error_msg="ERROR: Date input missing. Please enter one in."

			else:
				disallowed_characters = "() "
				for character in disallowed_characters:
					date1=date1.replace(character, "")
					date2=date2.replace(character, "")

				date1 = tuple(map(int, date1.split(',')))
				# print(date2)
				date2 = tuple(map(int, date2.split(',')))

				date1 = datetime(int(date1[2]), int(date1[0]), int(date1[1]))
				date2 = datetime(int(date2[2]), int(date2[0]), int(date2[1]))

			check_format=check_timesheet_format(window["-FILE1-"].get(), "BL")+check_timesheet_format(window["-FILE2-"].get(), "FG")+check_timesheet_format(window["-FILE3-"].get(), "CW")+check_timesheet_format(window["-FILE4-"].get(), "PC")

			if (window["-FILE1-"].get()=='' and window["-FILE2-"].get()=='' and window["-FILE3-"].get()==''):
				con=False
				error_msg="ERROR: No input timetables found..."

			elif window["-FILE4-"].get()=='':
				con=False
				error_msg="ERROR: No paychex timetable found..."

			elif check_format!='':
				con=False
				error_msg="ERROR: {}timesheet(s) have incorrect format(s)...".format(check_format)

			elif date1>date2:
				con=False
				error_msg="ERROR: Start date is after end date..."

			elif (window["-FOLDER-"].get()==''):
				con=False
				error_msg="ERROR: No output folder."

			elif (window["-NAMEs-"].get()!=''):
				if ', ' not in window["-NAMEs-"].get():
					con=False
					error_msg="ERROR: Name not in correct format."

			if con:
				#display progress msg/update progress bar
				# BL=convert_pd(window["-FILE1-"].get(), type='BL')
				# FG=convert_pd(window["-FILE2-"].get(), type='FG')
				# CP=convert_pd(window["-FILE3-"].get(), type='CP')
				# PC=convert_pd(window["-FILE4-"].get(), type='PC')
				progress_text.Update(wrapper('Creating reports for entered timesheets...'))
				progress_text.update(text_color='black')
				window['progress'].UpdateBar(1)
				output_msg = create_report({"BL":window["-FILE1-"].get(), "FG":window["-FILE2-"].get(), "CP":window["-FILE3-"].get()}, window["-FILE4-"].get(), window["-FOLDER-"].get(), date1, date2, window["-NAMEs-"].get().upper())
				
				window['progress'].UpdateBar(2)
				progress_text.Update(wrapper(output_msg))
				progress_text.update(text_color='black')
			else:
				# display error message
				progress_text.Update(wrapper(error_msg))
				progress_text.Update(text_color='red')
		if event == "Clear!":
			window["-DATE1-"].update('')
			window["-DATE2-"].update('')
			window['progress'].UpdateBar(0)
			progress_text.Update('')
			window["-FILE1-"].update('')
			window["-FILE2-"].update('')
			window["-FILE3-"].update('')
			window["-FILE4-"].update('')
			window["-NAMEs-"].update('')
			window["-FOLDER-"].update('')


except Exception as e:
	# sg.Print('Exception in my event loop for the program:', sg.__file__, e, keep_on_top=True)
	sg.popup_error_with_traceback(f'An error happened.  Here is the info:', traceback.format_exc())

window.close()