import pandas as pd
import numpy as np
from datetime import datetime
import os
import logging
from Levenshtein import distance

# NOLTE, WILLIAM


#create one whole report after getting converted timesheets
def create_report(rep_arr, paychex_addr, report_addr, d1, d2, fname=''):
	now=datetime.now()
	timestr = now.strftime("%Y%m%d-%H%M")
	report_addr=os.path.sep.join([report_addr, 'paychex_report_{}.xlsx'.format(timestr)])

	# payPD=pd.read_excel(paychex_addr, index_col=[1,0])
	paychex_rep=convert_pd(paychex_addr, d1, d2, fname=fname)
	# paychex_rep.to_excel(r"C:\Users\raymo\OneDrive\Documents\Paychex proj\test data\paychex_rep_nolte.xlsx")

	date_in=np.array(paychex_rep.index)
	date_in=np.unique(date_in)

	out_items = []

	out_msgs=[]
	out_msgz=[]
	

	writer = pd.ExcelWriter(report_addr, engine = 'xlsxwriter')

	counter=0



	for key in rep_arr.keys():
		if rep_arr[key]==None or rep_arr[key]=='':
			continue

		output_cols=["Error Type", "Employee ID", "Date", "Name", "Pay Type", "Paychex Hours", "Outside Hours"]
		df_output=pd.DataFrame(columns=output_cols)

		df=convert_pd(rep_arr[key], d1, d2, key, fname)
		df, paychex_rep=name_converter(df, paychex_rep)

		if len(df)==0:
			out_msgs.append(key)
			continue

		try:
			aewa = str(df['AEWA'][0])
		except:
			df_output.to_excel(writer, sheet_name=key, index=False)
			continue
		PD_key_selection=paychex_rep.loc[(paychex_rep['AEWA']==aewa)]
		for date in pd.to_datetime(np.unique(df.index)):
			PD_date_selection=PD_key_selection.loc[(PD_key_selection.index == date)]
			iter_len=len(PD_date_selection)
			# if(iter_len==0):
			# 	logger.info("[MISSING TIME LOG] for {} for {}".format(date.strftime('%Y-%m-%d'), key))
			# 	continue
			
			for name in np.unique(df.loc[date]['Last Name, First Name']):
				name_selection=df.loc[(df['Last Name, First Name']==name) & (df.index==date)]
				t = np.unique(name_selection['Pay Type'])
				for Htype in t:		
					hour_selection=name_selection.loc[(name_selection['Pay Type']==Htype)].iloc[0]
					try:
						PD_hours=PD_date_selection.loc[(PD_date_selection['Last Name, First Name']==name) & (PD_date_selection['Pay Type']==Htype)]['Billable Hours']
						if(hour_selection['Billable Hours']!=PD_hours[0]):
							# logger.info("[HOURS MISMATCH] {} - {} - {} - Paychex {} hrs/{} {} hrs".format(aewa, date.strftime('%Y-%m-%d'), name, PD_hours[0], key, hour_selection['Billable Hours']))
							df_output=df_output.append({"Error Type": "HOURS MISMATCH", "Employee ID": hour_selection['Employee ID'], "Date": date.strftime('%Y-%m-%d'), "Name": name, "Pay Type": Htype, "Paychex Hours": PD_hours[0], "Outside Hours": hour_selection['Billable Hours']}, ignore_index=True)
					except:
						# logger.info("[NO PAYCHEX HOURS] {} - {} - {} - [{}] {} hrs".format(aewa, date.strftime('%Y-%m-%d'), name, Htype, hour_selection['Billable Hours']))
						null_hour_df=paychex_rep.loc[(paychex_rep.index==date) & (paychex_rep['Pay Type']==Htype) & (paychex_rep['Last Name, First Name']==name)]
						if len(null_hour_df)==0:
							df_output=df_output.append({"Error Type": "NO PAYCHEX HOURS", "Employee ID": hour_selection['Employee ID'], "Date": date.strftime('%Y-%m-%d'), "Name": name, "Pay Type": Htype, "Paychex Hours": '[None]', "Outside Hours": hour_selection['Billable Hours']}, ignore_index=True)
						elif hour_selection['Billable Hours']!=null_hour_df['Billable Hours'][0] and (null_hour_df['AEWA'][0]=='<Unassigned>' or null_hour_df['AEWA'][0]==aewa):
							df_output=df_output.append({"Error Type": "HOURS MISMATCHX", "Employee ID": hour_selection['Employee ID'], "Date": date.strftime('%Y-%m-%d'), "Name": name, "Pay Type": Htype, "Paychex Hours": null_hour_df['Billable Hours'][0], "Outside Hours": hour_selection['Billable Hours']}, ignore_index=True)
						

				PD_hours_selection=np.unique(PD_date_selection.loc[(PD_date_selection['Last Name, First Name']==name)]['Pay Type'])
				PD_hourz=np.array(list(filter(lambda x: x not in t, PD_hours_selection)))
				for Htype in PD_hourz:
					if (PD_date_selection.loc[(PD_date_selection['Last Name, First Name']==name) & (PD_date_selection['Pay Type']==Htype)]['Billable Hours'][0]>0):
						# logger.info("[PAY TYPE MISMATCH] {} - {} - {} - Paychex [{}] Pay Type/{} hours".format(aewa, date.strftime('%Y-%m-%d'), name, Htype, PD_date_selection.loc[(PD_date_selection['Last Name, First Name']==name) & (PD_date_selection['Pay Type']==Htype)]['Billable Hours'][0]))
						df_output=df_output.append({"Error Type": "PAY TYPE MISMATCH", "Employee ID": hour_selection['Employee ID'], "Date": date.strftime('%Y-%m-%d'), "Name": name, "Pay Type": Htype, "Paychex Hours": PD_date_selection.loc[(PD_date_selection['Last Name, First Name']==name) & (PD_date_selection['Pay Type']==Htype)]['Billable Hours'][0], "Outside Hours": "[None]"}, ignore_index=True)
		
		if len(df_output)==0:
			out_msgz.append(key)
			continue

		df_hours=df_output.loc[df_output["Error Type"]=="HOURS MISMATCH"]

		hours_cols=['Name', 'Total Hours Discrepancy']
		df_totals=pd.DataFrame(columns=hours_cols)

		for name in np.unique(df_hours['Name']):
			tot_hours=0
			name_hours=df_hours.loc[df_hours['Name']==name]
			for i in range(len(name_hours)):
				tot_hours+=name_hours.iloc[i]["Paychex Hours"]-name_hours.iloc[i]["Outside Hours"]
			df_totals=df_totals.append({'Name': name, 'Total Hours Discrepancy': tot_hours}, ignore_index=True)


		df_output.to_excel(writer, sheet_name=key, index=False)
		df_totals.to_excel(writer, sheet_name="{} Totals".format(key), index=False)

		counter+=1
		out_items.append(key)

	writer.save()
	writer.close()


################### make another error: paychex not found in outside
	
	if counter>0:
		output_msg = 'Done! Report generated for {} timesheets can be found at {}.'.format(out_items, report_addr)

		if len(out_msgs)>0:
			output_msg+=" Date/name outside of range in {}.".format(out_msgs)

		if len(out_msgz)>0:		
			output_msg+=" No discrepancies in {}.".format(out_msgz)

	elif len(out_msgs)>0:
		output_msg="Done! Date/name outside of range in {}.".format(out_msgs)

		if len(out_msgz)>0:		
			output_msg+=" No discrepancies in {}.".format(out_msgz)

	elif len(out_msgz)>0:		
			output_msg="Done! No discrepancies in {}.".format(out_msgz)

	return output_msg




#reformats each timesheet for easier processing
#potential problems: pay type matters
def convert_pd(timesheet_addr, d1, d2, type='PC', fname=''):
	#types: BL, FG, CP, PC
	#preprocess each timesheet and convert to one standardized format
	#add together hours for same date and same category

	need_keys=['Apply To Date', "Employee ID", 'AEWA', 'Last Name, First Name', 'Billable Hours', 'Pay Type']
	report = pd.DataFrame(columns=need_keys)
	valid_dates = np.array(pd.date_range(d1, d2))
	file_extension=os.path.splitext(timesheet_addr)[1]
	if file_extension in ['.xlsx', '.xls', '.xlsm']:
		timesheet_pd=pd.read_excel(timesheet_addr)
	else:
		timesheet_pd=pd.read_csv(timesheet_addr)
	# Paychex
	if type=='PC':
		# timesheet_pd=pd.read_excel(timesheet_addr)
		timesheet_pd['Last Name, First Name']=timesheet_pd['Last Name']+', ' + timesheet_pd['First Name']
		timesheet_pd['Last Name, First Name']=timesheet_pd['Last Name, First Name'].str.upper()
		timesheet_pd['Apply To Date']=pd.to_datetime(timesheet_pd['Apply To Date'], infer_datetime_format=True)
		timesheet_pd.set_index('Apply To Date', inplace = True)

		if fname!='':	
			timesheet_pd=timesheet_pd.loc[(timesheet_pd['Last Name, First Name']==fname)]
			if len(timesheet_pd)==0:
				for n in np.unique(report['Last Name, First Name']):
					if distance(n, fname)<=3:
						fname=n
						break
				timesheet_pd=timesheet_pd.loc[(timesheet_pd['Last Name, First Name']==fname)]

		date_in=np.array(timesheet_pd.index)
		date_in=[i for i in date_in if i in valid_dates]
		date_in=np.unique(date_in)

		#create array of indeces with duplicates eliminated
		#nested for loop, top loop iterates through array of indeces, bottom iterates through people for each date and fills in report
		for date in date_in:
			try:
				current_names=np.unique(timesheet_pd['Last Name, First Name'][date].values)
			except AttributeError:
				current_names=[timesheet_pd.loc[date]['Last Name, First Name']]
			for name in current_names:
				name_selection=timesheet_pd.loc[(timesheet_pd['Last Name, First Name']==name) & (timesheet_pd.index==date)]
				for aewa in np.unique(name_selection['AEWA']):
					aewa_selection=name_selection.loc[(name_selection['AEWA']==aewa)]
					t = np.unique(aewa_selection['Pay Type'])
					for Htype in t:
						hour_selection=aewa_selection.loc[(aewa_selection['Pay Type']==Htype)]
						hours=0
						for i in range(len(hour_selection)):
							hours+=hour_selection.iloc[i]['Total Paid Duration']
						report = report.append({"AEWA": aewa, "Apply To Date": date, "Last Name, First Name":name.upper(), "Pay Type":Htype, "Billable Hours": round(hours,1)}, ignore_index=True)
				# report.loc[date]=[name_selection.iloc[0]['AEWA'], name, hours, name_selection.iloc[0]['Pay Type']]
		report.set_index('Apply To Date', inplace = True)


	#Beeline/Boeing
	#date no double dig
	elif type=='BL':
		# timesheet_pd=pd.read_excel(timesheet_addr)

		for i in range(len(timesheet_pd['Last Name, First Name Middle Name'])):
			name = timesheet_pd.loc[i, 'Last Name, First Name Middle Name']
			namex = name.split(" ")
			namex = namex[0]+' '+namex[1]
			timesheet_pd.loc[i, 'Last Name, First Name Middle Name']=namex.upper()

		timesheet_pd['Date']=pd.to_datetime(timesheet_pd['Date'], infer_datetime_format=True)
		timesheet_pd.set_index('Date', inplace = True)

		if fname!='':	
			ts_temp=timesheet_pd.loc[(timesheet_pd['Last Name, First Name Middle Name']==fname)]
			new_name=''
			if len(ts_temp)==0:
				for n in np.unique(timesheet_pd['Last Name, First Name Middle Name']):
					if distance(n, fname)<=3:
						new_name=n
						break
			if new_name!='':
				timesheet_pd=timesheet_pd.loc[(timesheet_pd['Last Name, First Name Middle Name']==fname)]
			else:
				timesheet_pd=ts_temp

		date_in=np.array(timesheet_pd.index)
		date_in=[i for i in date_in if i in valid_dates]
		date_in=np.unique(date_in)
		# print(date_in)

		for date in date_in:
			# print(timesheet_pd['Last Name, First Name Middle Name'][date])
			try:
				current_names=np.unique(timesheet_pd['Last Name, First Name Middle Name'][date].values)
			except AttributeError:
				current_names=[timesheet_pd.loc[date]['Last Name, First Name Middle Name']]
			for name in current_names:
				name_selection=timesheet_pd.loc[(timesheet_pd['Last Name, First Name Middle Name']==name) & (timesheet_pd.index==date)]
				hours=0
				for i in range(len(name_selection)):
					hours+=name_selection.iloc[i]['Units']
				if hours>0:
					# datetime_object = datetime.strptime(date, '%m/%d/%Y')
					# date_str=datetime_object.strftime('%m/%d/%Y')
					# namex = name.split(" ")
					# namex = namex[0]+' '+namex[1]
					try:
						tID=name_selection.iloc[0]['Timesheet Header ID']
					except:
						tID=''
					report = report.append({"AEWA": 100004, "Employee ID": tID ,"Apply To Date": date, "Last Name, First Name":name, "Pay Type": 'Work', "Billable Hours": round(hours,1)}, ignore_index=True)
					# report.loc[date]=[100004, name, hours, 'Work']
		report.set_index('Apply To Date', inplace = True)

		# report=report.loc[report['Last Name, First Name']==]
		

	# Fieldglass/Honeywell
	#y-mm-dd
	elif type=='FG':
		# timesheet_pd=pd.read_excel(timesheet_addr)
		timesheet_pd['Last Name, First Name']=timesheet_pd['Last Name']+', ' + timesheet_pd['First Name']
		timesheet_pd['Last Name, First Name']=timesheet_pd['Last Name, First Name'].str.upper()
		timesheet_pd['Time Entry Date']=pd.to_datetime(timesheet_pd['Time Entry Date'], infer_datetime_format=True)
		timesheet_pd.set_index('Time Entry Date', inplace = True)

		if fname!='':	
			ts_temp=timesheet_pd.loc[(timesheet_pd['Last Name, First Name']==fname)]
			new_name=''
			if len(ts_temp)==0:
				for n in np.unique(timesheet_pd['Last Name, First Name']):
					if distance(n, fname)<=3:
						new_name=n
						break
			if new_name!='':
				timesheet_pd=timesheet_pd.loc[(timesheet_pd['Last Name, First Name']==fname)]
			else:
				timesheet_pd=ts_temp

		date_in=np.array(timesheet_pd.index)
		date_in=[i for i in date_in if i in valid_dates]
		date_in=np.unique(date_in)

		for date in date_in:
			try:
				current_names=np.unique(timesheet_pd['Last Name, First Name'][date].values)
			except AttributeError:
				current_names=[timesheet_pd.loc[date]['Last Name, First Name']]
			except TypeError:
				continue
			for name in current_names:
				name_selection=timesheet_pd.loc[(timesheet_pd['Last Name, First Name']==name) & (timesheet_pd.index==date)]
				hours=0
				for i in range(len(name_selection)):
					hours+=name_selection.iloc[i]['Total Billable Hours']
				if hours>0:
					# datetime_object = datetime.strptime(date, '%m/%d/%Y')
					# date_str=datetime_object.strftime('%m/%d/%Y')
					report = report.append({"AEWA": 100000, "Apply To Date": date, "Last Name, First Name":name, "Pay Type": 'Work', "Billable Hours": round(hours, 1)}, ignore_index=True)
					# report.loc[date]=[100000, name, hours, 'Work']

		report.set_index('Apply To Date', inplace = True)


	#Coupa/Blue Origin
	elif type=='CP':
		# timesheet_pd=pd.read_excel(timesheet_addr)
		timesheet_pd['Date']=pd.to_datetime(timesheet_pd['Date'], infer_datetime_format=True)
		timesheet_pd['CW Name']=timesheet_pd['CW Name'].str.upper()
		timesheet_pd.set_index('Date', inplace = True)

		if fname!='':	
			ts_temp=timesheet_pd.loc[(timesheet_pd['Last Name, First Name']==fname)]
			new_name=''
			if len(ts_temp)==0:
				for n in np.unique(timesheet_pd['Last Name, First Name']):
					if distance(n, fname)<=3:
						new_name=n
						break
			if new_name!='':
				timesheet_pd=timesheet_pd.loc[(timesheet_pd['Last Name, First Name']==fname)]
			else:
				timesheet_pd=ts_temp

		date_in=np.array(timesheet_pd.index)
		date_in=[i for i in date_in if i in valid_dates]
		date_in=np.unique(date_in)

		for date in date_in:
			try:
				current_names=np.unique(timesheet_pd['CW Name'][date].values)
			except AttributeError:
				current_names=[timesheet_pd.loc[date]['CW Name']]
			for name in current_names:
				name_selection=timesheet_pd.loc[(timesheet_pd['CW Name']==name) & (timesheet_pd.index==date)]

				hours = 0
				for i in range(len(name_selection)):
					hours+=name_selection.iloc[i]['Hours']
				if hours>0:
					report = report.append({"AEWA": 100005, "Employee ID": name_selection.iloc[0]['CW Number'], "Apply To Date": date, "Last Name, First Name":name, "Pay Type": 'Work', "Billable Hours": round(hours, 1)}, ignore_index=True)

				# t = np.unique(name_selection['Hours Type'])
				# for Htype in t:
				# 	hour_selection=name_selection.loc[(name_selection['Hours Type']==Htype)]
				# 	hours=0
				# 	for i in range(len(hour_selection)):
				# 		hours+=hour_selection.iloc[i]['Hours']
				# 	if hours>0:
				# 		# datetime_object = datetime.strptime(date, '%m/%d/%Y')
				# 		# date_str=datetime_object.strftime('%m/%d/%Y')
				# 		Htypex='Work'
				# 		if(Htype!='Regular Hours'):
				# 			Htypex='Overtime'
				# 		report = report.append({"AEWA": 100005, "Apply To Date": date, "Last Name, First Name":name, "Pay Type": Htypex, "Billable Hours": hours}, ignore_index=True)
						# report.loc[date]=[100005, name, ot_hours, 'Overtime']
						# report.loc[date]=[100005, name, hours, 'Work']
		report.set_index('Apply To Date', inplace = True)

	return report

def name_converter(report, PD_report):
	simNames={}
	for name in np.unique(report['Last Name, First Name']):
		for full in np.unique(PD_report['Last Name, First Name']):
			if distance(name, full)<=3:
				simNames[name]=full
	
	pdNames={}
	for f in simNames:
		full=simNames[f]
		for name in np.unique(PD_report['Last Name, First Name']):
			if distance(name, full)<=3:
				pdNames[name]=full

	for name in simNames:
		report.loc[report['Last Name, First Name']==name, 'Last Name, First Name']=simNames[name]

	for name in pdNames:
		PD_report.loc[PD_report['Last Name, First Name']==name, 'Last Name, First Name']=pdNames[name]


	return (report, PD_report)


def check_timesheet_format(addr, dtype):
	output_msg=""

	if addr!='':
		file_extension=os.path.splitext(addr)[1]
		if file_extension in ['.xlsx', '.xls', '.xlsm']:
			df=pd.read_excel(addr)
		else:
			df=pd.read_csv(addr)
		if dtype=="BL":
			if 'Last Name, First Name Middle Name' not in df.keys():
				output_msg="BL "
		elif dtype=="FG":
			if 'Time Entry Date' not in df.keys():
				output_msg="FG "
		elif dtype=="CW":
			if 'CW Name' not in df.keys():
				output_msg="CW "
		elif dtype=="PC":
			if 'AEWA' not in df.keys():
				output_msg="PC "

	return output_msg