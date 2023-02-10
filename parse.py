#!/Users/jeff/.pyenv/shims/python3

def main():
	"""Main funcation"""
	import re
	import os
	import shutil
	import xlsxwriter
	import pandas as pd
	import numpy as np
	import matplotlib.pyplot as plt
	from matplotlib_venn import venn2

	# Clean up from last run
	dir_path = './output'
	images_path = './output/images'
	spreadsheet_path = './output/spreadsheet'
	try:
		shutil.rmtree(dir_path)
	except OSError as e:
		print("Error: %s : %s" % (dir_path, e.strerror))

	# Create subfolders
	os.mkdir(dir_path)
	os.mkdir(images_path)
	os.mkdir(spreadsheet_path)

	# Pre-declare XLSX Workbook
	workbook = xlsxwriter.Workbook(f'CTR_Percent_Graphs.xlsx')

	# Read and store content of excel file
	read_file = pd.read_excel("./input/input.xlsx")

	# Write the dataframe object into csv file
	read_file.to_csv("./input/input.csv", index=None, header=True)

	# read csv file and convert into a dataframe object
	df = pd.DataFrame(pd.read_csv("./input/input.csv"))

	# show the dataframe
	#print(df)

	# remove first 3 rows, then reset indexes
	df.drop([0,1,2], axis=0, inplace=True)
	df = df.reset_index(drop=True)

	# replace nan with empty string using replace() function
	df2 = df.replace(np.NaN, '', regex=True)

	# show the dataframe
	#print(df2)

	# convert df to list of lists
	df_list_of_lists = df2.values.tolist()
	#print(df_list_of_lists)

	# convert certain values to ints or floats with slicing, as well as some cleanup
	temp_l = []
	for l in df_list_of_lists:
		if l[3] == '':
			print('ERROR: No week value. Exiting...\n')
			exit()
		else:
			l[3] = int(l[3])
		l[4] = l[4].rstrip()
		l[7] = l[7].rstrip()
		l[8] = l[8].rstrip()
		if l[9] != '':
			l[9] = int(l[9])
		if l[10] != '':
			l[10] = float(l[10])
		if l[11] != '':
			l[11] = round(float(l[11]), 2)
		if l[12] != '':
			l[12] = round(float(l[12]), 2)
		if l[13] != '':
			l[13] = round(float(l[13])*100, 2)
		if l[14] != '':
			l[14] = round(float(l[14])*100, 2)
		#print(l)
		temp_l.append(l)
	df2 = temp_l
	del(df_list_of_lists)
	df_list_of_lists = df2.copy()
	del(df2)
	#print(df_list_of_lists)

	# Columns for each list in df_list_of_lists:
	##############################################################################
	#   week_number, start_date, end_date, reporting_days, vendor, ad_name_group,
	#   target_group, impressions, amount_spent, cpm, thru_plays, vtr_in_percent,
	#   ctr_in_percent

	vendor_names_list = []
	for l in df_list_of_lists:
		if l[4] not in vendor_names_list:
			vendor_names_list.append(l[4])
	print(f'\nRunning Script...\n\n\n  Vendor names are: {vendor_names_list}\n')

	temp_dict = {}
	temp_list = []
	cntr = 0
	vendor = ''
	for l in df_list_of_lists:
		if cntr == 0:
			temp_list.append(l[0])
			temp_list.append(l[7])
			temp_list.append(l[8].lstrip())
			temp_list.append(l[14])
			vendor = l[4]
			temp_dict[vendor] = temp_list
			cntr += 1

		elif cntr > 0 and vendor != l[4]:
			temp_dict[vendor] = temp_list
			temp_list = []
			temp_list.append(l[0])
			temp_list.append(l[7])
			temp_list.append(l[8].lstrip())
			temp_list.append(l[14])
			vendor = l[4]
			cntr = 0

		else:
			temp_list.append(l[0])
			temp_list.append(l[7])
			temp_list.append(l[8].lstrip())
			temp_list.append(l[14])
			temp_dict[vendor] = temp_list
			cntr += 1
	#print('\n')
	#print(temp_dict)
	#print('\n')

	temp2_dict = {}
	temp2_list = []
	for key,value in temp_dict.items():
		i = 0
		while i < len(value):
			temp2_list.append(value[i+1] + ' - ' + value[i+2] + ':' + str(value[i+3]))
			i += 4
		temp2_dict[key] = temp2_list
		#print(temp2_dict)
		#print('\n')

		temp3_dict = {}
		for k, v in temp2_dict.items():
			for tup in v:
				ad_target = re.sub(':.*$', '', tup)
				ctr = re.sub('^.*:', '', tup)
				if ad_target not in temp3_dict.keys():
					temp3_dict[ad_target] = [float(ctr)]
				else:
					temp3_dict[ad_target].append(float(ctr))
			#print(temp3_dict)

			# Set figure default figure size
			plt.rcParams["figure.figsize"] = (20, 18)

			for k, v in temp3_dict.items():
				print(f'Vendor is: {key}')
				c = f'Vendor is:  {key}'
				g = f'Ad Name/Target Group is:  {k}'
				print(f'  Ad Name/Target Group is: {k}')
				d = f'Ad Name/Target CTR % Values are:  {v}'
				print(f'  Ad Name/Target CTR % Values are: {v}\n')
				x_temp = np.array(range(len(v)))
				x = x_temp + 1
				y = np.array(v)
				#print(x)
				#print(y)
				target_group =re.sub(' - .*', '', k)
				graph_name = 'Vendor=' + key + ' AdName-TargetGroup=' + target_group
				graph_name2 = 'Vendor = ' + key
				plt.xticks(range(0, len(x) + 1))
				plt.xlim(0, len(x)+1)
				plt.plot(x, y, linestyle="-", marker="o", label=graph_name)
				plt.grid()
				plt.xlabel('Ad Campaign Day', fontsize=15, fontweight='bold', labelpad=5)
				plt.ylabel('CTR %', fontsize=15, fontweight='bold', labelpad=5)
				plt.title(graph_name2,fontsize=15, fontweight='bold', pad='5.0')
				plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05),
				fancybox = True, shadow = True, ncol = 3)

				for xx,yy in zip(x,y):
					label = "{:.2f}%".format(yy)
					plt.annotate(label,
								 (xx,yy),
								 textcoords="offset points",
								 xytext=(5,10),
								 ha='center')
			plt.savefig(graph_name, dpi=300)
			plt.clf()
			temp2_dict = {}
			temp2_list = []
			temp3_dict = {}

		worksheet = workbook.add_worksheet(f'{key}')
		worksheet.insert_image('A1', f'{graph_name}.png', {'x_scale': 0.8, 'y_scale': 0.8})

	workbook.close()

	print('\nCreating graphs. Please wait...')
	print('Done with graphs.\n')
	print('\nNow creating spreadsheets. Please wait...')
	print('Done with spreadsheets.\n')

	shutil.move('CTR_Percent_Graphs.xlsx', './output/spreadsheet/CTR_Percent_Graphs.xlsx')
	png_list = [f for f in os.listdir('.') if f.endswith(('.png'))]
	for f in png_list:
		print(f)
		shutil.move(f, './output/images')

if __name__ == '__main__':
	main()


