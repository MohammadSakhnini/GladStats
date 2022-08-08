import helpers
import matplotlib.pyplot as plt
from matplotlib import pyplot
import numpy as np
import pandas as pd




def BarChart(labels,value1,value2,title):
	x = np.arange(len(labels))  # the label locations
	width = 0.30  # the width of the bars

	fig, ax = plt.subplots()
	rects1 = ax.bar(x - width/2, value1, width, label='Win',color='#2FA345')
	rects2 = ax.bar(x + width/2, value2, width, label='Loss',color='#D42023')

	ax.set_ylabel('Count')
	ax.set_title(title)
	ax.set_xticks(x)
	ax.set_xticklabels(labels,fontsize=8)
	ax.legend()

	ax.bar_label(rects1, padding=2)
	ax.bar_label(rects2, padding=5)
	fig.tight_layout()

	#plt.show()

def WinLossPerMap():
	maps_df = pd.read_excel(helpers.SAVE_FILE_PATH,sheet_name="Maps", engine='openpyxl')
	labels = maps_df.columns.tolist()[1:]
	win = maps_df.loc[0].tolist()[1:]
	loss = maps_df.loc[1].tolist()[1:]
	
	BarChart(labels, win, loss,'WinLossPerMap')

def WinLossPerWeekday():
	data = helpers.WinLossPerDay()
	labels = [*data]
	win = list()
	loss = list()
	for i in data.values():
		win.append(i[0])
		loss.append(i[1] - i[0])

	BarChart(labels, win, loss,'WinLossPerWeekDay')

def WinLossPerComp():
	comps_df = pd.read_excel(helpers.SAVE_FILE_PATH,sheet_name="Comps", engine='openpyxl')
	labels = comps_df.columns.tolist()[1:]
	win = comps_df.loc[0].tolist()[1:]
	loss = comps_df.loc[1].tolist()[1:]

	BarChart(labels, win, loss,'WinLossPerComp')

def WinLossTop10():
	top10_df = pd.read_excel(helpers.SAVE_FILE_PATH,sheet_name="Top10", engine='openpyxl')
	labels = top10_df.columns.tolist()[1:]
	win = top10_df.iloc[0].tolist()[1:]
	loss = top10_df.iloc[1].tolist()[1:]

	BarChart(labels, win, loss,'WinLossTop10')

def MMRGraph():
	mmr_df = pd.read_excel(helpers.SAVE_FILE_PATH,sheet_name="DataSheet", engine='openpyxl')
	reversed_mmr = mmr_df["MMR"][::-1].tolist()
	plt.plot(reversed_mmr)
	plt.tick_params(axis='x',bottom=False,top=False,labelbottom=False)
	plt.show()
