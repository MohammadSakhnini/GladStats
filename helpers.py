import time
from datetime import date
from time import strftime
from time import gmtime
import calendar
import numpy as np
import pandas as pd
import hashmaps
from styleframe import StyleFrame
from datetime import datetime
from collections import Counter
COMPS = hashmaps.comps
MAPS = hashmaps.maps


READ_FILE_PATH = r"C:\Users\Mohammad Sakhnini\Desktop\Random_Files\GladPush stats\Data.xlsx"
SAVE_FILE_PATH = r"C:\Users\Mohammad Sakhnini\Desktop\Random_Files\GladPush stats\Formated_Data.xlsx"
df = pd.read_excel(READ_FILE_PATH, engine='openpyxl')

def GetTimeByStamp():
	timestamp = df["Timestamp"].tolist()
	if isinstance(timestamp,(list,np.dtype)):
		times = []
		for x in timestamp:
			times.append(time.strftime('%y-%m-%d %H:%M',time.localtime(x)))
		return np.array(times)

	return time.strftime('%H:%M', time.localtime(timestamp))

def GetWeekDayFromStamp():
	weekday_count = dict()
	for i in df["Timestamp"]:
		weekday_count[i] = weekday_count.get(i, 0) + 1
	return weekday_count

def WinLossPerDay():
	days = GetWeekDayFromStamp()
	for k,v in days.items():
		days[k] = [0,v]

	for x in range(len(df)):
		if df.iloc[x]["Timestamp"] == "Monday" and df.iloc[x]["Victory"] == True:
			days["Monday"][0] += 1
		if df.iloc[x]["Timestamp"] == "Tuesday" and df.iloc[x]["Victory"] == True:
			days["Tuesday"][0] += 1
		if df.iloc[x]["Timestamp"] == "Wednesday" and df.iloc[x]["Victory"] == True:
			days["Wednesday"][0] += 1
		if df.iloc[x]["Timestamp"] == "Thursday" and df.iloc[x]["Victory"] == True:
			days["Thursday"][0] += 1
		if df.iloc[x]["Timestamp"] == "Friday" and df.iloc[x]["Victory"] == True:
			days["Friday"][0] += 1
		if df.iloc[x]["Timestamp"] == "Saturday" and df.iloc[x]["Victory"] == True:
			days["Saturday"][0] += 1
		if df.iloc[x]["Timestamp"] == "Sunday" and df.iloc[x]["Victory"] == True:
			days["Sunday"][0] += 1
	return days

def PlayersSheet():
	player1,player2,player3 = [],[],[]
	class1,class2,class3 = [],[],[]
	columns = {}
	sheet = {}
	for i,comps in enumerate(df["EnemyComposition"].tolist()):
		cells = []
		comp = comps.split(",")
		sheet[i] = comp
	classes= []
	players = []
	for k,v in sheet.items():
		for x in range(3):
			if len(v) < 3:
				continue;
			class_spec = ("-".join(v[x].split("-",2)[:2]))
			name_server = ("-".join(v[x].split("-",2)[2:]))
			classes.append(class_spec)
			players.append(name_server)

	columns["player1"] = players[::3]
	columns["player2"] = players[1::3]
	columns["player3"] = players[2::3]
	columns["class1"] = classes[::3]
	columns["class2"] = classes[1::3]
	columns["class3"] = classes[2::3]

	df.drop("EnemyComposition",1,inplace=True)
	return columns

def GetMapsHeader():
	maps = []
	for k,v in MAPS.items():
		maps.append(v)
	return maps

def PreproccesData():
	temp = list()
	stamps = GetTimeByStamp()
	for stamp in stamps:
		weekday = calendar.day_name[datetime.strptime(stamp,'%y-%m-%d %H:%M').weekday()]
		temp.append(weekday)

	df["Timestamp"] = temp
	del temp
	df.drop("PlayersNumber",1,inplace=True)	
	df.drop("TeamComposition",1,inplace=True)
	df.drop("KillingBlows",1,inplace=True)
	df.drop("Damage",1,inplace=True)
	df.drop("Healing",1,inplace=True)
	df.drop("Honor",1,inplace=True)
	df.drop("Specialization",1,inplace=True)
	df.drop("isRated",1,inplace=True)
	df.drop("RatingChange",1,inplace=True)
	df["Map"] = df["Map"].map(MAPS)
	
def ScorePerMap():
	scoreBoard = {}
	countBoard = df["Map"].value_counts(ascending=True).to_dict()
	for i in GetMapsHeader():
		for k,v in countBoard.items():
			if k == i:
				scoreBoard[i] = [0,v]

	for x in range(len(df)):
		if df.iloc[x]["Map"] == "Ruins of Lordaeron" and df.iloc[x]["Victory"]:
			scoreBoard["Ruins of Lordaeron"][0] += 1

		if df.iloc[x]["Map"] == "Dalaran Sewers"and df.iloc[x]["Victory"]:
			scoreBoard["Dalaran Sewers"][0] += 1

		if df.iloc[x]["Map"] == "Tol'Viron Arena" and df.iloc[x]["Victory"]:
			scoreBoard["Tol'Viron Arena"][0] += 1

		if df.iloc[x]["Map"] == "Tiger's Peak" and df.iloc[x]["Victory"]:
			scoreBoard["Tiger's Peak"][0] += 1

		if df.iloc[x]["Map"] == "Black Rook Hold Arena" and df.iloc[x]["Victory"]:
			scoreBoard["Black Rook Hold Arena"][0] += 1

		if df.iloc[x]["Map"] == "Nagrand Arena" and df.iloc[x]["Victory"]:
			scoreBoard["Nagrand Arena"][0] += 1

		if df.iloc[x]["Map"] == "Ashamane's Fall" and df.iloc[x]["Victory"]:
			scoreBoard["Ashamane's Fall"][0] += 1

		if df.iloc[x]["Map"] == "Blade's Edge Arena" and df.iloc[x]["Victory"]:
			scoreBoard["Blade's Edge Arena"][0] += 1

		if df.iloc[x]["Map"] == "Hook Point" and df.iloc[x]["Victory"]:
			scoreBoard["Hook Point"][0] += 1

		if df.iloc[x]["Map"] == "Mugambala" and df.iloc[x]["Victory"]:
			scoreBoard["Mugambala"][0] += 1

		if df.iloc[x]["Map"] == "The Robodrome" and df.iloc[x]["Victory"]:
			scoreBoard["The Robodrome"][0] += 1

		if df.iloc[x]["Map"] == "Empyrean Domain" and df.iloc[x]["Victory"]:
			scoreBoard["Empyrean Domain"][0] += 1

	return scoreBoard

def MapsSheet():
	Scores = ScorePerMap()
	columns = {}
	columns[" "] = ["Win","Loss"]
	for v in Scores.values():
		v[1] = v[1] - v[0] 
	for k,v in Scores.items():
		columns[k] = v
	return columns

def GetComps():
	players_df = pd.read_excel(SAVE_FILE_PATH,sheet_name="PlayersSheet", engine='openpyxl')
	for x in range(len(players_df)):
		#RMP
		if "ROGUE" in players_df.iloc[x]["class1"] or "ROGUE" in  players_df.iloc[x]["class2"] or "ROGUE" in players_df.iloc[x]["class3"]:
			if "MAGE" in players_df.iloc[x]["class1"] or "MAGE" in  players_df.iloc[x]["class2"] or "MAGE" in players_df.iloc[x]["class3"]:
				COMPS["RMP"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["RMP"][0] += 1
		#Turbo
		if "Enhancement" in players_df.iloc[x]["class1"] or "Enhancement" in  players_df.iloc[x]["class2"] or "Enhancement" in players_df.iloc[x]["class3"]:
			if "WARRIOR" in players_df.iloc[x]["class1"] or "WARRIOR" in  players_df.iloc[x]["class2"] or "WARRIOR" in players_df.iloc[x]["class3"]:
				COMPS["Turbo"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Turbo"][0] += 1
		
		#Warr/Ret
		if "Retribution" in players_df.iloc[x]["class1"] or "Retribution" in  players_df.iloc[x]["class2"] or "Retribution" in players_df.iloc[x]["class3"]:
			if "WARRIOR" in players_df.iloc[x]["class1"] or "WARRIOR" in  players_df.iloc[x]["class2"] or "WARRIOR" in players_df.iloc[x]["class3"]:
				COMPS["Ret/Warr"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Ret/Warr"][0] += 1
		#DK
		if "DEATHKNIGHT" in players_df.iloc[x]["class1"] or "DEATHKNIGHT" in  players_df.iloc[x]["class2"] or "DEATHKNIGHT" in players_df.iloc[x]["class3"]:
			COMPS["DK/?"][1] += 1
			if df.iloc[x]["Victory"]:
				COMPS["DK/?"][0] += 1
		
		#WMP
		if "MAGE" in players_df.iloc[x]["class1"] or "MAGE" in  players_df.iloc[x]["class2"] or "MAGE" in players_df.iloc[x]["class3"]:
			if "WARRIOR" in players_df.iloc[x]["class1"] or "WARRIOR" in  players_df.iloc[x]["class2"] or "WARRIOR" in players_df.iloc[x]["class3"]:
				COMPS["WMP"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["WMP"][0] += 1
		#PHP
		if "Retribution" in players_df.iloc[x]["class1"] or "Retribution" in  players_df.iloc[x]["class2"] or "Retribution" in players_df.iloc[x]["class3"]:
			if "HUNTER" in players_df.iloc[x]["class1"] or "HUNTER" in  players_df.iloc[x]["class2"] or "HUNTER" in players_df.iloc[x]["class3"]:
				COMPS["PHP"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["PHP"][0] += 1
		#Thunder
		if "Elemental" in players_df.iloc[x]["class1"] or "Elemental" in  players_df.iloc[x]["class2"] or "Elemental" in players_df.iloc[x]["class3"]:
			if "WARRIOR" in players_df.iloc[x]["class1"] or "WARRIOR" in  players_df.iloc[x]["class2"] or "WARRIOR" in players_df.iloc[x]["class3"]:
				COMPS["Thunder"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Thunder"][0] += 1
		#Jungle
		if "Feral" in players_df.iloc[x]["class1"] or "Feral" in  players_df.iloc[x]["class2"] or "Feral" in players_df.iloc[x]["class3"]:
			if "HUNTER" in players_df.iloc[x]["class1"] or "HUNTER" in  players_df.iloc[x]["class2"] or "HUNTER" in players_df.iloc[x]["class3"]:
				COMPS["Jungle"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Jungle"][0] += 1

		#Caster Cleave
			#Warlock
		if "WARLOCK" in players_df.iloc[x]["class1"] or "WARLOCK" in  players_df.iloc[x]["class2"] or "WARLOCK" in players_df.iloc[x]["class3"]:
			if "Balance" in players_df.iloc[x]["class1"] or "Balance" in  players_df.iloc[x]["class2"] or "Balance" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1

		if "WARLOCK" in players_df.iloc[x]["class1"] or "WARLOCK" in  players_df.iloc[x]["class2"] or "WARLOCK" in players_df.iloc[x]["class3"]:
			if "Shadow" in players_df.iloc[x]["class1"] or "Shadow" in  players_df.iloc[x]["class2"] or "Shadow" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1

		if "WARLOCK" in players_df.iloc[x]["class1"] or "WARLOCK" in  players_df.iloc[x]["class2"] or "WARLOCK" in players_df.iloc[x]["class3"]:
			if "Elemental" in players_df.iloc[x]["class1"] or "Elemental" in  players_df.iloc[x]["class2"] or "Elemental" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1

		if "WARLOCK" in players_df.iloc[x]["class1"] or "WARLOCK" in  players_df.iloc[x]["class2"] or "WARLOCK" in players_df.iloc[x]["class3"]:
			if "MAGE" in players_df.iloc[x]["class1"] or "MAGE" in  players_df.iloc[x]["class2"] or "MAGE" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1


			#Boomi
		if "Balance" in players_df.iloc[x]["class1"] or "Balance" in  players_df.iloc[x]["class2"] or "Balance" in players_df.iloc[x]["class3"]:
			if "Shadow" in players_df.iloc[x]["class1"] or "Shadow" in  players_df.iloc[x]["class2"] or "Shadow" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1

		if "Balance" in players_df.iloc[x]["class1"] or "Balance" in  players_df.iloc[x]["class2"] or "Balance" in players_df.iloc[x]["class3"]:
			if "Elemental" in players_df.iloc[x]["class1"] or "Elemental" in  players_df.iloc[x]["class2"] or "Elemental" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1

		if "Balance" in players_df.iloc[x]["class1"] or "Balance" in  players_df.iloc[x]["class2"] or "Balance" in players_df.iloc[x]["class3"]:
			if "MAGE" in players_df.iloc[x]["class1"] or "MAGE" in  players_df.iloc[x]["class2"] or "MAGE" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1


			#Shadow Priest
		if "Shadow" in players_df.iloc[x]["class1"] or "Shadow" in  players_df.iloc[x]["class2"] or "Shadow" in players_df.iloc[x]["class3"]:
			if "Elemental" in players_df.iloc[x]["class1"] or "Elemental" in  players_df.iloc[x]["class2"] or "Elemental" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1

		if "Shadow" in players_df.iloc[x]["class1"] or "Shadow" in  players_df.iloc[x]["class2"] or "Shadow" in players_df.iloc[x]["class3"]:
			if "MAGE" in players_df.iloc[x]["class1"] or "MAGE" in  players_df.iloc[x]["class2"] or "MAGE" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1

			#Elemental
		if "Elemental" in players_df.iloc[x]["class1"] or "Elemental" in  players_df.iloc[x]["class2"] or "Elemental" in players_df.iloc[x]["class3"]:
			if "MAGE" in players_df.iloc[x]["class1"] or "MAGE" in  players_df.iloc[x]["class2"] or "MAGE" in players_df.iloc[x]["class3"]:
				COMPS["Caster Cleave"][1] += 1
				if df.iloc[x]["Victory"]:
					COMPS["Caster Cleave"][0] += 1
		#Resto Druid
		if "DRUID-Restoration" in players_df.iloc[x]["class1"] or "DRUID-Restoration" in  players_df.iloc[x]["class2"] or "DRUID-Restoration" in players_df.iloc[x]["class3"]:
			COMPS["Resto Druid"][1] += 1
			if df.iloc[x]["Victory"]:
				COMPS["Resto Druid"][0] += 1
		#Holy Priest
		if "PRIEST-Holy" in players_df.iloc[x]["class1"] or "PRIEST-Holy" in  players_df.iloc[x]["class2"] or "PRIEST-Holy" in players_df.iloc[x]["class3"]:
			COMPS["Holy Priest"][1] += 1
			if df.iloc[x]["Victory"]:
				COMPS["Holy Priest"][0] += 1
		#Holy Pala
		if "PALADIN-Holy" in players_df.iloc[x]["class1"] or "PALADIN-Holy" in  players_df.iloc[x]["class2"] or "PALADIN-Holy" in players_df.iloc[x]["class3"]:
			COMPS["Holy Pala"][1] += 1
			if df.iloc[x]["Victory"]:
				COMPS["Holy Pala"][0] += 1
		#Disc Priest
		if "PRIEST-Discipline" in players_df.iloc[x]["class1"] or "PRIEST-Discipline" in  players_df.iloc[x]["class2"] or "PRIEST-Discipline" in players_df.iloc[x]["class3"]:
			COMPS["Disc Priest"][1] += 1
			if df.iloc[x]["Victory"]:
				COMPS["Disc Priest"][0] += 1
		#Resto Sham
		if "SHAMAN-Restoration" in players_df.iloc[x]["class1"] or "SHAMAN-Restoration" in  players_df.iloc[x]["class2"] or "SHAMAN-Restoration" in players_df.iloc[x]["class3"]:
			COMPS["Resto Sham"][1] += 1
			if df.iloc[x]["Victory"]:
				COMPS["Resto Sham"][0] += 1
	return COMPS

def CompsSheet():
	comps = GetComps()
	columns = {}
	columns[" "] = ["Win","Loss"]
	for comp in comps.values():
		comp[1] = comp[1] - comp[0] 
	for k,v in comps.items():
		columns[k] = v
	return columns

def PlayerOccurance():
	players_df = pd.read_excel(SAVE_FILE_PATH,sheet_name="PlayersSheet", engine='openpyxl')
	player1 = Counter(players_df["player1"].value_counts().to_dict())
	player2 = Counter(players_df["player2"].value_counts().to_dict())
	player3 = Counter(players_df["player3"].value_counts().to_dict())
	players = dict(player1+player2+player3)
	players ={k: v for k, v in sorted(players.items(), key=lambda item: item[1],reverse=True)}
	top10 = list()
	del players_df
	for i,player in enumerate(players.keys()):
		top10.append(player)
		if i == 10:
			break
	return top10

def Top10Sheet():
	players_df = pd.read_excel(SAVE_FILE_PATH,sheet_name="PlayersSheet", engine='openpyxl')
	top10 = PlayerOccurance()
	results = dict()
	for i,player in enumerate(top10):
		top10[i] = player.split('-')[0]
	for player in top10:
		results[player] = [0,0]

	for player in results:
		for x in range(len(df) - 10):
			if player in players_df.iloc[x]["player1"] or player in players_df.iloc[x]["player2"] or player in players_df.iloc[x]["player3"]:
				results[player][1] += 1 
				if df.iloc[x]["Victory"]:
					results[player][0] += 1 
	columns = {}
	columns[" "] = ["Win","Loss"]
	for k,v in results.items():
		v[1] = v[1] - v[0]
		columns[k] = v
	return columns

def Init():
	try:
		PreproccesData()
		maps = MapsSheet()
		writer =  pd.ExcelWriter(SAVE_FILE_PATH, engine='openpyxl')
		pd.DataFrame(df).to_excel(writer,"DataSheet",header=True,index=False)
		pd.DataFrame(maps).to_excel(writer,'Maps',header=True,index=False)
		writer.save()
		playersheet = PlayersSheet()
		pd.DataFrame(playersheet).to_excel(writer,'PlayersSheet',header=True,index=False)
		writer.save()
		comps = CompsSheet()
		top10 = Top10Sheet()
		pd.DataFrame(comps).to_excel(writer,'Comps',header=True,index=False)
		pd.DataFrame(top10).to_excel(writer,'Top10',header=True,index=False)
		writer.save()
	except:
		print("Already formated")

##Init()
