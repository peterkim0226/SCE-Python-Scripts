# Author: Peter Kim
# Date Created: 2/8/2019
# This script will write TMW configurations for DERs to a .csv file
# BEFORE RUNNING THE SCRIPT
# 1. Create a folder called "Assets" in the same directory as the python file (Skip if Assets folder is created already)
# 2. Within the "Assets" folder, make sure the follow files are inside
#   a. 2030_Tags.csv
#   b. Channel_Tags.csv
#   c. DNP3_Tags.csv
#   d. der_locations.csv
#   e. DONTTOUCH.ini
#   f. DOTOUCH.ini
#   g. mp_locations.csv
#   h. MP_Tags.csv
# No further set up is necessary

import csv
import sys

assetsdir = 'C:/Users/kimp1/OneDrive - Edison International/10000DER to Camden/TMW/Assets/' #change this to match the path with the files above

def writeToCSV(list):
    # write to csv file
    with open('C:/Users/kimp1/OneDrive - Edison International/10000DER to Camden/TMW/tmwgtway.csv', mode='w', newline = '') as file:
        writer = csv.writer(file, delimiter=',')
        for line in list:
            writer.writerow(line)
    file.close()

def initipchannel(channelname, channeltags, writelist):
	for line in channeltags:
		templine = line.copy()
		templine[0] = channelname + templine[0]
		templine[len(templine)-1] = templine[len(templine)-1].replace('\n', '')
		# if 'SessionActiveControl' in templine[0]:
		# 	templine[0] = templine[0].replace('L1', 'L' + str(j+1))
		writelist.append(templine)

def populateDNP3List():
    with open(assetsdir + 'der_locations.csv', 'r') as file:
        return [line.split(',')[1] for line in file if not line.split(',')[1].startswith('AGG')]

def populateMPList():
    with open(assetsdir + 'mp_locations.csv', 'r') as file:
        return [line.split(',')[0].replace('\n', '') for line in file]

def populateIEEEList():
    with open(assetsdir + 'der_locations.csv', 'r') as file:
        return [line.split(',')[1] for line in file if line.split(',')[1].startswith('AGG')]

def readDNP3Tags():
    with open(assetsdir + 'DNP3_Tags.csv', 'r') as file:
        return [line.split(',') for line in file]

def readMPTags():
    with open(assetsdir + 'MP_Tags.csv', 'r') as file:
        return [line.split(',') for line in file]

def read2030Tags():
    with open(assetsdir + '2030_Tags.csv', 'r') as file:
        return [line.split(',') for line in file]

def readChannelTags():
    with open(assetsdir + 'Channel_Tags.csv', 'r') as file:
        return [line.split(',') for line in file]

########################################################################################################################
#Start of running code

DNP3Tags = readDNP3Tags()
MPTags = readMPTags()
IEEETags = read2030Tags()
channelTags = readChannelTags()

DNP3Names = populateDNP3List()
MPNames = populateMPList()
IEEENames = populateIEEEList()

writeList = []

# Generation of DNP3
for i in range(0, len(DNP3Names)):
	T1 = 0
	T30 = 0
	T41 = 0
	# initialize channel
	DERName = DNP3Names[i]
	channelName = DERName
	initipchannel(channelName, channelTags, writeList)
	for line in DNP3Tags:
			if 'BESS' in DERName and 'Pcurtail' in line[0]:
				continue
			elif 'PV' in DERName and any(x in line[0] for x in ('Pref', 'SoC')):
				continue
			currentList = line.copy()
			currentList[0] = DERName + currentList[0]
			currentList[9] = channelName
			#     currentList[10] = 'L' + str(i+1)
			currentList[10] = 'L1'
			currentList[19] = DERName + currentList[19]
			if currentList[12] == str(1):
					currentList[13] = T1
					T1 = T1 + 1
			elif currentList[12] == str(30):
					currentList[13] = T30
					T30 = T30 + 1
			elif currentList[12] == str(41):
					currentList[13] = T41
					T41 = T41 + 1
			currentList[len(currentList) - 1] = currentList[len(currentList) - 1].replace('\n', '')
			writeList.append(currentList)

#Generation of Monitoring Points
T30 = 0
# initialize channel
channelName = 'MonitoringPoints'
initipchannel(channelName, channelTags, writeList)
for j in range(0, len(MPNames)):
	MPName = MPNames[j]
	for line in MPTags:
			currentList = line.copy()
			currentList[0] = MPName + currentList[0]
			currentList[9] = channelName
			currentList[10] = 'L1'
			currentList[13] = T30
			T30 = T30 + 1
			currentList[19] = MPName + currentList[19]
			currentList[len(currentList) - 1] = currentList[len(currentList) - 1].replace('\n', '')
			writeList.append(currentList)

#Generation of 2030.5 DERs
# initialize channel
for j in range(0, len(IEEENames)):
	DERName = IEEENames[j]
	channelName = DERName
	initipchannel(channelName, channelTags, writeList)
	for line in IEEETags:
		currentList = line.copy()
		currentList[0] = currentList[0].replace('REPLACEME', DERName)
		if not currentList[9] == '':
			currentList[9] = channelName
			# currentList[10] = 'L' + str(j+1)
			currentList[10] = 'L1'
		currentList[16] = currentList[16].replace('REPLACEME', DERName)
		currentList[16] = currentList[16].replace('"', '')
		currentList[len(currentList) - 1] = currentList[len(currentList) - 1].replace('\n', '')
		writeList.append(currentList)

writeToCSV(writeList)
