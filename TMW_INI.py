# Author: Peter Kim
# Date Created: 2/8/2019
# This script will...
# -Write the TMW ini file
# -Write the AddIP.bat file
# BEFORE RUNNING THE SCRIPT
# 1. Create a folder called "Assets" in the same directory as the python file (Skip if Assets folder is created already)
# 2. Within the "Assets" folder, make sure the follow files are inside
#   a. AvailableIPs.txt
#   b. der_locations.csv
#   c. DONTTOUCH.ini
#   d. DOTOUCH.ini
#   e. READIP.bat

import re

assetsdir = 'C:/Users/kimp1/OneDrive - Edison International/10000DER to Camden/TMW/Assets/' #change this to match the path with the files above

def populatednp3():
    with open(assetsdir + 'der_locations.csv', 'r') as file:
        return [line.split(',')[1] for line in file if not line.split(',')[1].startswith('AGG')]

def populate20305():
    with open(assetsdir + 'der_locations.csv', 'r') as file:
        return [line.split(',')[1] for line in file if line.split(',')[1].startswith('AGG')]
    
def populateIPs(numIP, numDNP3, numAGG):
    retList = []
    with open(assetsdir + 'AvailableIPs.txt', 'r') as file:
        for line in file:
            if len(retList) == numIP:
                break
            if len(retList) == numDNP3 + 1:
                for i in range(numAGG):
                    retList.append(line.replace('\n',''))
                return retList
            else:
                retList.append(line.replace('\n',''))
    return retList

def populateChannelNames(DNP3List, AGGList):
    retList = []
    #DNP3
    if len(DNP3List) == 1:
        retList.append('DNP3_DERs')
    else:
        for i in range(len(DNP3List)):
            retList.append(DNP3List[i])
    #Monitoring Points
    retList.append('MonitoringPoints')
    #2030.5
    if len(AGGList) == 1:
        retList.append('IEEE_DERs')
    else:
        for i in range(len(AGGList)):
            retList.append(AGGList[i])
    return retList

########################################################################################################################
#Start of running code

DNP3_List = populatednp3()
AGG_List = populate20305()
DNP3ChannelCount = len(DNP3_List)
DNP3SessPerChan = 1
MPChannelCount = 1
AGGChannelCount = len(AGG_List)
AGGSessPerChan = 1

IPAddress = populateIPs(DNP3ChannelCount + MPChannelCount + AGGChannelCount, DNP3ChannelCount, AGGChannelCount)
ChannelAlias = populateChannelNames(DNP3_List, AGG_List)

writeINI = []
writeBatch = []

with open(assetsdir + 'DONTTOUCH.ini', mode='r') as file:
    for line in file:
        writeINI.append(line)

comIPRead = comIPPortRead = comNameRead = sessionRead = sessNameRead = sessLinkAddrRead = sessLocalAddrRead = False

with open(assetsdir + 'DOTOUCH.ini', 'r') as file:
    for line in file:
        # Do stuff here
        if 'PhysComLocalIpAddress' in line:
            if comIPRead:
                continue
            for i in range(0, len(IPAddress)):
                writeINI.append('PhysComLocalIpAddress[  ' + str(i) + ']=' + IPAddress[i] + '\n')
            comIPRead = True
        elif 'PhysComIpPort' in line:
            if comIPPortRead:
                continue
            #DNP3 DERs and MP
            writeINI.append('PhysComIpPort[0-' + str(DNP3ChannelCount + MPChannelCount - 1) +']=' + str(20000) + '\n')
            #IEEE 2030.5
            for i in range(0, AGGChannelCount):
                writeINI.append('PhysComIpPort[' + str(DNP3ChannelCount + MPChannelCount + i) +']=' + str(2000 + i) + '\n')
            comIPPortRead = True
        elif 'PhysComChnlName' in line:
            if comNameRead:
                continue
            for i in range(0, len(ChannelAlias)):
                writeINI.append('PhysComChnlName[  ' + str(i) + ']=' + ChannelAlias[i] + '\n')
            comNameRead = True
        elif 'SessionCommIndex' in line:
            if sessionRead:
                continue
            #DNP3
            for i in range(0, DNP3ChannelCount):
                for j in range(0, DNP3SessPerChan):
                    writeINI.append('SessionCommIndex[  ' + str(i*DNP3SessPerChan + j) + ']='
                                + str(i) + '\n')
            #MP
            for i in range(0, MPChannelCount):
                writeINI.append('SessionCommIndex[  ' + str(DNP3ChannelCount*DNP3SessPerChan + i) + ']=' + str(i + DNP3ChannelCount) + '\n')
            #IEEE 2030.5
            for i in range(0, AGGChannelCount):
                for j in range(0, AGGSessPerChan):
                    writeINI.append('SessionCommIndex[  ' + str((DNP3ChannelCount*DNP3SessPerChan + MPChannelCount) + i*AGGSessPerChan + j) + ']='
                                    + str(i + DNP3ChannelCount + MPChannelCount) + '\n')
            sessionRead = True
        elif 'SessionName' in line and not sessNameRead:
            if sessNameRead:
                continue
            #DNP3
            for i in range(0, DNP3ChannelCount):
                for j in range(0, DNP3SessPerChan):
                    writeINI.append('SessionName[' +str(i*DNP3SessPerChan + j) +
                                    ']=L' + str(j+1) + '\n')
            #MP
            for i in range(0, MPChannelCount): 
                writeINI.append('SessionName[' +str(DNP3ChannelCount*DNP3SessPerChan + i) + ']=L1\n')
            #IEEE 2030.5
            for i in range(0, AGGChannelCount):
                for j in range(0, AGGSessPerChan):
                    writeINI.append('SessionName[  ' + str((DNP3ChannelCount*DNP3SessPerChan + MPChannelCount) + i*AGGSessPerChan + j)
                                    + ']=L' + str(j+1) + '\n')
            sessNameRead = True
        elif 'SessionLinkAddress' in line:
            if sessLinkAddrRead:
                continue
            writeINI.append('SessionLinkAddress[0-' + str(DNP3ChannelCount*DNP3SessPerChan + MPChannelCount + AGGChannelCount*AGGSessPerChan - 1) + ']=1' + '\n')
            sessLinkAddrRead = True
        elif 'SessionLocalAddress' in line:
            if sessLocalAddrRead:
                continue
            writeINI.append('SessionLocalAddress[0-' + str(DNP3ChannelCount*DNP3SessPerChan + MPChannelCount - 1) + ']=21' + '\n')
            writeINI.append('SessionLocalAddress['  + str(DNP3ChannelCount*DNP3SessPerChan + MPChannelCount) + '-'
                                                    + str(DNP3ChannelCount*DNP3SessPerChan + MPChannelCount + AGGChannelCount*AGGSessPerChan - 1) + ']=3' + '\n')
            sessLocalAddrRead = True
        elif 'SessionOutgoingTimeZoneIgnoreDST' in line or 'SessionOutgoingTimeZoneIndex' in line or 'DNPEnableSecureAuthentication' in line or 'DNPTcpLinkStatusPeriod' in line:
            writeINI.append(re.sub('\[.*\]', '[0-' + str(DNP3ChannelCount*DNP3SessPerChan + MPChannelCount + AGGChannelCount*AGGSessPerChan - 1) + ']', line))
        # DbasSectorAddress
        elif 'DbasSectorAddress' in line:
            continue
        # I14AuthEnable
        elif 'I14AuthEnable' in line:
            continue
        # DERs
        elif '0- 10' in line:
            writeINI.append(    line.replace('0- 10', '0- ' + str(len(IPAddress) - 1))  )
        elif '[  0]' in line:
            continue
        # DNP3
        elif '1-' in line:
            writeINI.append(line.replace(re.search('1.+\]', line).group(0), '0-  ' + str(DNP3ChannelCount + MPChannelCount - 1) + ']'))
        # IEEE 2030.5
        elif '6-' in line:
            writeINI.append(line.replace(re.search('6.+\]', line).group(0), str(DNP3ChannelCount+MPChannelCount) + '-  '+str(len(IPAddress)-1) + ']'))

# Create a Batch File called AddIP.bat that corresponds to the IPs used to make the ini file
with open(assetsdir + 'READIP.bat', 'r') as file:
    for line in file:
        if 'REM List of IP addresses to be added' in line:
            writeBatch.append(line)
            for i in range(0, len(IPAddress)):
                writeBatch.append('set invIP' + str(i) + '="' + IPAddress[i] + '"\n')
        elif 'set invIP' in line:
            continue
        elif 'if %2==add (' in line:
            writeBatch.append(line)
            for i in range(0, len(IPAddress)):
                writeBatch.append('        netsh interface ipv4 add address %1 %invIP' + str(i) + '% %subMask%\n')
        elif 'netsh interface ipv4' in line:
            continue
        elif 'if %2==del (' in line:
            writeBatch.append(line)
            for i in range(0, len(IPAddress)):
                writeBatch.append('        netsh interface ipv4 del address %1 %invIP' + str(i) + '% %subMask%\n')
        else:
            writeBatch.append(line)

# Write the ini and batch file
with open('C:/Users/kimp1/OneDrive - Edison International/10000DER to Camden/TMW/tmwgtway.ini', 'w') as writefile:
    writefile.writelines(writeINI)
with open('C:/Users/kimp1/OneDrive - Edison International/10000DER to Camden/TMW/AddIP.bat', 'w') as writefile:
    writefile.writelines(writeBatch)
