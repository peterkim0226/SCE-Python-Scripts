#author: Peter Kim
#date created: 6/4/2019
#last updated: 6/14/2019
#This script is a 'cleaned up' version of the original addDER_Camden.py file with better legibility.
#6/14 Update: Can add signal variables to the der and its corresponding common (dsl) model
"""
Before starting, alter the derfilepath, pvMeasPath, and bessMeasPath to your needs.
Format of the derfilepath csv file will be |1.busbarlocation|2.dername|3.controlprotocol|4.derrating|5.energycapacity(-1 for PVs)|
For control protocol, the only choices can be 'DNP3' or '2030.5'
The unit for columns derrating and energycapacity is in kVA/kW
"""

derfilepath = 'der_locations.csv'
pvMeasPath = 'D:/Peter PF Files/1000 DER/pv_nrel_src.txt'       #path to pv profile
bessMeasPath = 'D:/Peter PF Files/1000 DER/'                    #path to bess DIRECTORY of profiles
nameplatePath = 'nameplate.txt'

tr_20305_singlephase = 'TTR_IEEE_12000_480_SingleEarthed'   # template for 2030.5 single phase der
tr_20305_twophase = 'TTR_IEEE_12000_480_SinglePhase'            # template for 2030.5 two phase der
tr_20305_threephase = 'TTR_IEEE_12000_480_ThreePhase'       # template for 2030.5 three phase der
tr_dnp3_twophase = 'TTR_DNP3_12000_480_SinglePhase'             # template for dnp3 two phase der
tr_dnp3_threephase = 'TTR_DNP3_12000_480'                   # template for dnp3 three phase der

# import all relevant modules
import powerfactory
import time
import re

def configReadDAT(readDAT, m, d, der, n, tag):
    readDAT.Multip = m              #Multiplicator
    readDAT.deadband = d            #OPC deadband
    readDAT.i_dat = 3               #DAT type = REAL
    readDAT.iStatus = 1073741824    #Status
    readDAT.pCalObj = der           #Object (get calculated data from)
    readDAT.pCalObjSim = der        #Sim Object (get calculated data from)
    readDAT.varCal = n              #Variable Name (Object)
    readDAT.varCalSim = n           #Variable Name (Sim Object)
    readDAT.sTagID = tag            #TagID

def configWriteDAT(writeDAT, m, d, ctrl, n, tag):
    writeDAT.Multip = m             #Multiplicator
    writeDAT.deadband = d           #OPC deadband
    writeDAT.i_dat = 3              #DAT type = REAL
    writeDAT.iStatus = 536870912    #Status
    writeDAT.pCtrl = ctrl           #Controller
    writeDAT.varName = n            #Variable name (for controller)
    writeDAT.sTagID = tag           #TagID

def configVmeasDAT(vmeasDAT, m, d, tag):
    vmeasDAT.Multip = m             #Multiplicator
    vmeasDAT.deadband = d           #OPC deadband
    vmeasDAT.iStatus = 1073741824   #Status
    vmeasDAT.sTagID = tag

def addpersonalbusbar(buslist, sysGrid, lib, mainbusbar, dername, protocol, nphase, phaselist):
    if protocol == 'DNP3':
        busname = re.sub('(BESS_)|(PV_)', '', dername) + '_DERs'
    else:
        busname = re.sub('(_BESS)|(_PV)', '', dername) + '_DERs'
    #if the location of the der (2030.5 and dnp3) has a transformer, 480V busbar, and an existing der, then just connect to that 480V busbar
    #otherwise, create a new transformer and new 480V busbar for that cluster
    existing480bar = app.GetCalcRelevantObjects(busname + '_480V' + '.ElmTerm')
    if canaddtoexisting480bar(existing480bar, busname, buslist, mainbusbar.loc_name):
        return addtoexistingpersonalbusbar(existing480bar[0]) 
    buscube = mainbusbar.CreateObject('StaCubic', busname + '_BarCub')
    configcubephase(buscube, nphase, phaselist)
    #create transformer
    tr = sysGrid.CreateObject('ElmTr2', busname + '_Tr')
    tr.typ_id = gettransformertype(lib, protocol, nphase)
    tr.bushv = buscube
    #create 480V busbar
    personalbar = sysGrid.CreateObject('ElmTerm', busname + '_480V')
    personalbar.uknom = 0.48
    if nphase == 2:   #2phase
        personalbar.phtech = 4   #2phase busbar
    buslist.append(personalbar.loc_name)
    trcube = personalbar.CreateObject('StaCubic', busname + '_TrCub')
    configcubephase(trcube, nphase, phaselist)
    tr.buslv = trcube
    return addtoexistingpersonalbusbar(personalbar)

def canaddtoexisting480bar(existing480bar, busname, buslist, mainbarname):
    if busname + '_480V' in buslist and existing480bar:
        existingmainbarname = existing480bar[0].GetConnectedCubicles()[0].obj_id.bushv.cterm.loc_name
        if existingmainbarname == mainbarname:
            return True
    return False

def gettransformertype(lib, protocol, nphase):
    #select transformer template
    if protocol == '2030.5':
        if nphase == 1:
            return lib.GetContents(tr_20305_singlephase + '.TypTr2')[0]  # 1 MVA Transformer Single Phase
        elif nphase == 2:
            return lib.GetContents(tr_20305_twophase + '.TypTr2')[0]  # 1 MVA Transformer Single Phase
        else:
            return lib.GetContents(tr_20305_threephase + '.TypTr2')[0]  # 1 MVA Transformer Three Phase
    else:
        if nphase == 2:
            return lib.GetContents(tr_dnp3_twophase + '.TypTr2')[0]  # 1 MVA Transformer Single Phase
        else:
            return lib.GetContents(tr_dnp3_threephase + '.TypTr2')[0]  # 1 MVA Transformer Three Phase

def addtoexistingpersonalbusbar(personalbar):
    dcube = personalbar.CreateObject('StaCubic', dername + '_Cub')
    return dcube

def configcubephase(dcube, nphase, phaselist):
    if nphase == 1:
        dcube.it2p1 = phaselist[0]
    elif nphase == 2:
        dcube.it2p1 = phaselist[0]
        dcube.it2p2 = phaselist[1]

def configder(der, dercube, size, phase, phaselist):
    if phase == 1:
        der.phtech = 2
    elif phase == 2:
        der.phtech = 4
    configcubephase(dercube, phase, phaselist)
    der.cCategory = 'stg'            # Plant Category
    der.iSimModel = 4           # Model = Constant Power
    # der.pgini = float(DERkVASize) / 1000  # Real power (kW to MW)             (Don't know what this is)
    der.sgn = float(size) / 1000  # size of der
    der.cosn = 1  # power factor
    der.bus1 = dercube  # Cubicle

def createmeasfile(compmodel, dername, pvpath, besspath):
    #Determines the measurement file's path based on type of DER
    if 'PV' in dername:
        measpath = pvpath
    else:
        measpath = besspath + dername + '.txt'
    # create measurement file
    meas = compmodel.CreateObject('ElmFile', 'Meas ' + dername)
    meas.iopt_imp = 1       #type of file = measurement file
    meas.f_name = measpath  #filename
    meas.approx = 0         #constant
    return meas

def createnameplatefile(compmodel, dername, nameplatepath):
    meas = compmodel.CreateObject('ElmFile', 'Meas ' + dername + ' Nameplate')
    meas.iopt_imp = 1       #type of file = measurement file
    meas.f_name = nameplatepath  #filename
    meas.approx = 0         #constant
    return meas

def configcommonmodel(commodel, pvcontrol, besscontrol, protocol, dername, size, energy):
    if protocol == '2030.5':
        protocolNum = 3
        initmva = 0
        initmwh = 0
    else:   #DNP3
        protocolNum = 2
        initmva = float(size) / 1000
        initmwh = float(energy) / 1000
    if 'PV' in dername:
        commodel.typ_id = pvcontrol
        commodel.params = [initmva, 1, 0.7, protocolNum]  # Size in kVA, Power factor, efficiency, 2=DNP3 / 3 =2030.5
    else:
        commodel.typ_id = besscontrol
        commodel.params = [initmva, initmwh, 0.5, 0.7, protocolNum]  # Size in kVA, Power factor, efficiency, 2=DNP3 / 3 =2030.5
    return commodel

def configvmeas(sysGrid, dername, busbar):
    v = sysGrid.CreateObject('StaVmea', 'Vmeas ' + dername)
    v.i_mode = 1
    v.pbusbar = busbar
    v.nphase = 3
    # if busbar.phtech == 0:     #3 phase scenario
    #     v.nphase = 3
    # else:
    #     v.nphase = 1
    #     v.it2p = busbar.GetConnectedCubicles()[0].it2p1
    return v

def configcompmodel(compmodel, pvframe, bessframe, vmeas, nameplatefile, measfile, der, commonmodel):
    # configure composite model
    if 'PV' in der.loc_name:
        compmodel.typ_id = pvframe
    else:
        compmodel.typ_id = bessframe
    modelparams = compmodel.pelm
    modelparams[0] = vmeas
    modelparams[1] = measfile
    modelparams[2] = nameplatefile
    modelparams[3] = der
    modelparams[4] = commonmodel
    compmodel.pelm = modelparams

def addDATs(der, dername, protocol, commonmodel, dercube):
    if protocol == '2030.5':   #IEEE 2030.5
        addDAT2030(der, dername, commonmodel, dercube)
    else:                   #DNP3
        addDATDNP3(der, dername, commonmodel, dercube)
    #vmeas
    vmeas_extname = dername + '_Vmeas'
    vmeas_mult = 0.00001
    vmeas_dead = 0.1
    vmeasdat = dercube.CreateObject('StaExtvmea', vmeas_extname)
    configVmeasDAT(vmeasdat, vmeas_mult, vmeas_dead, vmeas_extname)


def addDAT2030(der, dername, commonmodel, dercube):
    #pmeas
    pmeas_extname = dername + '_Pmeas'
    pmeas_mult = 0.1
    pmeas_dead = 0.001
    pmeasdat = dercube.CreateObject('StaExtdatmea', pmeas_extname)
    configReadDAT(pmeasdat, pmeas_mult, pmeas_dead, der, 'm:Psum:bus1', pmeas_extname)
    #pref_pct
    prefpct_extname = dername + '_ImmediateControls_WMaxLimPct'
    prefpct_mult = 1
    prefpct_dead = 0.1
    prefpctdat = dercube.CreateObject('StaExtdatmea', dername + '_Pref_pct')
    configWriteDAT(prefpctdat, prefpct_mult, prefpct_dead, commonmodel, 's:Pref_pct', prefpct_extname)
    #qmeas
    qmeas_extname = dername + '_Qmeas'
    qmeas_mult = 0.1
    qmeas_dead = 0.001
    qmeasdat = dercube.CreateObject('StaExtdatmea', qmeas_extname)
    configReadDAT(qmeasdat, qmeas_mult, qmeas_dead, der, 'm:Qsum:bus1', qmeas_extname)
    #qref_pct
    qrefpct_extname = dername + '_ImmediateControls_VArWMaxPct'
    qrefpct_mult = 1
    qrefpct_dead = 0.001
    qrefpctdat = dercube.CreateObject('StaExtdatmea', dername + '_Qref_pct')
    configWriteDAT(qrefpctdat, qrefpct_mult, qrefpct_dead, commonmodel, 's:Qref_pct', qrefpct_extname)
    if 'BESS' in dername:
        #soc
        soc_extname = dername + '_SoC'
        soc_mult = 0.01
        soc_dead = 0.01
        socdat = dercube.CreateObject('StaExtdatmea', soc_extname)
        configReadDAT(socdat, soc_mult, soc_dead, commonmodel, 's:SoC', soc_extname)
    #status
    status_extname = dername + '_Status'
    status_mult = 1
    status_dead = 1 #Not used for Status DAT
    statusdat = dercube.CreateObject('StaExtdatmea', status_extname)
    configReadDAT(statusdat, status_mult, status_dead, commonmodel, 'e:outserv', status_extname)
    statusdat.i_dat = 0

def addDATDNP3(der, dername, commonmodel, dercube):
    #pmeas
    pmeas_extname = dername + '_Pmeas'
    pmeas_mult = 1
    pmeas_dead = 0.01
    pmeasdat = dercube.CreateObject('StaExtdatmea', pmeas_extname)
    configReadDAT(pmeasdat, pmeas_mult, pmeas_dead, der, 'm:Psum:bus1', pmeas_extname)
    #qmeas
    qmeas_extname = dername + '_Qmeas'
    qmeas_mult = 1
    qmeas_dead = 0.01
    qmeasdat = dercube.CreateObject('StaExtdatmea', qmeas_extname)
    configReadDAT(qmeasdat, qmeas_mult, qmeas_dead, der, 'm:Qsum:bus1', qmeas_extname)
    #qref
    qref_extname = dername + '_Qref'
    qref_mult = 0.001
    qref_dead = 0.001
    qrefdat = dercube.CreateObject('StaExtdatmea', qref_extname)
    configWriteDAT(qrefdat, qref_mult, qref_dead, commonmodel, 's:Qref', qref_extname)
    #status
    status_extname = dername + '_Status'
    status_mult = 1
    status_dead = 1 #Not used for status DAT
    statusdat = dercube.CreateObject('StaExtdatmea', status_extname)
    configReadDAT(statusdat, status_mult, status_dead, commonmodel, 'e:outserv', status_extname)
    statusdat.i_dat = 0

    if 'PV' in dername:  # PV
        addDATPV(dername, commonmodel, dercube)
    else:  # BESS
        addDATBESS(dername, commonmodel, dercube)

def addDATPV(dername, commonmodel, dercube):
    #pcurtail
    pcurtail_extname = dername + '_Pcurtail'
    pcurtail_mult = 0.001
    pcurtail_dead = 0.001
    pcurtaildat = dercube.CreateObject('StaExtdatmea', pcurtail_extname)
    configWriteDAT(pcurtaildat, pcurtail_mult, pcurtail_dead, commonmodel, 's:Pcurtail', pcurtail_extname)

def addDATBESS(dername, commonmodel, dercube):
    #pref
    pref_extname = dername + '_Pref'
    pref_mult = 0.001
    pref_dead = 0.001
    prefdat = dercube.CreateObject('StaExtdatmea', pref_extname)
    configWriteDAT(prefdat, pref_mult, pref_dead, commonmodel, 's:Pref', pref_extname)
    #soc
    soc_extname = dername + '_SoC'
    soc_mult = 0.01
    soc_dead = 0.01
    socdat = dercube.CreateObject('StaExtdatmea', soc_extname)
    configReadDAT(socdat, soc_mult, soc_dead, commonmodel, 's:SoC', soc_extname)

def addsignalvariables(allCalc, der, commonmodel):
    # add signal variables for der
    resultvar_der = allCalc.CreateObject('IntMon', der.loc_name)
    resultvar_der.obj_id = der
    resultvar_der.vars = ['m:P:bus1', 'm:Q:bus1', 'm:U1l:bus1', 'm:cosphi:bus1', 'e:pf_recap', 'e:outserv', 'm:Psum:bus1', 'm:Qsum:bus1']
    # add signal variables for common model
    resultvar_cm = allCalc.CreateObject('IntMon', commonmodel.loc_name)
    resultvar_cm.obj_id = commonmodel
    if 'PV' in commonmodel.loc_name:
        resultvar_cm.vars = ['s:Pcurtail', 's:Qref', 's:Pout', 's:Qout', 's:E', 's:V_meas', 's:Pref_pct', 's:Qref_pct', 's:PFref', 'e:pf_recap']
    elif 'BESS' in commonmodel.loc_name:
        resultvar_cm.vars = ['s:Pref', 's:Qref', 's:Pout', 's:Qout', 's:SoC', 's:V_meas', 's:Pref_pct', 's:Qref_pct', 's:PFref', 'e:pf_recap']
    app.PrintInfo('Signal Variable has been added for ' + der.loc_name)


#####################################################################################

startTime = time.time()     #start recording time of code simulation

networkName = '*.ElmNet'

# create powerfactory app object
app = powerfactory.GetApplication()
app.SetUserBreakEnabled(0)
netData = app.GetProjectFolder('netdat')
library = app.GetProjectFolder('equip')     # equipment folder (used to retrieve Transformer Type)
sysGrid = netData.GetChildren(0, networkName, 0)[0]  # This returns the Camden Network
studyCase = app.GetActiveStudyCase()    # gets the active study case
allCalc = studyCase.GetChildren(0, '*.ElmRes', 0)[0]    # gets the 'All Calculations' Folder

# Retrieving user-defined 'SIMPLE' BESS and PV frame
blkPath = app.GetProjectFolder('blk')
pvBlk = blkPath.GetChildren(0, '*SIMPLE_PV.BlkDef', 0)[0]
bessBlk = blkPath.GetChildren(0, '*SIMPLE_STORAGE.BlkDef', 0)[0]
pvCtrl = blkPath.GetChildren(0, '*Simple PV.BlkDef', 0)[0]
bessCtrl = blkPath.GetChildren(0, '*Simple BESS.BlkDef', 0)[0]

derlist = [i.loc_name for i in app.GetCalcRelevantObjects('ElmGenstat')]
busbarlist = [i.loc_name for i in app.GetCalcRelevantObjects('ElmTerm')]

#Addition of DERs begins here
#Format of the csv file will be |1.busbarlocation|2.dername|3.controlprotocol|4.derrating|5.energycapacity(-1 for PVs)|
with open(derfilepath, 'r') as f:
    for line in f:
        line = line.replace('\n','')    #Removes 'next line' character from the line
        mainbusbarname = line.split(',')[0].strip()
        dername = line.split(',')[1].strip()
        derprotocol = line.split(',')[2].strip()
        dersize = line.split(',')[3].strip()
        derenergy = line.split(',')[4].strip()
        #check if der already exists; skip if true
        if dername in derlist:
            app.PrintInfo('ERROR. Cannot add DER --->' + dername)
            app.PrintInfo('DER already exists for ---->' + dername)
            app.PrintInfo('----------------------------------------------------------------------------------')
            continue
        #check if busbar location exists; skip if false
        if mainbusbarname not in busbarlist:
            app.PrintInfo('ERROR. Cannot add DER --->' + dername)
            app.PrintInfo('Busbar does not exist ---->' + mainbusbarname)
            app.PrintInfo('----------------------------------------------------------------------------------')
            continue
        derlist.append(dername)
        mainbusbar12kV = app.GetCalcRelevantObjects(mainbusbarname + '.ElmTerm')[0]
        mainbusbarcube = mainbusbar12kV.GetConnectedCubicles()[0]
        #Find out phase tech
        derphase = mainbusbarcube.nphase       #number of phases
        derphaselist = []
        if derphase == 1:       #1PH PH-E
            derphaselist.append(mainbusbarcube.it2p1)
        elif derphase == 2:     #1PH PH-PH
            derphaselist.append(mainbusbarcube.it2p1)
            derphaselist.append(mainbusbarcube.it2p2)
        app.PrintInfo('Configuring DER... -> ' + dername)
        dercube = addpersonalbusbar(busbarlist, sysGrid, library, mainbusbar12kV, dername, derprotocol, derphase,
                                    derphaselist)  # separate 480V busbar created to connect ONLY to DER; returns cubicle of der
        busbar480V = dercube.cterm
        der = sysGrid.CreateObject('ElmGenstat', dername)   #create der
        configder(der, dercube, dersize, derphase, derphaselist) # configure der
        compmodel = sysGrid.CreateObject('ElmComp', 'Control for ' + dername)  # create composite model
        measfile = createmeasfile(compmodel, dername, pvMeasPath, bessMeasPath)  # create measurement file
        nameplatefile = createnameplatefile(compmodel, dername, nameplatePath)  # create nameplate measurement file
        commonmodel = compmodel.CreateObject('ElmDsl', 'Interface for ' + dername)  #create common model
        configcommonmodel(commonmodel, pvCtrl, bessCtrl, derprotocol, dername, dersize, derenergy)  # configure common model
        vmeas = configvmeas(sysGrid, dername, busbar480V)  # create new Vmeas (not external DAT); used for configuring the composite model
        configcompmodel(compmodel, pvBlk, bessBlk, vmeas, nameplatefile, measfile, der, commonmodel)  # configure composite model
        addDATs(der, dername, derprotocol, commonmodel, dercube)  # create DATs
        addsignalvariables(allCalc, der, commonmodel)  # create signal variables
        app.PrintInfo('----------------------------------------------------------------------------------')

totalTime = time.time() - startTime        #calculate run time
totalMin = int(totalTime/60)
totalSec = float(int((totalTime % 60)*100)/100)

app.PrintInfo('')
app.PrintInfo('----------------------------------------------------------------------------------')
app.PrintInfo('                                Done Scripting!!!                                 ')
app.PrintInfo('                         Run time: ' + str(totalMin) + ' minute(s) '
                                              + str(totalSec) + ' second(s)                      ')
app.PrintInfo('----------------------------------------------------------------------------------')
app.PrintInfo('')
