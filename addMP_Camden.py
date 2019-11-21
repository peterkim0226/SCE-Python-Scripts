#author: Peter Kim (Peter.1.Kim@sce.com)
#date created: 4/24/2019
#last updated: 6/14/2019
#This script will located pre-existing monitoring points that were added to PowerFactory and will add signal variables and DAT points to it
#***NOTE*** This script DOES NOT add monitoring points. The monitoring points must already exist in the model as cubicles

import powerfactory

imeas_mult = 1              #Multiplicator for Imeas
imeas_deadband = 0.01           #Deadband for Imeas
vmeas_mult = 0.00001        #Multiplicator for Vmeas
vmeas_deadband = 0.1            #Deadband for Vmeas
pmeas_mult = -1             #Multiplicator for Pmeas
pmeas_deadband = 0.01           #Deadband for Pmeas
qmeas_mult = -1             #Multiplicator for Qmeas
qmeas_deadband = 0.001          #Deadband for Qmeas
status_mult = 1             #Multiplicator for Status
status_deadband = 0.01          #Deadband for Status

networkName = '*.ElmNet'

# create powerfactory app object
app = powerfactory.GetApplication()
netData = app.GetProjectFolder('netdat')
sysGrid = netData.GetChildren(0, networkName, 0)[0]  # This returns the Cobolt Network

# study case folder
studyCaseFolder = app.GetProjectFolder('study')
# gets the first (default) study case of the study case folder (assumes that no custom study case exists)
studyCase = studyCaseFolder.GetContents()[0]
# gets the 'All Calculations' Folder
allCalc = studyCase.GetChildren(0, '*.ElmRes', 0)[0]

for cube in app.GetCalcRelevantObjects('MP*.StaCubic'):
    # if 'Ti' not in cube.loc_name:     #Replace Ti with another circuit to get monitoring points from that circuit
    #     continue
    cubeName = cube.loc_name
    busOfCube = cube.cterm

    app.PrintInfo('Adding measurements to cubicle ' + cubeName)

    # add Vmeas (non-external)
    newVmeas = sysGrid.CreateObject('StaVmea', 'Vmeas ' + cubeName)
    newVmeas.i_mode = 1
    newVmeas.pbusbar = busOfCube

    ImeasDAT = cube.CreateObject('StaExtdatmea', cubeName + '_Imeas')
    ImeasDAT.Multip = imeas_mult  # Multiplicator
    ImeasDAT.deadband = imeas_deadband  # OPC deadband
    ImeasDAT.i_dat = 3  # DAT type = REAL
    ImeasDAT.iStatus = 1073741824  # Status
    ImeasDAT.pCalObj = ImeasDAT.ccubic
    ImeasDAT.pCalObjSim = ImeasDAT.ccubic
    ImeasDAT.varCal = 'm:I:bus1'  # Variable Name (Object)
    ImeasDAT.varCalSim = 'm:I:bus1'  # Variable Name (Sim Object)
    ImeasDAT.sTagID = cubeName + '_Imeas'  # TagID

    VmeasDAT = cube.CreateObject('StaExtvmea', cubeName + '_Vmeas')
    VmeasDAT.Multip = vmeas_mult  # Multiplicator
    VmeasDAT.deadband = vmeas_deadband  # OPC deadband
    VmeasDAT.i_dat = 3  # DAT type = REAL
    VmeasDAT.iStatus = 1073741824  # Status
    VmeasDAT.sTagID = cubeName + '_Vmeas'  # TagID

    PmeasDAT = cube.CreateObject('StaExtdatmea', cubeName + '_Pmeas')
    PmeasDAT.Multip = pmeas_mult  # Multiplicator
    PmeasDAT.deadband = pmeas_deadband  # OPC deadband
    PmeasDAT.i_dat = 3  # DAT type = REAL
    PmeasDAT.iStatus = 1073741824  # Status
    PmeasDAT.pCalObj = PmeasDAT.ccubic
    PmeasDAT.pCalObjSim = PmeasDAT.ccubic
    PmeasDAT.varCal = 'm:Psum:bus1'  # Variable Name (Object)
    PmeasDAT.varCalSim = 'm:Psum:bus1'  # Variable Name (Sim Object)
    PmeasDAT.sTagID = cubeName + '_Pmeas'  # TagID

    QmeasDAT = cube.CreateObject('StaExtdatmea', cubeName + '_Qmeas')
    QmeasDAT.Multip = qmeas_mult  # Multiplicator
    QmeasDAT.deadband = qmeas_deadband  # OPC deadband
    QmeasDAT.i_dat = 3  # DAT type = REAL
    QmeasDAT.iStatus = 1073741824  # Status
    QmeasDAT.pCalObj = QmeasDAT.ccubic
    QmeasDAT.pCalObjSim = QmeasDAT.ccubic
    QmeasDAT.varCal = 'm:Qsum:bus1'  # Variable Name (Object)
    QmeasDAT.varCalSim = 'm:Qsum:bus1'  # Variable Name (Sim Object)
    QmeasDAT.sTagID = cubeName + '_Qmeas'  # TagID

    app.PrintInfo('Measurements added to cube ' + cubeName)
    app.PrintInfo('Adding variables to ---> ' + cubeName)

    # new variable selection variable to be added
    newVar = allCalc.CreateObject('IntMon', cubeName)
    # referenced
    newVar.obj_id = cube.obj_id
    # variables of the object to be selected
    newVar.vars = ['m:P:bus1', 'm:Q:bus1', 'm:U1l:bus1', 'm:cosphi:bus1', 'e:pf_recap', 'm:I:bus1', 'e:outserv', 'm:Psum:bus1', 'm:Qsum:bus1']

    app.PrintInfo('Variables added to ---> ' + cubeName)
    app.PrintInfo('-----------------------------------------------------------------------------')

app.PrintInfo('Script Complete :)')
