# -*- coding: utf-8 -*-

""" ***************************************************************************

Python Script Name : OrcaPySM1A

Description : 
    
    Python Scripts to Automate OrcaFlex Simulation
    For Spread Moored Vessels
    
    Package Includes Two Python Script Files with an Input Excel WorkBook. The 
    Following are the names:
    
        Input.xlsx
        OrcaPySM1A.py
        OrcaPySM1B.py
    
    Note: The Wave Loads on the vessel are required to be imported seperately 
    from an OrcaWave Result File or any other valid / compatible seakeeping 
    tool.
    
    Read Below steps on how to use these scripts 
    
    Step 1: 
    -------
        Prepare the Input Excel Work Book "Inputs.xlsx"
    Fill the Input Excel sheets with appropraite details to generate the 
    OrcaFlex Model of the Spread Moored Vessel System. The parameters and their
    Description are given in the Template sheet provided with this script
    This Input Excel File and the Python scripts are required to be in the 
    same directory. Lets call this Directory as Parent Directory.

    Step 2:
    -------
        Run the Python Script : OrcaPySM1.py
    If all input data is sufficient and valid, then this script generates 
    a Folder named INTACT in the same parent directory.
    This INTACT folder shall have the following  
    1. OrcaFlex Model Data File with .yml extension
    2. OrcaFlex Simulation File with adjusted Mooring Lines & statics analysed
        
    Step 3:
    -------
        Import the Vessel Sea Keeping Analysis Results (RAOs & QTFs)
    Open the Generated Simulation File, In the Vessel Type data import the
    Wave Loads  Motions from the Sea keeping analysis results file of 
    OrcaWave. Make sure the Conventions are consistent. Once Everything the 
    Vessel Type and Vessel are found to be modelled correctly, then red do 
    static analysis and save the simulation file with the same name in the same
    INTACT directory.
    
    Step 4: 
    -------
        Run the Python Script : OrcaPySM2.py
    This Script file will be in the parent directory. This cript reads the 
    Input Excel Sheet and the Intact Static Simulation File. Based on the Cases
    listed in the Input Excel Sheet this script generates one Simulation File  
    for each of the Intact Dynamic Analysis Case. This SCript shall also 
    read the Damage Cases list and generates a seperate folder named 
    "DAMAGE" in the parent directory, with simulation files corresponding to
    each of the single line damage analysis cases
    
    These Generated Files can be Batch Processed and the final simulation 
    results can be further post processed.
    
@author: Praveen Kumar Ch (praveench1888@gmail.com)

*************************************************************************** """

import OrcFxAPI
import numpy as np
import pandas as pd
import math
import os
import shutil

''' ---------------------------------------------------------------------------
    Name of the Input Excel File
--------------------------------------------------------------------------- '''
INPUT_FILE = 'Input.xlsx'

# Function to create a valid file name
def filename_valid(filename):
    invalid = '<>:"/\|?* '
    for char in invalid:
    	filename = filename.replace(char, '')	
    return filename

INTACT_DIR = 'INTACT'

if os.path.exists(INTACT_DIR):
    shutil.rmtree(INTACT_DIR)

os.mkdir(INTACT_DIR)

''' ---------------------------------------------------------------------------
    CREATE MAIN MODEL OBJECT
--------------------------------------------------------------------------- '''
# The main Orcaflex MODEL Object
model_0 = OrcFxAPI.Model()

''' ---------------------------------------------------------------------------
                  General Analysis Data
--------------------------------------------------------------------------- '''
gen=model_0.general
gen.DynamicsSolutionMethod = 'Implicit time domain'
gen.StageCount = 2
gen.StageDuration[0]=8.0
gen.StageDuration[1]=3600


''' ---------------------------------------------------------------------------
                    INITIAL ENVIRONMENT SETUP
--------------------------------------------------------------------------- '''

# Reading Data from Input Excel File - General Sheet
DF_GN = pd.read_excel(INPUT_FILE, sheet_name='General', index_col=0, usecols='A:B', header=1)

GRS = DF_GN.VAL['GRS']
GXDIR = DF_GN.VAL['GXDIR']

# Location Identification Tag
LOC_TAG = filename_valid(DF_GN.VAL['LOC_TAG'])

# Setting Calm Sea Environement
model_0.environment.WaterDepth = DF_GN.VAL['SEA_DEPTH']
model_0.environment.NumberOfWaveTrains = 1
model_0.environment.WaveDirection = 0
model_0.environment.WaveHeight = 0
model_0.environment.WaveType = 'Airy'
model_0.environment.RefCurrentSpeed = 0
model_0.environment.RefCurrentDirection = 0
model_0.environment.WindSpeed = 0
model_0.environment.WindDirection = 0


''' ---------------------------------------------------------------------------
    CREATE VESSEL TYPE OBJECT
--------------------------------------------------------------------------- '''

# Reading Vessel General Data from Iput excel sheet
DF_VES_GEN = pd.read_excel(INPUT_FILE, sheet_name='Ves_Gen', index_col=0, usecols='A:B',header=1)

# Vessel Identification Tag
VES_TAG = filename_valid(DF_VES_GEN.VAL['TAG'])

# Create Vessel Type Object
vesselType_0 = model_0.CreateObject(OrcFxAPI.ObjectType.VesselType, DF_VES_GEN.VAL['TYPE'])

# Structure of vesselType_0
vesselType_0.Length = DF_VES_GEN.VAL['LENGTH']
vesselType_0.Mass = DF_VES_GEN.VAL['MASS']
vesselType_0.MomentOfInertiaTensorX[0] = vesselType_0.Mass * DF_VES_GEN.VAL['Kxx']**2
vesselType_0.MomentOfInertiaTensorY[1] = vesselType_0.Mass * DF_VES_GEN.VAL['Kyy']**2
vesselType_0.MomentOfInertiaTensorZ[2] = vesselType_0.Mass * DF_VES_GEN.VAL['Kzz']**2
vesselType_0.MomentOfInertiaTensorY[0] = vesselType_0.Mass * DF_VES_GEN.VAL['Kxy']**2
vesselType_0.MomentOfInertiaTensorZ[1] = vesselType_0.Mass * DF_VES_GEN.VAL['Kyz']**2
vesselType_0.MomentOfInertiaTensorZ[0] = vesselType_0.Mass * DF_VES_GEN.VAL['Kxz']**2

vesselType_0.CentreOfMassX = DF_VES_GEN.VAL['LCG']
vesselType_0.CentreOfMassY = DF_VES_GEN.VAL['TCG']
vesselType_0.CentreOfMassZ = DF_VES_GEN.VAL['VCG']

vesselType_0.StiffnessInertiaDampingRefOriginx = DF_VES_GEN.VAL['LENGTH']/2
vesselType_0.StiffnessInertiaDampingRefOriginy = 0
vesselType_0.StiffnessInertiaDampingRefOriginz = DF_VES_GEN.VAL['DRAFT']

vesselType_0.HydrostaticReferenceOriginDatumPositionz = 0
vesselType_0.HydrostaticReferenceOriginDatumOrientationx = 0
vesselType_0.HydrostaticReferenceOriginDatumOrientationy = 0

# Penning Vessel Geometry with Rectangular Box Geometry
if DF_VES_GEN.VAL['XREF'] == 'FP':
    XFWD = 0
    if DF_VES_GEN.VAL['XDIR'] == 'AFT':
        XAFT = DF_VES_GEN.VAL['LENGTH']
    else:
        XAFT = -DF_VES_GEN.VAL['LENGTH']
elif DF_VES_GEN.VAL['XREF'] == 'AP':
    XAFT = 0
    if DF_VES_GEN.VAL['XDIR'] == 'FWD':
        XFWD = DF_VES_GEN.VAL['LENGTH']
    else:
        XFWD = -DF_VES_GEN.VAL['LENGTH']
else:
    if DF_VES_GEN.VAL['XDIR'] == 'FWD':
        XAFT = -DF_VES_GEN.VAL['LENGTH']/2
        XFWD = DF_VES_GEN.VAL['LENGTH']/2
    else:
        XAFT = DF_VES_GEN.VAL['LENGTH']/2
        XFWD = -DF_VES_GEN.VAL['LENGTH']/2

if DF_VES_GEN.VAL['XDIR'] == 'FWD':
        YSTBD = -DF_VES_GEN.VAL['BREADTH']/2
        YPORT = DF_VES_GEN.VAL['BREADTH']/2
elif DF_VES_GEN.VAL['XDIR'] == 'AFT':
        YPORT = -DF_VES_GEN.VAL['BREADTH']/2
        YSTBD = DF_VES_GEN.VAL['BREADTH']/2

if DF_VES_GEN.VAL['ZREF'] == 'BL':
    ZBL = 0
    ZMD = DF_VES_GEN.VAL['DEPTH']
elif DF_VES_GEN.VAL['ZREF'] == 'DRAFT':
    ZBL = -DF_VES_GEN.VAL['DRAFT']
    ZMD = DF_VES_GEN.VAL['DEPTH']-DF_VES_GEN.VAL['DRAFT']
else:
    ZBL = -DF_VES_GEN.VAL['DRAFT']
    ZMD = DF_VES_GEN.VAL['DEPTH']-DF_VES_GEN.VAL['DRAFT']

vesselType_0.WireFrameType = 'Edges'

vesselType_0.NumberOfVertices = 8
vesselType_0.VertexX[0] = XAFT
vesselType_0.VertexX[1] = XAFT
vesselType_0.VertexX[2] = XAFT
vesselType_0.VertexX[3] = XAFT
vesselType_0.VertexX[4] = XFWD
vesselType_0.VertexX[5] = XFWD
vesselType_0.VertexX[6] = XFWD
vesselType_0.VertexX[7] = XFWD

vesselType_0.VertexY[0] = YPORT
vesselType_0.VertexY[1] = YPORT
vesselType_0.VertexY[2] = YSTBD
vesselType_0.VertexY[3] = YSTBD
vesselType_0.VertexY[4] = YPORT
vesselType_0.VertexY[5] = YPORT
vesselType_0.VertexY[6] = YSTBD
vesselType_0.VertexY[7] = YSTBD

vesselType_0.VertexZ[0] = ZMD
vesselType_0.VertexZ[1] = ZBL
vesselType_0.VertexZ[2] = ZBL
vesselType_0.VertexZ[3] = ZMD
vesselType_0.VertexZ[4] = ZMD
vesselType_0.VertexZ[5] = ZBL
vesselType_0.VertexZ[6] = ZBL
vesselType_0.VertexZ[7] = ZMD

vesselType_0.NumberOfEdges = 12
vesselType_0.EdgeFrom[0], vesselType_0.EdgeTo[0] = 1, 2
vesselType_0.EdgeFrom[1], vesselType_0.EdgeTo[1] = 2, 3
vesselType_0.EdgeFrom[2], vesselType_0.EdgeTo[2] = 3, 4
vesselType_0.EdgeFrom[3], vesselType_0.EdgeTo[3] = 4, 1
vesselType_0.EdgeFrom[4], vesselType_0.EdgeTo[4] = 5, 6
vesselType_0.EdgeFrom[5], vesselType_0.EdgeTo[5] = 6, 7
vesselType_0.EdgeFrom[6], vesselType_0.EdgeTo[6] = 7, 8
vesselType_0.EdgeFrom[7], vesselType_0.EdgeTo[7] = 8, 5
vesselType_0.EdgeFrom[8], vesselType_0.EdgeTo[8] = 1, 5
vesselType_0.EdgeFrom[9], vesselType_0.EdgeTo[9] = 2, 6
vesselType_0.EdgeFrom[10], vesselType_0.EdgeTo[10] = 3, 7
vesselType_0.EdgeFrom[11], vesselType_0.EdgeTo[11] = 4, 8

''' Set Wind and Current Areas and Point of Action '''

# Read data from Ves_Area Sheet of Input Excel file
DF_VES_AREA = pd.read_excel(INPUT_FILE, sheet_name='Ves_Area', index_col=0, usecols='A:J',header=2, nrows=3)

vesselType_0.CurrentCoeffSurgeArea = DF_VES_AREA.SURGE_AREA['CURRENT']
vesselType_0.CurrentCoeffSwayArea = DF_VES_AREA.SWAY_AREA['CURRENT']
vesselType_0.CurrentCoeffHeaveArea = DF_VES_AREA.HEAVE_AREA['CURRENT']
vesselType_0.CurrentCoeffRollAreaMoment = DF_VES_AREA.ROLL_AREAMOM['CURRENT']
vesselType_0.CurrentCoeffPitchAreaMoment = DF_VES_AREA.PITCH_AREAMOM['CURRENT']
vesselType_0.CurrentCoeffYawAreaMoment = DF_VES_AREA.YAW_AREAMOM['CURRENT']

vesselType_0.CurrentCoeffOriginX = DF_VES_AREA.X_ORG['CURRENT']
vesselType_0.CurrentCoeffOriginY = DF_VES_AREA.Y_ORG['CURRENT']
vesselType_0.CurrentCoeffOriginZ = DF_VES_AREA.Z_ORG['CURRENT']

vesselType_0.WindCoeffSurgeArea = DF_VES_AREA.SURGE_AREA['WIND']
vesselType_0.WindCoeffSwayArea = DF_VES_AREA.SWAY_AREA['WIND']
vesselType_0.WindCoeffHeaveArea = DF_VES_AREA.HEAVE_AREA['WIND']
vesselType_0.WindCoeffRollAreaMoment = DF_VES_AREA.ROLL_AREAMOM['WIND']
vesselType_0.WindCoeffPitchAreaMoment = DF_VES_AREA.PITCH_AREAMOM['WIND']
vesselType_0.WindCoeffYawAreaMoment = DF_VES_AREA.YAW_AREAMOM['WIND']

vesselType_0.WindCoeffOriginX = DF_VES_AREA.X_ORG['WIND']
vesselType_0.WindCoeffOriginY = DF_VES_AREA.Y_ORG['WIND']
vesselType_0.WindCoeffOriginZ = DF_VES_AREA.Z_ORG['WIND']

''' Set Wind and Current Load Coefficients '''

# Read Vessel Current Coefficients from Input Excel File
DF_VES_CURR = pd.read_excel(INPUT_FILE, sheet_name='Ves_Curr', index_col=None, usecols='A:G',header=1)

nCD = len(DF_VES_CURR)

vesselType_0.CurrentCoeffSymmetry='xz plane'
vesselType_0.NumberOfCurrentCoeffDirections=nCD

for i in range(nCD):
    vesselType_0.CurrentCoeffDirection[i]=DF_VES_CURR.DIR[i]
    vesselType_0.CurrentCoeffSurge[i]=DF_VES_CURR.SURGE[i]
    vesselType_0.CurrentCoeffSway[i]=DF_VES_CURR.SWAY[i]
    vesselType_0.CurrentCoeffHeave[i]=DF_VES_CURR.HEAVE[i]
    vesselType_0.CurrentCoeffRoll[i]=DF_VES_CURR.ROLL[i]
    vesselType_0.CurrentCoeffPitch[i]=DF_VES_CURR.PITCH[i]
    vesselType_0.CurrentCoeffYaw[i]=DF_VES_CURR.YAW[i]

DF_VES_WIND = pd.read_excel(INPUT_FILE, sheet_name='Ves_Wind', index_col=None, usecols='A:G',header=1)

nWD = len(DF_VES_WIND)

vesselType_0.WindCoeffSymmetry='xz plane'
vesselType_0.NumberOfWindCoeffDirections=nWD

for i in range(nWD):
    vesselType_0.WindCoeffDirection[i]=DF_VES_WIND.DIR[i]
    vesselType_0.WindCoeffSurge[i]=DF_VES_WIND.SURGE[i]
    vesselType_0.WindCoeffSway[i]=DF_VES_WIND.SWAY[i]
    vesselType_0.WindCoeffHeave[i]=DF_VES_WIND.HEAVE[i]
    vesselType_0.WindCoeffRoll[i]=DF_VES_WIND.ROLL[i]
    vesselType_0.WindCoeffPitch[i]=DF_VES_WIND.PITCH[i]
    vesselType_0.WindCoeffYaw[i]=DF_VES_WIND.YAW[i]

''' ---------------------------------------------------------------------------
    CREATE VESSEL OBJECT
--------------------------------------------------------------------------- '''

# Create a vessel object & set its Type, Connections, Position, Orientation
vessel_0 = model_0.CreateObject(OrcFxAPI.ObjectType.Vessel, DF_VES_GEN.VAL['NAME'])
vessel_0.type = vesselType_0.Name
vessel_0.Connection = 'Free'

# Note that OrcaFlex Reference System is always RHS
# The position of vessel defined in Global Reference Frame
vessel_0.InitialX = DF_VES_GEN.VAL['XPOS']

if DF_GN.VAL['GRS']=='RHS':
    vessel_0.InitialY = DF_VES_GEN.VAL['YPOS']   
else:
    vessel_0.InitialY = -DF_VES_GEN.VAL['YPOS']

vessel_0.InitialZ = DF_VES_GEN.VAL['ZPOS']

vessel_0.InitialHeel = DF_VES_GEN.VAL['HEEL']
vessel_0.InitialTrim = DF_VES_GEN.VAL['TRIM']

if DF_GN.VAL['GRS']=='RHS':
    vessel_0.InitialHeading = DF_VES_GEN.VAL['HEADING']  
else:
    vessel_0.InitialHeading = 360-DF_VES_GEN.VAL['HEADING']

vessel_0.IncludedInStatics = '6 DOF'

''' ---------------------------------------------------------------------------
                                LINE TYPES
--------------------------------------------------------------------------- '''

# Reading Line Type Data From Excel
DF_LT = pd.read_excel(INPUT_FILE, sheet_name='Line_Types', index_col=0, usecols='A:M', header=3)
nLT = DF_LT.shape[0]

''' ----- Creating Line Type Objects ----- '''

lineTypes = list()
for i in range(nLT):
    lineType = model_0.CreateObject(OrcFxAPI.ObjectType.LineType, name=DF_LT.index[i])
    if DF_LT.WIZARD[i]:
        if 'Rope' in DF_LT.LTYP[i] or 'wire' in DF_LT.LTYP[i]:
            lineType.WizardCalculation = DF_LT.LTYP[i]
            lineType.RopeNominalDiameter = DF_LT.NOM_DIA[i]
            lineType.RopeConstruction = DF_LT.SUBTYP[i]
            lineType.InvokeWizard()
        if 'Chain' in DF_LT.LTYP[i]:
            lineType.WizardCalculation = DF_LT.LTYP[i]
            lineType.ChainBarDiameter = DF_LT.NOM_DIA[i]
            lineType.ChainLinkType = DF_LT.SUBTYP[i]
            lineType.InvokeWizard()
    lineTypes.append(lineType)
    


'''---------------------------------------------------------------------------
    Creating Clump Types (Buoys)
---------------------------------------------------------------------------'''
DF_CB = pd.read_excel(INPUT_FILE, sheet_name='Clump_Buoy', index_col=0, usecols='A:E', header=3)
nCB = DF_CB.shape[0]

for i in range(nCB):
    clumpType = model_0.CreateObject(OrcFxAPI.ObjectType.ClumpType, DF_CB.index[i])
    clumpType.Mass = DF_CB.MASS[i]
    clumpType.Volume = DF_CB.VOLUME[i]
    clumpType.Height = DF_CB.HEIGHT[i]
    clumpType.Offset = DF_CB.OFFSET[i]
    clumpType.AlignWith = 'Global axes'
    clumpType.PenWidth = 10

''' ---------------------------------------------------------------------------
    MOORING LINES
--------------------------------------------------------------------------- '''

# Reading Fairlead Locations from Excel Sheet
DF_FL = pd.read_excel(INPUT_FILE, sheet_name='Ves_FL', index_col=0, header=2,usecols='A:D')

# Reading Initial Mooring Line Configuration from Excel Sheet
DF_ML = pd.read_excel(INPUT_FILE, sheet_name='Moor_Lines', index_col=0, usecols='A:AC',header=3)

nLines = DF_ML.shape[0]
lines = list()

for i in range(nLines):

    # Creating Line object
    line = model_0.CreateObject(OrcFxAPI.ObjectType.Line, name=DF_ML.index[i])

    # Set General parameters
    line.IncludeTorsion = 'No'
    line.TopEnd = 'End A'
    line.Representation = 'Finite element'
    line.LengthAndEndOrientations = 'Explicit'

    # End A connection - Top End
    line.EndAConnection = vessel_0.Name
    FLID = DF_ML.ENDA_CONN[i]
    line.EndAX = DF_FL.X_FL[FLID]
    
    if DF_VES_GEN.VAL['VRS'] == 'RHS':
        line.EndAY = DF_FL.Y_FL[FLID]
    else:
        line.EndAY = - DF_FL.Y_FL[FLID]
            
    line.EndAZ = DF_FL.Z_FL[FLID]

    # End B Connections - Bottom
    line.EndBConnection = DF_ML.ENDB_CONN[i]
    COS_HEADING = math.cos(math.radians(vessel_0.InitialHeading))
    SIN_HEADING = math.sin(math.radians(vessel_0.InitialHeading))

    xBV = DF_FL.X_FL[FLID]+DF_ML.HORZ_DIST[i] * \
        math.cos(math.radians(DF_ML.AZIMUTH[i]))
    yBV = DF_FL.Y_FL[FLID]+DF_ML.HORZ_DIST[i] * \
        math.sin(math.radians(DF_ML.AZIMUTH[i]))

    xBG1 = xBV*COS_HEADING-yBV*SIN_HEADING
    yBG1 = xBV*SIN_HEADING+yBV*COS_HEADING

    line.EndBX = vessel_0.InitialX+xBG1
    line.EndBY = vessel_0.InitialY+yBG1
    line.EndBZ = 0

    if line.EndBConnection == 'Anchored':
        line.EndBHeightAboveSeabed = DF_ML.VERT_POS[i]+model_0.environment.WaterDepth

    if line.EndBConnection == 'Fixed':
        line.EndBZ = DF_ML.VERT_POS[i]

    # Sections and Line Types
    line.NumberOfSections = int(DF_ML.N_SECS[i])
    for j in range(line.NumberOfSections):
        LTID = DF_ML.iloc[i, 16+(j*3)]
        SEC_LEN = DF_ML.iloc[i, 16+(j*3)+1]
        SEG_LEN = DF_ML.iloc[i, 16+(j*3)+2]
        line.LineType[j] = model_0[LTID].Name
        line.Length[j] = SEC_LEN
        line.TargetSegmentLength[j] = SEG_LEN

    # Buoys and Clump weights
    line.NumberOfAttachments = int(DF_ML.N_BUOYS[i])
    for j in range(line.NumberOfAttachments):
        CBID = DF_ML.iloc[i, 9+(j*2)]
        SEG_LEN = DF_ML.iloc[i, 9+(j*2)+1]
        line.AttachmentType[j] = model_0[CBID].Name
        line.Attachmentz[j] = SEG_LEN
        line.AttachmentzRelativeTo[j] = 'End B'

    line.SetLayAzimuth = 'Yes'
    lines.append(line)


''' ---------------------------------------------------------------------------
 Saving the Intact Initial SET UP 
------------------------------------------------------------------------------ '''

BASENAME=VES_TAG+'_'+LOC_TAG

fileName = os.path.join(INTACT_DIR, BASENAME+'_INIT_SETUP.yml')
model_0.SaveData(fileName)


''' ---------------------------------------------------------------------------
    LINE SETUP WIZARD
--------------------------------------------------------------------------- '''

# Settings for Line Setup Wizard
ILSW = 0
model_0.general.LineSetupCalculationMode = "Calculate line lengths"
model_0.general.LineSetupMaxDamping = 20
model_0.general.LineSetupTolerance = 0.01


for i in range(nLines):
    j = i
    # For Line Setup wizard : Pre tension
    if (DF_ML.LAY_SETUP[i] == "PRE_TENS"):
        ILSW = 1
        lines[j].LineSetupIncluded = 'Yes'
        lines[j].LineSetupTargetVariable = 'Tension'
        lines[j].LineSetupLineEnd = 'End A'
        lines[j].LineSetupArclength = 0.0
        lines[j].LineSetupTargetValue = DF_ML.PRE_TENS[i]
    else:
        lines[j].LineSetupIncluded = 'Yes'
        lines[j].LineSetupTargetVariable = 'No target'

vessel_0.IncludedInStatics = 'None'

if ILSW==1:
    model_0.InvokeLineSetupWizard()

vessel_0.IncludedInStatics = '6 DOF'

model_0.CalculateStatics()

''' ---------------------------------------------------------------------------
 Saving the STATIC Simulation Setup
------------------------------------------------------------------------------ '''
fileName = os.path.join(INTACT_DIR, BASENAME+'_INTACT_STATICS.sim')
model_0.SaveSimulation(fileName)

