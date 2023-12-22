# -*- coding: utf-8 -*-
""" ***************************************************************************

Python Script Name : OrcaPySM1B

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

DAMAGE_DIR = 'DAMAGE'

if os.path.exists(DAMAGE_DIR):
    shutil.rmtree(DAMAGE_DIR)

os.mkdir(DAMAGE_DIR)

''' ---------------------------------------------------------------------------
Load the Data from existing Intact Static File
----------------------------------------------------------------------------'''

# Reading Data from Input Excel File - General Sheet
DF_GN = pd.read_excel(INPUT_FILE, sheet_name='General', index_col=0, usecols='A:B', header=1)

GRS = DF_GN.VAL['GRS']
GXDIR = DF_GN.VAL['GXDIR']

# Location Identification Tag
LOC_TAG = filename_valid(DF_GN.VAL['LOC_TAG'])

# Reading Vessel General Data from Iput excel sheet
DF_VES_GEN = pd.read_excel(INPUT_FILE, sheet_name='Ves_Gen', index_col=0, usecols='A:B',header=1)

# Vessel Identification Tag
VES_TAG = filename_valid(DF_VES_GEN.VAL['TAG'])

BASENAME=VES_TAG+'_'+LOC_TAG
INTACT_DIR = 'INTACT'

fileName = os.path.join(INTACT_DIR, BASENAME+'_INTACT_STATICS.sim')

model_0 = OrcFxAPI.Model(fileName)


''' --------------------------------------------------------------------------
Intact Dynamic Setup
----------------------------------------------------------------------------'''

#
vessel_0 = model_0[DF_VES_GEN.VAL['NAME']]
vessel_0.IncludedInStatics = '6 DOF'
vessel_0.PrimaryMotion = 'Calculated (6 DOF)'
vessel_0.SuperimposedMotion = 'None'
vessel_0.IncludeAppliedLoads = 'No'
vessel_0.IncludeWaveLoad1stOrder = 'Yes'
vessel_0.IncludeWaveDriftLoad2ndOrder = 'Yes'
vessel_0.IncludeWaveDriftDamping = 'Yes'
vessel_0.IncludeSumFrequencyLoad = 'No'
vessel_0.IncludeAddedMassAndDamping = 'Yes'
vessel_0.IncludeManoeuvringLoad = 'Yes'
vessel_0.IncludeOtherDamping = 'Yes'
vessel_0.IncludeCurrentLoad = 'Yes'
vessel_0.IncludeWindLoad = 'Yes'
vessel_0.PrimaryMotionIsTreatedAs = 'Both low and wave frequency'
vessel_0.PrimaryMotionDividingPeriod = 40.0
vessel_0.CalculationMode = 'Filtering'
vessel_0.CalculateHydrostaticStiffnessAnglesBy = 'Orientation'
#


# Reading the intact Case Matrix from Input Excel sheet
DF_ICM = pd.read_excel(INPUT_FILE, sheet_name='IntactCases', index_col=None, usecols='A:J',header=3)

# Loop for each Case
nICM = len(DF_ICM)

env = model_0.environment

for ic in range(nICM):

    # Computing Direction
    DIR_REF=DF_ICM.DIR_REF[ic]
    DIR_CONV = DF_ICM.DIR_CONV[ic]
    
    if DIR_REF == 'GLOBX':
        LAG_ANGLE = 0
    elif DIR_REF == 'NORTH':
        LAG_ANGLE = GXDIR
    elif DIR_REF == 'EAST':
        LAG_ANGLE = GXDIR-90
    elif DIR_REF == 'SOUTH':
        LAG_ANGLE = GXDIR-180
    elif DIR_REF == 'WEST':
        LAG_ANGLE = GXDIR-270
    elif DIR_REF == 'VESX+':
        LAG_ANGLE = vessel_0.InitialHeading
    elif DIR_REF == 'VESX-':
        LAG_ANGLE = vessel_0.InitialHeading + 180
        
    DIR_ORCA = list()
    
        
    if DIR_CONV == 'ANTICLOCKWISE':
        TEMP1=DF_ICM.DIR[ic]
    else:
        TEMP1=360-DF_ICM.DIR[ic]
    
    TEMP2=TEMP1+LAG_ANGLE

    DIRECTION = (TEMP2%360)

    # Wave Params
    model_0.environment.NumberOfWaveTrains = 1
    env.WaveType = DF_ICM.WAVE_TYPE[ic]
    env.WaveDirection = DIRECTION

    if env.WaveType == 'JONSWAP' or env.WaveType == 'ISSC':
        env.WaveHs = DF_ICM.Hs[ic]
        env.WaveTp = DF_ICM.Tp[ic]
        env.WaveGamma = DF_ICM.GAMMA[ic]
    else:
        env.WaveHeight = DF_ICM.Hs[ic]
        env.WavePeriod = DF_ICM.Tp[ic]

    # Wind Params
    env.WindDirection = DIRECTION
    env.WindSpeed = DF_ICM.Vw[ic]

    # Current Params
    env.RefCurrentSpeed = DF_ICM.Vc[ic]
    env.RefCurrentDirection = DIRECTION

    # Save the File
    fileName = os.path.join(INTACT_DIR, BASENAME +
                            '_INTACT_DYNAMICS_'+str(DF_ICM.CASE_ID[ic]).replace(' ', '_')+'.sim')

    model_0.CalculateStatics()

    model_0.SaveSimulation(fileName)
    
    

''' --------------------------------------------------------------------------
DAMAGE Dynamic Setup
----------------------------------------------------------------------------'''

# Reading the intact Case Matrix from Input Excel sheet
DF_DCM = pd.read_excel(INPUT_FILE, sheet_name='DamageCases', index_col=None, usecols='A:K',header=3)

# Loop for each Case
nDCM = len(DF_DCM)
''
for ic in range(nDCM):
    
    # Opening the Intact Static File
    del model_0
    fileName = os.path.join(INTACT_DIR, BASENAME+'_INTACT_STATICS.sim')
    model_0=OrcFxAPI.Model(fileName)
    
    # Creating a Copy of the sim file for damage case
    fileName = os.path.join(DAMAGE_DIR, BASENAME +'_DAMAGE_DYNAMICS_'+str(DF_DCM.CASE_ID[ic]).replace(' ', '_')+'.sim')
    model_0.SaveSimulation(fileName)

    model_0.DestroyObject(DF_DCM.DAM_LIN[ic])
    
    vessel_0 = model_0[DF_VES_GEN.VAL['NAME']]
    vessel_0.IncludedInStatics = '6 DOF'
    vessel_0.PrimaryMotion = 'Calculated (6 DOF)'
    vessel_0.SuperimposedMotion = 'None'
    vessel_0.IncludeAppliedLoads = 'No'
    vessel_0.IncludeWaveLoad1stOrder = 'Yes'
    vessel_0.IncludeWaveDriftLoad2ndOrder = 'Yes'
    vessel_0.IncludeWaveDriftDamping = 'Yes'
    vessel_0.IncludeSumFrequencyLoad = 'No'
    vessel_0.IncludeAddedMassAndDamping = 'Yes'
    vessel_0.IncludeManoeuvringLoad = 'Yes'
    vessel_0.IncludeOtherDamping = 'Yes'
    vessel_0.IncludeCurrentLoad = 'Yes'
    vessel_0.IncludeWindLoad = 'Yes'
    vessel_0.PrimaryMotionIsTreatedAs = 'Both low and wave frequency'
    vessel_0.PrimaryMotionDividingPeriod = 40.0
    vessel_0.CalculationMode = 'Filtering'
    vessel_0.CalculateHydrostaticStiffnessAnglesBy = 'Orientation'
    
    env = model_0.environment
    
    
    # Computing Direction
    DIR_REF=DF_DCM.DIR_REF[ic]
    DIR_CONV = DF_DCM.DIR_CONV[ic]
    
    if DIR_REF == 'GLOBX':
        LAG_ANGLE = 0
    elif DIR_REF == 'NORTH':
        LAG_ANGLE = GXDIR
    elif DIR_REF == 'EAST':
        LAG_ANGLE = GXDIR-90
    elif DIR_REF == 'SOUTH':
        LAG_ANGLE = GXDIR-180
    elif DIR_REF == 'WEST':
        LAG_ANGLE = GXDIR-270
    elif DIR_REF == 'VESX+':
        LAG_ANGLE = vessel_0.InitialHeading
    elif DIR_REF == 'VESX-':
        LAG_ANGLE = vessel_0.InitialHeading + 180
        
    DIR_ORCA = list()
    
        
    if DIR_CONV == 'ANTICLOCKWISE':
        TEMP1=DF_DCM.DIR[ic]
    else:
        TEMP1=360-DF_DCM.DIR[ic]
    
    TEMP2=TEMP1+LAG_ANGLE

    DIRECTION = (TEMP2%360)

    # Wave Params
    model_0.environment.NumberOfWaveTrains = 1
    env.WaveType = DF_DCM.WAVE_TYPE[ic]
    env.WaveDirection = DIRECTION

    if env.WaveType == 'JONSWAP' or env.WaveType == 'ISSC':
        env.WaveHs = DF_DCM.Hs[ic]
        env.WaveTp = DF_DCM.Tp[ic]
    else:
        env.WaveHeight = DF_DCM.Hs[ic]
        env.WavePeriod = DF_DCM.Tp[ic]

    # Wind Params
    env.WindDirection = DIRECTION
    env.WindSpeed = DF_DCM.Vw[ic]

    # Current Params
    env.RefCurrentSpeed = DF_DCM.Vc[ic]
    env.RefCurrentDirection = DIRECTION

    
    model_0.CalculateStatics()

    model_0.SaveSimulation(fileName)





