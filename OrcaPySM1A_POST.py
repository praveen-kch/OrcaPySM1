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
        OrcaPySM1A_POST.py
        OrcaPySM1B.py
        OrcaPySM1B_POST.py
    
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
        Run the Python Script : OrcaPySM1A.py
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
    
        Run the Python Script : OrcaPySM1A_POST.py
        
        This will generate an Ouput Excel Sheet with Static Analysis Results.
    
    Step 4: 
    -------
        Run the Python Script : OrcaPySM1B.py
    This Script file will be in the parent directory. This cript reads the 
    Input Excel Sheet and the Intact Static Simulation File. Based on the Cases
    listed in the Input Excel Sheet this script generates one Simulation File  
    for each of the Intact Dynamic Analysis Case. This SCript shall also 
    read the Damage Cases list and generates a seperate folder named 
    "DAMAGE" in the parent directory, with simulation files corresponding to
    each of the single line damage analysis cases
    
    These Generated Files can be Batch Processed and the final simulation 
    results can be further post processed.
    
    After Runing all the Intact dynamic simulations:
        
        Run the Python Script : OrcaPySM1B_POST.py
        
        This will generate an Ouput Excel Sheet with Dynamic Analysis Results.   
    
@author: Praveen Kumar Ch (praveench1888@gmail.com)

*************************************************************************** """
import OrcFxAPI
import numpy as np
import pandas as pd
import math
import os
import shutil

# Function to create a valid file name
def filename_valid(filename):
    invalid = '<>:"/\|?* '
    for char in invalid:
    	filename = filename.replace(char, '')	
    return filename

''' ---------------------------------------------------------------------------
    Name of the Input Excel File
--------------------------------------------------------------------------- '''
INPUT_FILE = 'Input.xlsx'

INTACT_DIR = 'INTACT'
DAMAGE_DIR = 'DAMAGE'

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


DF_ML = pd.read_excel(INPUT_FILE, sheet_name='Moor_Lines', index_col=0, usecols='A:AC',header=3)

lines = DF_ML.index
nLines=len(lines)

vesName = DF_VES_GEN.VAL['NAME']


# Reading the intact Case Matrix from Input Excel sheet
DF_ICM = pd.read_excel(INPUT_FILE, sheet_name='IntactCases', index_col=None, usecols='A:J',header=3)

# Loop for each Case
nICM = len(DF_ICM)

''' ------------------------------------------------------------------------
Intact Static Results
--------------------------------------------------------------------------'''

fileName = os.path.join(INTACT_DIR, BASENAME+'_INTACT_STATICS.sim')
model_0 = OrcFxAPI.Model(fileName)

# Fetching and Writing - Tensions
StaticLineForces = np.zeros((nLines,8),dtype=float)

for i in range(nLines):
    
    line=model_0[lines[i]]
    StaticLineForces[i,0] = line.StaticResult('Effective Tension',OrcFxAPI.oeEndA)
    StaticLineForces[i,1] = line.StaticResult('End GX force',OrcFxAPI.oeEndA)
    StaticLineForces[i,2] = line.StaticResult('End GY force',OrcFxAPI.oeEndA)
    StaticLineForces[i,3] = line.StaticResult('End GZ force',OrcFxAPI.oeEndA)
    StaticLineForces[i,4] = line.StaticResult('Effective Tension',OrcFxAPI.oeEndB)
    StaticLineForces[i,5] = line.StaticResult('End GX force',OrcFxAPI.oeEndB)
    StaticLineForces[i,6] = line.StaticResult('End GY force',OrcFxAPI.oeEndB)
    StaticLineForces[i,7] = line.StaticResult('End GZ force',OrcFxAPI.oeEndB)

DataNames = ['End A - Effective Tensions (kN)','End A - GX force (kN)', 'End A - GY force (kN)','End A - GZ force (kN)','End B - Effective Tensions (kN)','End B - GX forc (kN)', 'End B - GY force (kN)','End B - GZ force (kN)']
DF_ISR_TEN = pd.DataFrame(StaticLineForces,index=lines,columns=DataNames)


# Fetching Vessel Excustions
VesselExcursions = list()

ves=model_0[vesName]

temp=ves.StaticResult('X')
VesselExcursions.append(temp)

temp=ves.StaticResult('Y')
VesselExcursions.append(temp)

temp=ves.StaticResult('Z')
VesselExcursions.append(temp)

temp=ves.StaticResult('Rotation 1')
VesselExcursions.append(temp)

temp=ves.StaticResult('Rotation 2')
VesselExcursions.append(temp)

temp=ves.StaticResult('Rotation 3')
VesselExcursions.append(temp)


DF_ISR_EXC = pd.DataFrame(VesselExcursions,index=['X','Y','Z','Roll','Pitch','Yaw'],columns=[vesName])

with pd.ExcelWriter('output.xlsx',mode='w') as writer:  
    DF_ISR_TEN.to_excel(writer,sheet_name='Intact-Static Tensions')
    DF_ISR_EXC.to_excel(writer,sheet_name='Intact-Static Excursions')