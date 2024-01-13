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
Intact Dynamic Results
--------------------------------------------------------------------------'''

LineParmList = ['Effective Tension','End GX force','End GY force','End GZ force','Effective Tension','End GX force','End GY force','End GZ force']
LineOEList = [OrcFxAPI.oeEndA,OrcFxAPI.oeEndA,OrcFxAPI.oeEndA,OrcFxAPI.oeEndA,OrcFxAPI.oeEndB,OrcFxAPI.oeEndB,OrcFxAPI.oeEndB,OrcFxAPI.oeEndB]
LineSheetNames = ['End A EFF TEN ','End A GX F','End A GY F','End A GZ F','End B EFF TEN ','End B GX F','End B GY F','End B GZ F']

VesParmList = ['X','Y','Z','Rotation 1','Rotation 2','Rotation 3']
VOE = OrcFxAPI.oeVessel((0,0,0))
VesSheetNames = 'Vessel Excursions'

nLineParms = len(LineParmList)
nVesParms = len(VesParmList)

MPV_MAX_DATA_LINE = np.zeros([nICM,nLines,nLineParms])
MPV_MIN_DATA_LINE = np.zeros([nICM,nLines,nLineParms])
MAX_DATA_LINE = np.zeros([nICM,nLines,nLineParms])
MIN_DATA_LINE = np.zeros([nICM,nLines,nLineParms])
RMS_DATA_LINE = np.zeros([nICM,nLines,nLineParms])

MPV_MAX_DATA_VES = np.zeros([nICM,nVesParms])
MPV_MIN_DATA_VES = np.zeros([nICM,nVesParms])
MAX_DATA_VES = np.zeros([nICM,nVesParms])
MIN_DATA_VES = np.zeros([nICM,nVesParms])
RMS_DATA_VES = np.zeros([nICM,nVesParms])

    
for i in range(nICM):
    
    # Save the File
    fileName = os.path.join(INTACT_DIR, BASENAME +
                            '_INTACT_DYNAMICS_'+str(DF_ICM.CASE_ID[i]).replace(' ', '_')+'.sim')

    model_0 = OrcFxAPI.Model(fileName)
    
    ''' ------------------------------------------------------------------------
    Line Forces / Tensions
    --------------------------------------------------------------------------'''
    
    for j in range(nLines):
        
        obj=model_0[lines[j]]
        
        for k in range(nLineParms):
            
            # Most Probable
            extrmStats=obj.ExtremeStatistics(LineParmList[k],OrcFxAPI.Period(OrcFxAPI.PeriodNum.WholeSimulation),LineOEList[k])
            
            # # Rayleigh Distribution
            # # # Maximum
            esSpec=OrcFxAPI.RayleighStatisticsSpecification(ExtremesToAnalyse=0)
            extrmStats.Fit(esSpec)
            query = OrcFxAPI.RayleighStatisticsQuery(StormDurationHours=3, RiskFactor=1)
            extrms=extrmStats.Query(query)
            MPV_MAX_DATA_LINE[i,j,k]=extrms.MostProbableExtremeValue

            # # # Maximum
            esSpec=OrcFxAPI.RayleighStatisticsSpecification(ExtremesToAnalyse=1)
            extrmStats.Fit(esSpec)
            query = OrcFxAPI.RayleighStatisticsQuery(StormDurationHours=3, RiskFactor=1)
            extrms=extrmStats.Query(query)
            MPV_MIN_DATA_LINE[i,j,k]=extrms.MostProbableExtremeValue
            
            # Max and Min
            stats=obj.AnalyseExtrema(LineParmList[k],OrcFxAPI.Period(OrcFxAPI.PeriodNum.WholeSimulation),LineOEList[k])
            MAX_DATA_LINE[i,j,k]=stats.Max
            MIN_DATA_LINE[i,j,k]=stats.Min
            
            # RMS Value
            stats=obj.TimeSeriesStatistics(LineParmList[k],OrcFxAPI.Period(OrcFxAPI.PeriodNum.WholeSimulation),LineOEList[k])
            RMS_DATA_LINE[i,j,k]=stats.RMS

    ''' ------------------------------------------------------------------------
    Vessel Excursions
    --------------------------------------------------------------------------'''      
    obj = model_0[vesName]
    
    
    for k in range(nVesParms):
        
        extrmStats=obj.ExtremeStatistics(VesParmList[k],OrcFxAPI.Period(OrcFxAPI.PeriodNum.WholeSimulation),VOE)
        
        # # Rayleigh Distribution
        # # # Maximum
        esSpec=OrcFxAPI.RayleighStatisticsSpecification(ExtremesToAnalyse=0)
        extrmStats.Fit(esSpec)
        query = OrcFxAPI.RayleighStatisticsQuery(StormDurationHours=3, RiskFactor=1)
        extrms=extrmStats.Query(query)
        MPV_MAX_DATA_VES[i,k]=extrms.MostProbableExtremeValue
    
        # # # Maximum
        esSpec=OrcFxAPI.RayleighStatisticsSpecification(ExtremesToAnalyse=1)
        extrmStats.Fit(esSpec)
        query = OrcFxAPI.RayleighStatisticsQuery(StormDurationHours=3, RiskFactor=1)
        extrms=extrmStats.Query(query)
        MPV_MIN_DATA_VES[i,k]=extrms.MostProbableExtremeValue
        
        # Max and Min
        stats=obj.AnalyseExtrema(VesParmList[k],OrcFxAPI.Period(OrcFxAPI.PeriodNum.WholeSimulation),VOE)
        MAX_DATA_VES[i,k]=stats.Max
        MIN_DATA_VES[i,k]=stats.Min
        
        # RMS Value
        stats=obj.TimeSeriesStatistics(VesParmList[k],OrcFxAPI.Period(OrcFxAPI.PeriodNum.WholeSimulation),VOE)
        RMS_DATA_VES[i,k]=stats.RMS
    

with pd.ExcelWriter('output.xlsx',mode='a',if_sheet_exists='replace') as writer:  
    
    for i in range(nLineParms):
        DF = pd.DataFrame(MPV_MAX_DATA_LINE[:,:,0],index=DF_ICM.CASE_ID,columns=lines)
        DF.to_excel(writer,'MPV_MAX_'+LineSheetNames[i])
        DF = pd.DataFrame(MPV_MIN_DATA_LINE[:,:,0],index=DF_ICM.CASE_ID,columns=lines)
        DF.to_excel(writer,'MPV_MIN_'+LineSheetNames[i])
        DF = pd.DataFrame(MAX_DATA_LINE[:,:,0],index=DF_ICM.CASE_ID,columns=lines)
        DF.to_excel(writer,'MAX_'+LineSheetNames[i])
        DF = pd.DataFrame(MIN_DATA_LINE[:,:,0],index=DF_ICM.CASE_ID,columns=lines)
        DF.to_excel(writer,'MIN_'+LineSheetNames[i])        
        DF = pd.DataFrame(RMS_DATA_LINE[:,:,0],index=DF_ICM.CASE_ID,columns=lines)
        DF.to_excel(writer,'RMS_'+LineSheetNames[i])           

with pd.ExcelWriter('output.xlsx',mode='a',if_sheet_exists='replace') as writer:  
    
    for i in range(nVesParms):
        DF = pd.DataFrame(MPV_MAX_DATA_VES[:,:],index=DF_ICM.CASE_ID,columns=VesParmList)
        DF.to_excel(writer,'MPV_MAX_'+VesSheetNames)
        DF = pd.DataFrame(MPV_MIN_DATA_VES[:,:],index=DF_ICM.CASE_ID,columns=VesParmList)
        DF.to_excel(writer,'MPV_MIN_'+VesSheetNames)
        DF = pd.DataFrame(MAX_DATA_VES[:,:],index=DF_ICM.CASE_ID,columns=VesParmList)
        DF.to_excel(writer,'MAX_'+VesSheetNames)
        DF = pd.DataFrame(MIN_DATA_VES[:,:],index=DF_ICM.CASE_ID,columns=VesParmList)
        DF.to_excel(writer,'MIN_'+VesSheetNames)        
        DF = pd.DataFrame(RMS_DATA_VES[:,:],index=DF_ICM.CASE_ID,columns=VesParmList)
        DF.to_excel(writer,'RMS_'+VesSheetNames)
        