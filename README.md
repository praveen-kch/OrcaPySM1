# OrcaPySM1

Python Scripts to Automate OrcaFlex Modelling : Spread Moored Vessels

Description : 

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


