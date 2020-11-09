# -*- coding: utf-8 -*-
"""
Created on Mon Oct 19 14:10:05 2020

@author: Marc Bodem

Template for Calling CST from Python
- Lossy Load Waveguide (Microwave Studio)
- Store Parameter to load one sample point at a time
- change one parameter at a time
- Obtain S-Parameter values

"""
import sys
sys.path.append(r"C:\Program Files (x86)\CST Studio Suite 2020\AMD64\python_cst_libraries")
import cst
import cst.interface
import cst.results
import numpy as np
import time
import shutil

start = time.time()

# Local path to CST project file --> Please adapt
cst_path = r'D:\Documents\Vorlagen\CST_Python_Interface' # path
cst_project = '\Lossy Loaded Waveguide' # CST project

cst_project_path = cst_path + cst_project + '.cst'
cst_result_folder = cst_path + cst_project + '\Result'


# Delete all old results (if exist)
# This option is only necessary in the Parameter Sweep Option - here it could only be a nice feature in order to save computing time
try:
    shutil.rmtree(cst_result_folder)
except OSError as e:
    print(e)
else:
    print("The results directory is deleted successfully")


# Define parameter value for change
sample_point = 0.06


# Define list with frequency points of interest, here 80-120 GHz, Number of calculated points are set to 1001 in CST file, to get the desired frequencies.
freq_range = [80,90,100,110,120]
# This line mights be adapted, depending on the number of frequency result points and the specified frequency range
freq_range_pos_float = 0.1*(np.array(freq_range)-freq_range[0])*250
freq_range_pos= freq_range_pos_float.astype(int)


# Initialize lists with S-parameter results
S_all = []
S_real_all = []
S_imag_all = []
SdB_all = []
all_S_Para=[]


# Open CST as software
mycst=cst.interface.DesignEnvironment()

# Open CST Project
mycst1=cst.interface.DesignEnvironment.open_project(mycst,cst_project_path)


#Delete All Sequences before starting to get a fresh file
dele = 'Sub Main () \n dim objName as object \n set objName = Result1D("S-Parameters") \n DeleteAt("truemodelchange") \nEnd Sub ()'
mycst1.schematic.execute_vba_code(dele, timeout=None)
    
#VBA Code for Parameter change and rebuild 
par_change = 'Sub Main () \n StoreParameter("width", '+str(sample_point)+') \nRebuildOnParametricChange (bfullRebuild, bShowErrorMsgBox)\nEnd Sub' 
   
#execute VBA Code above
mycst1.schematic.execute_vba_code(par_change, timeout=None)
   
#start solver
mycst1.modeler.run_solver()

#get project for results
project = cst.results.ProjectFile(cst_project_path, allow_interactive=True)

#get the results
results = project.get_3d().get_result_item(r"1D Results\S-Parameters\S1,1")

#get frequencies
freqs = np.array(results.get_xdata())

#get S21 Parameter
S_Para = np.array(results.get_ydata())

# Initialize value list for one MC sample point over all frequency points
freq_pos = []
freq = []
S = []
SdB = []
S_real = []
S_imag = []
    
    
# Get results for each freq. point of interest from CST
for j in range(len(freq_range)):
    freq_pos_j = freq_range_pos[j]
    freq_pos.append(freq_pos_j)
        
    freq_value_j = freqs[freq_pos_j]
    freq.append(freq_value_j)
        
    S_real_j = S_Para[freq_pos_j].real
    S_real.append(S_real_j)
        
    S_imag_j = S_Para[freq_pos_j].imag
    S_imag.append(S_imag_j)
        
    S_j=np.sqrt(S_real_j**2+S_imag_j**2)
    S.append(S_j)
        
    S_dB_j = 20*np.log10(S_j)
    SdB.append(S_dB_j)


#close CST
cst.interface.DesignEnvironment.close(mycst) 

# Return results 
print('\n frequency range in GHz:', freq_range)
print('frequency range as positions in CST:', freq_pos)
print('S-parameter:',S)
print('S-parameter (real part):',S_real)
print('S-parameter (imag part):',S_imag)
print('S-parameter (in dB):',SdB)

# Return runtime
end = time.time()
print('\n Runtime: {:5.3f}seconds'.format(end-start))
