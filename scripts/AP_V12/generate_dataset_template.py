r'''
Creates a dataset template for training by generating random values for input 
variables according to specified distributions with defined parameters.
    
Distribution Parameters:
normal: mean, standard deviation (mean, sd)
alpha: shape, location, scale (a, loc, scale)
beta: shape, shape, location, scale (a, b, loc, scale)
gamma: shape, location, scale (a, loc, scale)
triang: shape, location, scale (c, loc, scale)
pareto: shape, location, scale (b, loc, scale)
bernoulli: probability of low value, probability of high value (pl, ph)
  
Input Variable Types:
xlsm: Calculator variables
bkp: Aspen non-Fortran variables
bkp_fortran: Aspen Fortran variables

Usage:
python path\to\generate_dataset_template.py
'''


__author__ = 'Chao Wu'


import os
import numpy as np
import pandas as pd
from scipy import stats


OUTPUT_FILE = 'path\to\training_data.xlsx'
CONFIG_FILE = 'path\to\var_info.xlsx'
NRUNS = 100


def parse_config_file(config_file):
    '''
    Parameters
    ----------
    config_file: str
        Path to the configuration file.
    
    Returns
    -------
    inputsInfo: pd.DataFrame
        A DataFrame containing the input information.
    outputInfo: pd.DataFrame
        A DataFrame containing the output information.
    '''
    
    configInfo = pd.read_excel(config_file, sheet_name=['Inputs', 'Output'])
    inputsInfo = configInfo['Inputs']
    outputInfo = configInfo['Output']
    
    return inputsInfo, outputInfo


def generate_input_values(inputs_info, nruns):
    '''
    Parameters
    ----------
    inputs_info: pd.DataFrame
        A DataFrame containing information about the input variables with columns 
        'Input variable', 'Type', 'Location', 'Bounds', 'Distribution' and 
        'Parameters'.
    nruns: int
        # of runs to generate.
    
    Returns
    -------
    inputsValues: pd.DataFrame
        A DataFrame containing the generated input values.
    '''
    
    inputsValues = pd.DataFrame(
        columns=['Input variable', 'Type', 'Location', 'Values']
    )
    for _, cols in inputs_info.iterrows():
        (inputVar, varType, local, bnds, distName, params) = cols
        print(f'generating random values of {inputVar}')
        
        lb, ub = map(float, bnds.split(','))
        dist = getattr(stats, distName)
        
        if lb >= ub:
            raise ValueError('lower bound should be less than upper bound')

        if distName == 'uniform':
            values = dist.rvs(loc=lb, scale=ub-lb, size=nruns)

        elif distName == 'bernoulli':
            pl, ph = map(float, params.split(','))
            labels = dist.rvs(pl, size=nruns)
            values = [lb if label else ub for label in labels]
        
        else:
            *shapeParams, loc, scale = map(float, params.split(','))
            
            values = []
            count = 0
            while count < nruns:
                value = dist.rvs(*shapeParams, loc=loc, scale=scale)
                if lb <= value <= ub:
                    count += 1
                    values.append(value)
        
        values = ','.join(np.array(values).astype(str))
        
        inputsValues.loc[inputVar, :] = [inputVar, varType, local, values]
    
    return inputsValues


def write_to_excel(out_file, inputs_values, output_info):
    '''
    Parameters
    ----------
    out_file: str
        Path to the output file.
    inputs_values: pd.DataFrame
        A DataFrame containing information about the input values with columns 
        'Input variable', 'Type', 'Location', and 'Values'.
    output_info: pd.DataFrame 
        A DataFrame containing information about the output variables with columns 
        'Output variable', and 'Location'.
    '''
    
    outDir = os.path.dirname(out_file)
    os.makedirs(outDir, exist_ok = True)
    
    output_info = output_info.copy()
    output_info['Values'] = 'NaN'
    
    with pd.ExcelWriter(out_file) as writer:
        inputs_values.to_excel(writer, sheet_name='Inputs', index=False)
        output_info.to_excel(writer, sheet_name='Output', index=False)
    
    


if __name__ == '__main__':
    
    inputsInfo, outputInfo = parse_config_file(CONFIG_FILE)
    inputsValues = generate_input_values(inputsInfo, NRUNS)
    write_to_excel(OUTPUT_FILE, inputsValues, outputInfo)
    