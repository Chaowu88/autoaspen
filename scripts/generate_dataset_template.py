#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '05/21/2021'


r'''
This script generates dataset template for training. Random values of the input variables are generated according to defined distributions in format of:
    Distribution   Parameters
    normal         mean,sd
    alpha          a,loc,scale
    Beta           a,b,loc,scale
    gamma          a,loc,scale            
    triangular     c,loc,scale
    pareto         b,loc,scale
    bernoulli      pl,ph          # prob. of low value, prob. of high value
  
Input variable types are:
    xlsm for calculator variables
    bkp for Aspen non-Fortran variables
    bkp_fortran for Aspen Fortran variables

Usage
python path\to\generate_dataset_template.py
'''


OUTPUT_FILE = 'path\to\training\training_data.xlsx'
CONFIG_FILE = 'path\to\var_infos.xlsx'
NRUNS = 100


import os
import numpy as np
import pandas as pd
from scipy import stats


def parse_config_file(config_file):
    '''
    Parameters
    ----------
    config_file: str
        path of config file
    
    Returns
    -------
    inputsInfo, outputInfo: df
    '''
    
    configInfo = pd.read_excel(config_file, sheet_name = ['Inputs', 'Output'])
    inputsInfo = configInfo['Inputs']
    outputInfo = configInfo['Output']
    
    return inputsInfo, outputInfo


def generate_input_values(inputs_info, nruns):
    '''
    Parameters
    ----------
    inputs_info: df
        columns are ['Input variable', 'Type', 'Location', 'Bounds', 'Distribution', 'Parameters']
    nruns: int
        # of runs
    
    Returns
    -------
    inputsValues: df
    '''
    
    inputsValues = pd.DataFrame(columns = ['Input variable', 'Type', 'Location', 'Values'])
    for _, [inputVar, varType, local, bnds, distName, params] in inputs_info.iterrows():
        
        print('generating random values of %s' % inputVar)
        
        lb, ub = map(float, bnds.split(','))
        dist = getattr(stats, distName)
        
        if lb >= ub:
            raise ValueError('lower bound should be less than upper bound')

        if distName == 'uniform':
            values = dist.rvs(loc = lb, scale = ub-lb, size = nruns)

        elif distName == 'bernoulli':
            pl, ph = map(float, params.split(','))
            labels = dist.rvs(pl, size = nruns)
            values = [lb if label else ub for label in labels]
        
        else:
            *shapeParams, loc, scale = map(float, params.split(','))
            
            values = []
            count = 0
            while count < nruns:
                value = dist.rvs(*shapeParams, loc = loc, scale = scale)
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
        output file
    inputs_values: df
        columns are ['Input variable', 'Type', 'Location', 'Values']
    output_info: df 
        columns are ['Output variable', 'Location']
    '''
    
    outDir = os.path.dirname(out_file)
    os.makedirs(outDir, exist_ok = True)
    
    output_info = output_info.copy()
    output_info['Values'] = 'NaN'
    
    with pd.ExcelWriter(out_file) as writer:
        inputs_values.to_excel(writer, sheet_name = 'Inputs', index = False)
        output_info.to_excel(writer, sheet_name = 'Output', index = False)
    
    


if __name__ == '__main__':
    
    inputsInfo, outputInfo = parse_config_file(CONFIG_FILE)
    inputsValues = generate_input_values(inputsInfo, NRUNS)
    write_to_excel(OUTPUT_FILE, inputsValues, outputInfo)
    
