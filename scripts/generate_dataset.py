#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '08/24/2022'


r'''
This script generates training dataset using Aspen model and .xslm calculator.
If excel raise an error, try run the script again first. If no success, delete the inputs of the run in Values column,
and try again.

Usage
python path\to\generate_dataset.py
'''


DATASET_FILE = 'path\to\training\training_data.xlsx'
ASPEN_FILE = 'path\to\aspenmodel.bkp'
CALCULATOR_FILE = 'path\to\excel_calculator.xlsm'


import os
import re
from collections import namedtuple
import numpy as np
import pandas as pd
from pythoncom import CoInitialize
from win32com.client import DispatchEx


class Excel():
    '''
    Parameters
    ----------
    file: str
        excel file
    '''

    def __init__(self, excelFile):
        '''
        Parameters
        ----------
        file: str
            excel file
        '''
        
        CoInitialize()
        self.excelCOM = DispatchEx('Excel.Application')
        self.excelBook = self.excelCOM.Workbooks.Open(excelFile)


    def get_cell(self, sheet, loc = None, row = None, col= None):
        '''
        Parameters
        ----------
        sheet: str
            sheet name
        loc: cell location
            equivalent to col+row
        row: int or str
            row index
        col: str
            column index
        
        Returns
        -------
        cellValue: num or str
            cell value (after calculation)
        '''
        
        sht = self.excelBook.Worksheets(sheet)
        
        if loc != None:
            cellValue = sht.Evaluate(loc).Value
        
        elif row != None and col != None:
            cellValue = sht.Cells(row, col).Value
        
        return cellValue

    
    def set_cell(self, value, sheet, loc = None, row = None, col= None):
        '''
        Parameters
        ----------
        value: num or str
            value to set
        sheet: str
            sheet name
        loc: cell location
            equivalent to col+row
        row: int or str
            row index
        col: str
            column index
        '''
        
        sht = self.excelBook.Worksheets(sheet)
        
        if loc != None:
            sht.Evaluate(loc).Value = value
        
        elif row != None and col != None:
            sht.Cells(row, col).Value = value
    
    
    def load_aspenModel(self, aspenFile):
        '''
        Parameters
        ----------
        aspenFile: str
            aspen file
        '''
        
        self.set_cell(aspenFile, 'Set-up', 'B1')
    
        self.run_macro('sub_ClearSumData_ASPEN')
        self.run_macro('sub_GetSumData_ASPEN')
        
        
    def run_macro(self, macro):
        '''
        Parameters
        ----------
        macro: str
            macro
        '''
            
        self.excelCOM.Run(macro)    
        
            
    def close(self):
        
        self.excelBook.Close(SaveChanges = 0)    
        
        
class Aspen():
    '''
    Parameters
    ----------
    file: str
        excel file
        '''

    def __init__(self, aspenFile):
        '''
        Parameters
        ----------
        file: str
            excel file
        '''
        
        CoInitialize()
        self.file = aspenFile
        self.COM = DispatchEx('Apwn.Document')
        self.COM.InitFromArchive2(self.file)
        
        
    def get_value(self, aspenPath):
        '''
        Parameters
        ----------
        aspenPath: str
            path in ASPEN tree
        
        Returns
        -------
        value: num or str
            value in ASPEN tree node
        '''
        
        value = self.COM.Tree.FindNode(aspenPath).Value
        
        return value
        
        
    def set_value(self, aspenPath, value, ifFortran):
        '''
        Parameters
        ----------
        aspenPath: str
            path in ASPEN tree
        value: float or str
            value to set
        ifFortran: bool
            whether it is a Fortran variable
        '''
        
        if ifFortran:
            oldValue = self.COM.Tree.FindNode(aspenPath).Value
            self.COM.Tree.FindNode(aspenPath).Value = re.sub(r'(?<== ).+', str(value), oldValue)
        else:
            self.COM.Tree.FindNode(aspenPath).Value = float(value)
        

    def run_model(self):
    
        self.COM.Reinit()
        self.COM.Engine.Run2()
        
        
    def save_model(self, saveFile):
        '''
        Parameters
        ----------
        saveFile: str
            file name to save (.bkp)
        '''
    
        self.COM.SaveAs(saveFile)
        
    
    def close(self):
        
        self.COM.Close()
    
        
def parse_data_file(data_file):
    '''
    Parameters
    ----------
    data_file: str
        path of data file
    
    Returns
    -------
    inputInfos: df
    outputInfo: df
    '''
    
    dataInfo = pd.read_excel(data_file, sheet_name = ['Inputs', 'Output'])
    inputInfo = dataInfo['Inputs']
    outputInfo = dataInfo['Output']
    
    return inputInfo, outputInfo
    
    
def run_and_update(data_file, input_infos, output_info, aspen_file, calculator_file):    
    '''
    Parameters
    ----------
    data_file: str
        dataset file
    input_infos: df
        columns are ['Input variable', 'Type', 'Location', 'Values']
    output_info: df
        columns are ['Output variable', 'Location', 'Values']
    aspen_file: str
        Aspen model file
    calculator_file: str
        .xslm calculator file
    '''
    
    *others, values = output_info.squeeze()
    if isinstance(values, str):
        values = list(map(float, values.split(',')))
    elif np.isnan(values):
        values = []
    else:
        raise TypeError("what's in the Values column of Output sheet?")
        
    OutputInfo = namedtuple('OutputInfo', ['name', 'loc', 'values'])
    outputInfo = OutputInfo(*others, values)
    nrunsCompl = len(outputInfo.values)

    InputInfo = namedtuple('InputInfo', ['name', 'type', 'loc', 'values'])
    inputInfos = []
    for _, [*others, values] in input_infos.iterrows():
        
        values = list(map(float, values.split(',')))
        inputInfos.append(InputInfo(*others, values))
    nruns = len(values)
    
    nrunsLeft = nruns - nrunsCompl
    if nrunsLeft > 0:
        print('%s runs in total, %s runs left' % (nruns, nrunsLeft))

        # run
        outDir = os.path.dirname(data_file)
        tmpDir = outDir + '/tmp'
        os.makedirs(tmpDir, exist_ok = True)
    
        aspenModel = Aspen(aspen_file)
        calculator = Excel(calculator_file)

        for i in range(nrunsCompl, nruns):
            print('run %s:' % (i+1))
            
            # set Aspen variables
            for inputInfo in inputInfos:
                if inputInfo.type == 'bkp':
                    aspenModel.set_value(inputInfo.loc, inputInfo.values[i], False)   
                elif inputInfo.type == 'bkp_fortran':
                    aspenModel.set_value(inputInfo.loc, inputInfo.values[i], True)
                else:
                    continue

            # run Aspen model
            aspenModel.run_model()
            
            tmpFile = '%s/%s.bkp' % (tmpDir, i)
            aspenModel.save_model(tmpFile)

            # set calculator variables
            for inputInfo in inputInfos:
                if inputInfo.type == 'xlsm':
                    inputSheet, inputCell = inputInfo.loc.split('!')
                    calculator.set_cell(inputInfo.values[i], inputSheet, loc = inputCell)
                else:
                    continue
            
            # run calculator
            calculator.load_aspenModel(tmpFile)
            calculator.run_macro('solvedcfror')
            
            outputSheet, outputCell = outputInfo.loc.split('!')
            output = calculator.get_cell(outputSheet, loc = outputCell)
            outputInfo.values.append(output)
            
            # update dataset
            outputValues = ','.join(map(str, outputInfo.values))
            output_info = pd.DataFrame([[outputInfo.name, outputInfo.loc, outputValues]],
                                       columns = ['Output variable', 'Location', 'Values'])
            
            with pd.ExcelWriter(data_file) as writer:
                input_infos.to_excel(writer, sheet_name = 'Inputs', index = False)
                output_info.to_excel(writer, sheet_name = 'Output', index = False)
                writer.save()
            
            print('done')
            
        aspenModel.close()
        calculator.close()
        
    else:
        print('all runs completed')
    
    
    
    
if __name__ == '__main__':
    
    inputsInfo, outputInfo = parse_data_file(DATASET_FILE)
    run_and_update(DATASET_FILE, inputsInfo, outputInfo, ASPEN_FILE, CALCULATOR_FILE)
