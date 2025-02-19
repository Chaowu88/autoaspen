r'''
Generates a training dataset using the Aspen model and the .xslm calculator.

If an error occurs in Excel, begin by deleting the unsuccessful .bkp file and then attempt to 
run the script again. If the issue persists, you can use the "delete_samples.py" script to 
clear the input values in the "Values" column and then make another attempt.

Usage:
python path\to\generate_dataset.py
'''


__author__ = 'Chao Wu'


import os
import re
from collections import namedtuple
import numpy as np
import pandas as pd
from pythoncom import CoInitialize
from win32com.client import DispatchEx
from time import sleep


DATASET_FILE = 'path\to\training_data.xlsx'
ASPEN_FILE = 'path\to\aspen_model.bkp'
CALCULATOR_FILE = r'path\to\calculator.xlsm'


class Excel():

    def __init__(self, excelFile):
        '''
        Parameters
        ----------
        excelFile: str
            Path to the Excel file.
        '''
        
        CoInitialize()
        self.excelCOM = DispatchEx('Excel.Application')
        self.excelBook = self.excelCOM.Workbooks.Open(excelFile)


    def get_cell(self, sheet, loc = None, row = None, col= None):
        '''
        Parameters
        ----------
        sheet: str
            Name of the sheet containing the cell.
        loc: str
            Location of the cell in Excel's A1 notation (e.g., "A1", "B2", "C3", etc.).
            A cell location can be specified either by the loc argument or the row and col 
            arguments.
        row: int or str
            Index of the row containing the cell.
        col: str
            Index of the column containing the cell.
        
        Returns
        -------
        cellValue: float or str
            The value of the cell after calculation.    
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
        value: float or str
            The value to write to the cell.
        sheet: str
            Name of the sheet containing the cell.
        loc: str
            Location of the cell in Excel's A1 notation (e.g., "A1", "B2", "C3", etc.).
            A cell location can be specified either by the loc argument or the row and col 
            arguments.
        row: int or str
            Index of the row containing the cell.
        col: str
            Index of the column containing the cell.
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
            Path to the Aspen Plus file.
        '''
        
        self.set_cell(aspenFile, 'Set-up', 'B1')
    
        self.run_macro('sub_ClearSumData_ASPEN')
        self.run_macro('sub_GetSumData_ASPEN')
        
        
    def run_macro(self, macro):
        '''
        Parameters
        ----------
        macro: str
            The name of the Excel macro to run.
        '''

        sleep(3)
        
        self.excelCOM.Run(macro)    
        
            
    def close(self):
        
        self.excelBook.Close(SaveChanges = 0)
        self.excelCOM.Application.Quit()
        
        
class Aspen():

    def __init__(self, aspenFile):
        '''
        Parameters
        ----------
        aspenFile: str
            Path to the Aspen Plus file.
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
            Path to the node in the Aspen Plus tree.
        
        Returns
        -------
        value: float or str
            The value of the node.
        '''
        
        value = self.COM.Tree.FindNode(aspenPath).Value
        
        return value
        
        
    def set_value(self, aspenPath, value, ifFortran):
        '''
        Parameters
        ----------
        aspenPath: str
            Path to the node in the Aspen Plus tree.
        value: float or str
            The value to write to the node.
        ifFortran: bool
            Whether the node is a Fortran variable
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
            The file name to save the Aspen Plus file (.bkp) to. 
        '''
    
        self.COM.SaveAs(saveFile)
        
    
    def close(self):
        
        self.COM.Close()
    
        
def parse_data_file(data_file):
    '''
    Parameters
    ----------
    data_file: str
        Path to the data file.
    
    Returns
    -------
    inputInfos: pd.DataFrame
        A DataFrame containing the input information.
    outputInfo: pd.DataFrame
        A DataFrame containing the output information.
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
        Path to the data file.
    input_infos: pd.DataFrame
        A DataFrame containing the input information with columns 'Input variable', 'Type', 
        'Location', and 'Values'.
    output_info: pd.DataFrame
        A DataFrame containing the output information with columns 'Output variable', 
        'Location', and 'Values'.
    aspen_file: str
        Path to the Aspen model file.
    calculator_file: str
        Path to the Excel calculator file.
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
        print(f'{nruns} runs in total, {nrunsLeft} runs left')

        outDir = os.path.dirname(data_file)
        tmpDir = outDir + '/tmp'
        os.makedirs(tmpDir, exist_ok = True)
    
        aspenModel = Aspen(aspen_file)
        calculator = Excel(calculator_file)

        for i in range(nrunsCompl, nruns):
            print(f'run {i+1}:')
            
            for inputInfo in inputInfos:
                if inputInfo.type == 'bkp':
                    aspenModel.set_value(inputInfo.loc, inputInfo.values[i], False)   
                elif inputInfo.type == 'bkp_fortran':
                    aspenModel.set_value(inputInfo.loc, inputInfo.values[i], True)
                else:
                    continue

            aspenModel.run_model()
            
            tmpFile = f'{tmpDir}/{i}.bkp'
            aspenModel.save_model(tmpFile)

            for inputInfo in inputInfos:
                if inputInfo.type == 'xlsm':
                    inputSheet, inputCell = inputInfo.loc.split('!')
                    calculator.set_cell(inputInfo.values[i], inputSheet, loc = inputCell)
                else:
                    continue
            
            calculator.load_aspenModel(tmpFile)
            calculator.run_macro('solvedcfror')
            
            outputSheet, outputCell = outputInfo.loc.split('!')
            output = calculator.get_cell(outputSheet, loc = outputCell)
            outputInfo.values.append(output)
            
            outputValues = ','.join(map(str, outputInfo.values))
            output_info = pd.DataFrame(
                [[outputInfo.name, outputInfo.loc, outputValues]],
                columns = ['Output variable', 'Location', 'Values']
            )
            
            with pd.ExcelWriter(data_file) as writer:
                input_infos.to_excel(writer, sheet_name = 'Inputs', index = False)
                output_info.to_excel(writer, sheet_name = 'Output', index = False)
            
            print('done')
            
        aspenModel.close()
        calculator.close()
        
    else:
        print('all runs completed')
    
    
    
    
if __name__ == '__main__':
    
    inputsInfo, outputInfo = parse_data_file(DATASET_FILE)
    run_and_update(DATASET_FILE, inputsInfo, outputInfo, ASPEN_FILE, CALCULATOR_FILE)
