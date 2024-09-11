r'''
Generates a training dataset using the Aspen model and the .xslm calculator.

If an error occurs in Excel, begin by deleting the unsuccessful .bkp file and then 
attempt to run the script again. If the issue persists, you can use the 
"delete_samples.py" script to clear the input values in the "Values" column and then 
make another attempt.

The zetoolkit.dll from Aspen Plus V11 and later versions only supports 64-bit Excel. For 32-bit Excel, zetoolkit.dll from older versions of Aspen Plus should be provided.
If ZETOOLKIT_FILE is not specified, the default path 
"C:\Program Files\AspenTech\AprSystem Vxxx\Engine\Xeq\zetoolkit.dll" will be used.

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
from winreg import (CreateKey, EnumKey, EnumValue, OpenKey, QueryValueEx, CloseKey, 
                    HKEY_CLASSES_ROOT, HKEY_LOCAL_MACHINE)


DATASET_FILE = 'path\to\training_data.xlsx'
ASPEN_FILE = 'path\to\aspen_model.bkp'
CALCULATOR_FILE = 'path\to\calculator.xlsm'
ZETOOLKIT_FILE = 'path\to\zetoolkit.dll'


class Excel():

    def __init__(self, excelFile, aspenVersion=None, zetoolkitFile=None):
        '''
        Parameters
        ----------
        excelFile: str
            Path to the Excel file.
        aspenVersion: str
            Aspen Plus version.
        zetoolkitFile: str
            Path to the zetoolkit.dll.
        '''
        
        CoInitialize()
        self.excelCOM = DispatchEx('Excel.Application')
        self.excelBook = self.excelCOM.Workbooks.Open(excelFile)
        if zetoolkitFile is None:
            if aspenVersion is None:
                raise ValueError('zetoolkit.dll should be provided')
            else:
                zetoolkitFile = (r'C:\Program Files\AspenTech\AprSystem '
                                 rf'{aspenVersion}\Engine\Xeq\zetoolkit.dll')
        self._detect_bitness_compatibility(zetoolkitFile)
        self.set_cell(zetoolkitFile, 'Set-up', 'TKDLL')


    def _detect_bitness_compatibility(self, zetoolkitFile):
        '''
        Parameters
        ----------
        zetoolkitFile: str
            Path to the zetoolkit.dll.
        '''

        officeKeyPath = r'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        with OpenKey(HKEY_LOCAL_MACHINE, officeKeyPath) as officeKey:
            officeBit = QueryValueEx(officeKey, 'Platform')[0]
            if officeBit == 'x86':
                officeBit = 32
            elif officeBit == 'x64':
                officeBit = 64

        tkdllBit = None
        with open(zetoolkitFile, 'br') as f:
            for line in f:
                if b'PE\x00\x00L' in line:
                    tkdllBit = 32
                    break
                elif b'PE\x00\x00d' in line:
                    tkdllBit = 64
                    break

        if officeBit != tkdllBit:
            raise ValueError('Excel bitness does not match zetoolkit.dll')


    def get_cell(self, sheet, loc=None, row=None, col=None):
        '''
        Parameters
        ----------
        sheet: str
            Name of the sheet containing the cell.
        loc: str
            Location of the cell in Excel's A1 notation (e.g., "A1", "B2", "C3", 
            etc.). A cell location can be specified either by the loc argument or 
            the row and col arguments.
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

    
    def set_cell(self, value, sheet, loc=None, row=None, col=None):
        '''
        Parameters
        ----------
        value: float or str
            The value to write to the cell.
        sheet: str
            Name of the sheet containing the cell.
        loc: str
            Location of the cell in Excel's A1 notation (e.g., "A1", "BKPNAME", 
            etc.). A cell location can be specified either by the loc argument or 
            the row and col arguments.
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
        
        self.set_cell(aspenFile, 'Set-up', loc='BKPNAME')

        self.run_macro('sub_GetSumData_ASPEN')
        
        
    def run_macro(self, macro):
        '''
        Parameters
        ----------
        macro: str
            The name of the Excel macro to run.
        '''
            
        self.excelCOM.Application.Run(f'{self.excelBook.Name}!{macro}')


    def calculate_dcfror(self):
        self.run_macro('sub_SolveProductCost_DCFROR')
        
            
    def close(self):
        self.excelBook.Close(SaveChanges=False)
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
        self.version = self._get_aspen_version()
        if self.version is not None:
            self.aspenCOM = DispatchEx('Apwn.Document')
            self.aspenCOM.InitFromArchive2(aspenFile)


    @staticmethod
    def _get_aspen_version():
        keys = CreateKey(HKEY_CLASSES_ROOT, '')
        i = 0
        count = 0
        while True:
            try:
                key = EnumKey(keys, i)
                if key == 'Apwn.Document':
                    aspenKey = CreateKey(keys, key)
                    iconKey = CreateKey(aspenKey, 'DefaultIcon')
                    version = re.search(
                        r'V\d+.\d+', 
                        EnumValue(iconKey, 0)[1]
                    ).group()
                    print(f'Aspen Plus {version} found')
                    count += 1
                i += 1
            
            except:
                if count == 0:
                    version = None
                    print("can't find Aspen Plus installed")
                break

        return version
        
        
    def get_value(self, loc):
        '''
        Parameters
        ----------
        loc: str
            Path to the node in the Aspen Plus tree.
        
        Returns
        -------
        value: float or str
            The value of the node.
        '''
        
        return self.aspenCOM.Tree.FindNode(loc).Value
        
        
    def set_value(self, loc, value, ifFortran):
        '''
        Parameters
        ----------
        loc: str
            Path to the node in the Aspen Plus tree.
        value: float or str
            The value to write to the node.
        ifFortran: bool
            Whether the node is a Fortran variable
        '''
        
        node = self.aspenCOM.Tree.FindNode(loc)

        if ifFortran:
            oldValue = node.Value
            node.Value = re.sub(r'(?<== ).+', str(value), oldValue)
        else:
            node.Value = float(value)
        

    def run_model(self):
        self.aspenCOM.Reinit()
        self.aspenCOM.Engine.Run2()
        
        
    def save_model(self, saveFile):
        '''
        Parameters
        ----------
        saveFile: str
            The file name to save the Aspen Plus file (.bkp) to. 
        '''
    
        self.aspenCOM.SaveAs(saveFile)
        
    
    def close(self):
        self.aspenCOM.Close()
    
        
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
    
    dataInfo = pd.read_excel(data_file, sheet_name=['Inputs', 'Output'])
    inputInfo = dataInfo['Inputs']
    outputInfo = dataInfo['Output']
    
    return inputInfo, outputInfo
    
    
def run_and_update(
        data_file, 
        input_infos, 
        output_info, 
        aspen_file, 
        calculator_file,
        zetoolkit_file=None
    ):    
    '''
    Parameters
    ----------
    data_file: str
        Path to the data file.
    input_infos: pd.DataFrame
        A DataFrame containing the input information with columns 'Input variable', 
        'Type', 'Location', and 'Values'.
    output_info: pd.DataFrame
        A DataFrame containing the output information with columns 'Output 
        variable', 'Location', and 'Values'.
    aspen_file: str
        Path to the Aspen model file.
    calculator_file: str
        Path to the Excel calculator file.
    zetoolkit_file: str
        Path to the zetoolkit.dll.
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
        # run
        outDir = os.path.dirname(data_file)
        tmpDir = outDir + '/tmp'
        os.makedirs(tmpDir, exist_ok=True)
    
        aspenModel = Aspen(aspen_file)
        calculator = Excel(calculator_file, aspenModel.version, zetoolkit_file)
        
        print(f'{nruns} runs in total, {nrunsLeft} runs left')
        for i in range(nrunsCompl, nruns):
            print(f'run {i+1}:')
            
            # set Aspen variables
            for inputInfo in inputInfos:
                if inputInfo.type == 'bkp':
                    aspenModel.set_value(
                        inputInfo.loc, inputInfo.values[i], False
                    )   
                elif inputInfo.type == 'bkp_fortran':
                    aspenModel.set_value(
                        inputInfo.loc, inputInfo.values[i], True
                    )
                else:
                    continue

            # run Aspen model
            aspenModel.run_model()
            
            tmpFile = f'{tmpDir}/{i}.bkp'
            aspenModel.save_model(tmpFile)

            # set calculator variables
            for inputInfo in inputInfos:
                if inputInfo.type == 'xlsm':
                    inputSheet, inputCell = inputInfo.loc.split('!')
                    calculator.set_cell(
                        inputInfo.values[i], inputSheet, loc=inputCell
                    )
                else:
                    continue
            
            # run calculator
            calculator.load_aspenModel(tmpFile)
            calculator.calculate_dcfror()
            
            outputSheet, outputCell = outputInfo.loc.split('!')
            output = calculator.get_cell(outputSheet, loc=outputCell)
            outputInfo.values.append(output)
            
            # update dataset
            outputValues = ','.join(map(str, outputInfo.values))
            output_info = pd.DataFrame(
                [[outputInfo.name, outputInfo.loc, outputValues]],
                columns = ['Output variable', 'Location', 'Values']
            )
            
            with pd.ExcelWriter(data_file) as writer:
                input_infos.to_excel(writer, sheet_name='Inputs', index=False)
                output_info.to_excel(writer, sheet_name='Output', index=False)
            
            print('done')
        
        aspenModel.close()
        calculator.close()
        
    else:
        print('all runs completed')
    
    
    
    
if __name__ == '__main__':
    inputsInfo, outputInfo = parse_data_file(DATASET_FILE)
    if 'ZETOOLKIT_FILE' not in globals():
        ZETOOLKIT_FILE = None
    run_and_update(
        DATASET_FILE, 
        inputsInfo, 
        outputInfo, 
        ASPEN_FILE, 
        CALCULATOR_FILE,
        ZETOOLKIT_FILE
    )
