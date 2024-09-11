r'''
Removes sample(s) from the "Values" column of the training data.

Usage:
python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\AutoAspen2\delete_samples.py
'''


__author__ = 'Chao Wu'


import pandas as pd


OUT_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\test_delete_samples\test_out.xlsx'
DATA_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\test_delete_samples\test.xlsx'
DELETE_INDEX = [0, 2]   # starting from 0
DELETE_RESULT = True   # whether to remove the corresponding values in the results as well


def delete_data(out_file, data_file, index, delete_result=True):
    '''
    Parameter
    ---------
    out_file: str
        Path to the output file.
    data_file: str
        Path to the data file.
    index: list
        List of indices to delete, starting from 0.
    delete_result: bool
        whether to delete corresponding values in the results.
    '''

    dataInfo = pd.read_excel(data_file, sheet_name=['Inputs', 'Output'])
    
    outputInfo = dataInfo['Output'].squeeze()
    results = pd.Series(outputInfo['Values'].split(','))
    
    inputInfo = dataInfo['Inputs']
    samples = pd.DataFrame(inputInfo['Values'].str.split(',').tolist(), 
                           index=inputInfo['Input variable'],
                           columns=results.index)

    if delete_result:
        results = results.drop(index)
    samples = samples.drop(columns=index)

    outputInfo['Values'] = ','.join(results)
    outputInfo = outputInfo.to_frame().T
    inputInfo['Values'] = samples.apply(lambda r: ','.join(r), axis=1).values

    with pd.ExcelWriter(out_file) as writer:
        inputInfo.to_excel(writer, sheet_name='Inputs', index=False)
        outputInfo.to_excel(writer, sheet_name='Output', index=False)
        writer.save()




if __name__ == '__main__':
    delete_data(OUT_FILE, DATA_FILE, DELETE_INDEX)