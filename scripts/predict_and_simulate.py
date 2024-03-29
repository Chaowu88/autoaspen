r'''
Predicts the output Minimum Selling Price (MSP) using a trained regression model and conducts 
the following analyses:
1. Sensitivity analysis with single query input.
2. Response analysis with two query inputs.
3. Monte Carlo simulation with three or more query inputs.

NOTES:
1. Variables can be formatted as "varname" or "varname (unit)".
2. Variable names must match between the "xxx-input(s)" sheet and the "baseline" sheet.
3. The variable order in the "baseline" sheet should match that in the "Inputs" sheet of the 
   generate_dataset_template.py config file.

Usage:
python path\to\predict_and_simulate.py
'''


__author__ = 'Chao Wu'


import os
import re
from itertools import product
from collections import namedtuple
import warnings
warnings.filterwarnings('ignore')
import numpy as np
import pandas as pd
from scipy import stats
from joblib import load
import matplotlib.pyplot as plt
import seaborn as sns


OUT_DIR = 'path\to\output'
CONFIG_FILE = 'path\to\config.xlsx'
MODEL_FILE = 'path\to\model.mod'
LABEL = 'MFSP (\$ GGE$^{-1}$)'


class BaseHandler:
    
    def __init__(self, config, baseline):
        '''
        Parameters
        ----------
        config: pd.DataFrame
            A DataFrame containing the configuration for the random input samples with columns 
            'Input variable', 'Bounds', 'Distribution', 'Parameters', and 'Size'.
        baseline: pd.DataFrame
            A DataFrame containing the baseline values for the input variables with columns 
            'Input variable' and 'Baseline value'.
        '''
        
        self.config = config
        self.baseline = baseline
        
        
    @staticmethod
    def generate_random_values(dist_name, size, bounds, *params):
        '''
        Parameters
        ----------
        dist_name: str
            The name of the distribution to use to generate the random values. Supported 
            distributions are: "normal", "alpha", "beta", "gamma", "triang", "pareto" and 
            "bernoulli".
        size: int
            The number of random values to generate.
        bounds: tuple
            A tuple of two values, representing the lower and upper bounds of the 
            random values.
        params: tuple
            A tuple of parameters for the specified distribution.
                normal: mean, sd
                alpha: a, loc, scale
                beta: a, b, loc, scale
                gamma: a, loc, scale
                triang: c, loc, scale
                pareto: b, loc, scale
                bernoulli: pl, ph
        
        Returns
        -------
        values: np.array
            A NumPy array containing the generated random values.
        '''
        
        dist = getattr(stats, dist_name)
        lb, ub = bounds
        
        if dist_name == 'uniform':
            values = dist.rvs(loc = lb, scale = ub-lb, size = size)

        elif dist_name == 'bernoulli':
            pl, ph = params
            labels = dist.rvs(pl, size = size)
            values = [lb if label else ub for label in labels]
        
        else:
            *shapeParams, loc, scale = params
            
            values = []
            count = 0
            while count < size:
                value = dist.rvs(*shapeParams, loc = loc, scale = scale)
                if lb <= value <= ub:
                    count += 1
                    values.append(value)
        values = np.array(values)            
                    
        return values

    
    def load_model(self, model_file, scaler_file = None):
        '''
        Parameters
        ----------
        model_file: str
            Path to the model file.
        scaler_file: str or None
            Path to the scaler file.
        '''
        
        self.model = load(model_file)

        if scaler_file:
            self.scaler = load(scaler_file)
        
        
    def simulate(self):
        
        self.outputs = []
        Output = namedtuple('Output', ['name', 'values'])
        for singleInput in self.inputs:
            print(singleInput.name, ': simulating')
            
            if hasattr(self, 'scaler'):
                inputData = self.scaler.transform(singleInput.data)
            else:
                inputData = singleInput.data
            predicted = self.model.predict(inputData)
            
            singleOutput = Output(singleInput.name, predicted)
            self.outputs.append(singleOutput)
            
    
    def plot_hist_and_save(self, out_dir, folder_name, label, percentile = 5):
        '''
        Parameters
        ----------
        out_dir: str
            Path to the output directory.
        folder_name: str
            Name of the folder containing the data to plot.
        label: str
            The label for the x-axis.
        percentile: float
            A value in the range [0, 100]. Lines indicating the percentile% and 
            (100 - percentile)% will be plotted.
        '''
        
        for singleOutput in self.outputs:
            print(singleOutput.name, ': plotting')
            
            varName = get_var_name(singleOutput.name)
            values = singleOutput.values
            fileName = get_var_name(label)
            
            saveDir = f'{out_dir}/{folder_name}/{varName}'
            saveDir = make_dir(saveDir)
            
            fig, ax1 = plt.subplots()
            sns.distplot(
                values, 
                rug = True, 
                kde = False, 
                hist = True, 
                ax = ax1,
                hist_kws = {'color': 'forestgreen'}
            )
            ax1.set_xlabel(label, fontsize = 20)
            ax1.set_ylabel('Count', color = 'forestgreen', fontsize = 20)
            ax1.tick_params(labelsize = 15)
            
            ax2 = ax1.twinx()
            sns.distplot(
                values, 
                rug = True, 
                kde = True, 
                hist = False, 
                ax = ax2, 
                hist_kws = {'color': 'royalblue'}, kde_kws = {'linewidth': 2.5}
            )   
            ax2.set_ylabel('')
            ax2.set_yticks([])
            ax2.spines['left'].set_visible(False)
            ax2.spines['right'].set_visible(False)
            ax2.spines['top'].set_visible(False)
            ax2.spines['bottom'].set_visible(False)
            
            bins = int(values.size/10) if values.size > 20 else values.size
            counts, edges = np.histogram(values, bins = bins)
            x = (edges[:-1]+edges[1:])/2
            y = np.cumsum(counts)/np.sum(counts)
            p1, p2 = np.percentile(values, [percentile, 100-percentile])
            
            ax3 = ax1.twinx()
            ax3.plot(x, y, color = 'peru', linewidth = 2.5)
            ax3.set_ylabel(
                'Cumulative probabilty', 
                color = 'peru', 
                fontsize = 20, 
                labelpad = 25, 
                rotation = 270
            )
            ax3.tick_params(labelsize = 15)

            ax3.vlines(x = p1, ymin = 0, ymax = 0.97, linestyles = 'dashed', color = 'gray')
            ax3.text(x = p1, y = 1, s = round(p1, 2), transform = ax3.transData, ha = 'center')
            ax3.vlines(x = p2, ymin = 0, ymax = 0.97, linestyles = 'dashed', color = 'gray')
            ax3.text(x = p2, y = 1, s = round(p2, 2), transform = ax3.transData, ha = 'center')
            
            fig.savefig(f'{saveDir}/{fileName}.jpg', dpi = 300, bbox_inches = 'tight')
            plt.close(fig = fig)
            
            pd.Series(values).to_excel(f'{saveDir}/{fileName}.xlsx', header = False, index = False)
    
    
    def plot_contour_and_save(self, out_dir, folder_name, label):
        '''
        Parameters
        ----------
        out_dir: str
            Path to the output directory.
        folder_name: str
            Name of the folder containing the data to plot.
        label: str
            The label for the x-axis.
        '''
        
        for singleInput, singleOutput in zip(self.inputs, self.outputs):
            print(singleOutput.name, ': plotting')
            
            varName = get_var_name(singleOutput.name)
            fileName = get_var_name(label)
            
            saveDir = f'{out_dir}/{folder_name}/{varName}'
            saveDir = make_dir(saveDir)
            
            xvar, yvar = singleInput.name.split('_')
            xvar = correct_unit(xvar)
            yvar = correct_unit(yvar)
            
            X, Y = np.meshgrid(singleInput.xvalues, singleInput.yvalues)
            Z = singleOutput.values.reshape(singleInput.xvalues.size, singleInput.yvalues.size).T
            
            fig, ax = plt.subplots()
    
            ctf = ax.contourf(X, Y, Z, 50, cmap = plt.cm.get_cmap('RdBu').reversed())
            ax.set_xlabel(xvar, fontsize = 20)
            ax.set_ylabel(yvar, fontsize = 20)
            ax.tick_params(labelsize = 15)
            
            cbar = fig.colorbar(ctf)
            cbar.set_label(label, labelpad = 25, rotation = 270, fontsize = 20)
            
            if Z.max() - Z.min() > 0.001:    
                ct = ax.contour(
                    X, Y, Z, 
                    ctf.levels[1::6], 
                    colors = 'dimgray', 
                    linewidths = 1, 
                    linestyles ='dashed'
                )
                ax.clabel(ct, ctf.levels[1::6], inline = True, fontsize = 10, colors = 'k')
            
            fig.savefig(f'{saveDir}/{fileName}.jpg', dpi = 300, bbox_inches = 'tight')
            plt.close(fig = fig)
            
            pd.DataFrame(Z).to_excel(f'{saveDir}/{fileName}.xlsx', header = xvar, index = yvar)
        

class OneInputHandler(BaseHandler):
    
    def generate_input_matrix(self):
        
        self.inputs = []
        Input = namedtuple('Input', ['name', 'data'])
        for _, [inputVar, bnds, distName, params, size] in self.config.iterrows():
            
            inputVar = inputVar.strip()
            print(inputVar, ': generating input values')
            
            bnds = map(float, bnds.split(','))
            if isinstance(params, str):
                params = map(float, params.split(','))
            elif np.isnan(params):
                params = ()
            size = int(size)
            values = self.generate_random_values(distName, size, bnds, *params)
            
            baseInput = self.baseline.set_index('Input variable').T
            baseInputMat = pd.concat([baseInput]*size, ignore_index = True)
            
            baseInputMat[inputVar] = values
            
            singleInput = Input(inputVar, baseInputMat.values)
            self.inputs.append(singleInput)
            
    
    def plot_and_save(self, out_dir, label):
        '''
        Parameters
        ----------
        out_dir: str
            Path to the output directory.
        label: str
            The label for the x-axis.
        '''
        
        self.plot_hist_and_save(out_dir, 'one_input', label)
        

class TwoInputsHandler(BaseHandler):
    
    def generate_input_matrix(self):
        
        self.inputs = []
        Input = namedtuple('Input', ['name', 'xvalues', 'yvalues', 'data'])
        for _, [inputVars, bndss, sizes] in self.config.iterrows():
            
            inputVarX, inputVarY = [var.strip() for var in inputVars.split('|')]
            print(inputVarX+'_'+inputVarY, ': generating input values')
            
            bndsX, bndsY = (tuple(map(float, bnds.split(','))) for bnds in bndss.split('|'))
            sizeX, sizeY = map(int, sizes.split('|'))
            
            valuesX = np.linspace(*bndsX, sizeX)
            valuesY = np.linspace(*bndsY, sizeY)
            
            baseInput = self.baseline.set_index('Input variable').T
            baseInputMat = pd.concat([baseInput]*sizeX*sizeY, ignore_index = True)
            
            baseInputMat[[inputVarX, inputVarY]] = list(product(valuesX, valuesY))
            
            singleInput = Input(inputVarX+'_'+inputVarY, valuesX, valuesY, baseInputMat.values)
            self.inputs.append(singleInput)
        
    
    def plot_and_save(self, out_dir, label):
        '''
        Parameters
        ----------
        out_dir: str
            Path to the output directory.
        label: str
            The label for the x-axis.
        '''
        
        self.plot_contour_and_save(out_dir, 'two_input', label)
    
    
class MoreInputsHandler(BaseHandler):
    
    def generate_input_matrix(self):
            
        self.inputs = []
        Input = namedtuple('Input', ['name', 'data'])
        for _, [inputVars, bndss, distNames, paramss, size] in self.config.iterrows():
        
            baseInput = self.baseline.set_index('Input variable').T
            baseInputMat = pd.concat([baseInput]*size, ignore_index = True)
            
            inputVars = [var.strip() for var in inputVars.split('|')]
            print('_'.join(inputVars), ': generating input values')
            
            bndss = bndss.split('|')
            distNames = distNames.split('|')
            paramss = paramss.split('|')
            size = int(size)
            for inputVar, bnds, distName, params in zip(inputVars, bndss, distNames, paramss):
            
                bnds = map(float, bnds.split(','))
                if params == '':
                    params = ()
                else:
                    params = map(float, params.split(','))
                values = self.generate_random_values(distName, size, bnds, *params)
                
                baseInputMat[inputVar] = values
            
            singleInput = Input('_'.join(inputVars), baseInputMat.values)
            self.inputs.append(singleInput)
                
    
    def plot_and_save(self, out_dir, label):
        '''
        Parameters
        ----------
        out_dir: str
            Path to the output directory.
        label: str
            The label for the x-axis.
        '''
        
        self.plot_hist_and_save(out_dir, 'more_input', label)
        

def parse_config_file(config_file):
    '''
    Parameters
    ----------
    config_file: str
        Path to the configuration file. Note that the order of variables in "Baseline" sheet 
        should be identical with the order of features used in model. 
    
    Returns
    -------
    oneInput: pd.DataFrame
        A DataFrame containing the configuration for one input variable.
    twoInputs: pd.DataFrame
        A DataFrame containing the configuration for two input variables.
    moreInputs: pd.DataFrame
        A DataFrame containing the configuration for more than two input variables.
    baseline: pd.DataFrame
        A DataFrame containing the baseline values for the input variables.
    '''
    
    configInfo = pd.read_excel(
        config_file, 
        sheet_name = ['One-input', 'Two-inputs', 'More-inputs', 'Baseline']
    )
    oneInput = configInfo['One-input']
    twoInputs = configInfo['Two-inputs']
    moreInputs = configInfo['More-inputs']
    baseline = configInfo['Baseline']
    baseline['Input variable'] = baseline['Input variable'].str.strip()
    
    return oneInput, twoInputs, moreInputs, baseline
    
    
def get_var_name(name_with_unit):        
    '''
    Parameters
    ----------
    name_with_unit: str
        Variable name with unit.
    
    Returns
    -------
    name: str
        Variable name.
    '''
    
    name = re.sub(r'\s*\(.*?\)\s*', '', name_with_unit)

    return name
    

def correct_unit(name_with_unit):
    '''
    Parameters
    ----------
    name_with_unit: str
        Variable name with unit.
    
    Returns
    -------
    name_with_corr_unit: str
        Variable name with correlated unit.
    '''    

    name_with_unit = re.sub(r'\$', '\$', name_with_unit)   
    name_with_corr_unit = re.sub(r'\((.+)/(.+)\)', '(\g<1> \g<2>$^{-1}$)', name_with_unit)

    return name_with_corr_unit

    
def make_dir(directory):
    '''
    Parameters
    ----------
    directory: str
        Path to the directory to create.
    '''
    
    try:
        os.makedirs(directory, exist_ok = True)
    
    except FileNotFoundError:
        directory = directory[:220].strip()
        os.makedirs(directory, exist_ok = True)
        
    finally:
        return directory
        
        
        

if __name__ == '__main__':
    
    *configs, baseline = parse_config_file(CONFIG_FILE)
    Handlers = [OneInputHandler, TwoInputsHandler, MoreInputsHandler]
    labels = ['one input variable', 'two input variables', 'more input variables']
    
    for config, Handler, label in zip(configs, Handlers, labels):
        print(f'\nhandle {label}:')
        
        if not config.empty:
            handler = Handler(config, baseline)
            handler.generate_input_matrix()
            handler.load_model(MODEL_FILE)
            handler.simulate()
            handler.plot_and_save(OUT_DIR, LABEL)
            