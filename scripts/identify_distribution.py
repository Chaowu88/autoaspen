#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '06/09/2021'
__version__ = '1.2'


r'''
This script identifies the distribution of a continuous variable by fitting to the following unimodal distributions: "alpha", "beta", "triangular", "normal", "gamma" and "pareto", and optionally generate random values subject to the identified distribution

python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\AutoAspen2\identify_distribution.py
'''


OUT_DIR = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\AutoAspen_paper\case_1_cellulosic\distributions\loan_interest'
DATA_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Data\AutoAspen_paper\case_1_cellulosic\loan_interest.xlsx'
DATA_LABEL = 'Monthly values'
DISTRIBUTIONS = ['alpha', 'gamma', 'beta', 'triang', 'norm', 'pareto', 'uniform']

NEED_RANDOM_VALUES = False
SIZE = 5000   # works if NEED_RANDOM_VALUES == True


import os
from collections import namedtuple
import numpy as np
import pandas as pd
from scipy import stats
from scipy.stats import kstest
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings("ignore")


def read_data(data_file):
    '''
    Parameters
    ----------
    data_file: str
        data file
    
    Returns
    -------
    data: ser
    '''
    
    data = pd.read_excel(data_file, header = 0, index_col = 0, squeeze = True)
    
    return data
    
    
def identify_distribution(data, distributions):
    '''
    Parameters
    ----------
    data: ser
    distributions: list of str
        distributions to test
    
    Returns
    FitInfo: list of namedtuples
    '''
    
    fitInfos = []
    FitInfo = namedtuple('FitInfo', ['dist_name', 'shape_params', 'loc', 'scale', 'pvalue', 'pdf'])
    for distName in distributions:
        
        # fit to known distribution
        dist = getattr(stats, distName)
        params = dist.fit(data)
        
        pvalue = kstest(data, distName, args = params)[1]
        
        *shapeParams, loc, scale = params
        
        print('%s pvalue: %.4f\nparams: %s, loc: %.4f, scale: %.4f' % (distName, pvalue, shapeParams, loc, scale))
        
        # generate PDF of fitted distribution
        xstart = dist.ppf(0.01, *shapeParams, loc = loc, scale = scale)
        xend = dist.ppf(0.99, *shapeParams, loc = loc, scale = scale)
        xend = min(xend, data.max()*1.2)
        xs = np.linspace(xstart, xend, 1000)
        
        PDFarr = dist.pdf(xs, *params[:-2], loc = params[-2], scale = params[-1])
        PDFser = pd.Series(PDFarr, index = xs)
        
        fitInfo = FitInfo(distName, shapeParams, loc, scale, pvalue, PDFser)
        fitInfos.append(fitInfo)
        
    return fitInfos
    

def plot_results(out_dir, data, data_label, fit_infos):
    '''
    Parameters
    ----------
    out_dir: str
        output directory
    data: ser
    data_label: str
        data label for xaxes
    fit_infos: list of namedtuples
        fields are ['dist_name', 'shape_params', 'loc', 'scale', 'pvalue', 'pdf']
    '''
    
    os.makedirs(out_dir, exist_ok = True)
    
    plt.hist(data, bins = 50 if data.size > 100 else 10)
    plt.xlabel(data_label)
    plt.ylabel('Count')

    ax = plt.twinx()
    for fitInfo in fit_infos:
        ax.plot(fitInfo.pdf.index, fitInfo.pdf.values, label = fitInfo.dist_name)

    ax.set_ylabel('Probability density function')
    ax.legend()

    plt.savefig('%s/fitted_distributions.jpg' % out_dir, dpi = 300, bbox_inches = 'tight')


def generate_random_values(out_dir, data, fit_infos, size):
    '''
    Parameters
    ----------
    out_dir: str
        output directory
    data: ser
    fit_infos: list of namedtuples
        fields are ['dist_name', 'shape_params', 'loc', 'scale', 'pvalue', 'pdf']
    size: int
        # of random values to generate
    '''
    
    os.makedirs(out_dir, exist_ok = True)
    
    lb, ub = data.min(), data.max()
    
    maxPvalue = 0
    for fitInfo in fit_infos:
        if fitInfo.pvalue >= maxPvalue:
            maxPvalue = fitInfo.pvalue
            distName = fitInfo.dist_name
            shapeParams = fitInfo.shape_params
            loc = fitInfo.loc
            scale = fitInfo.scale
    
    dist = getattr(stats, distName)
    
    values = []
    count = 0
    while count < size:
        value = dist.rvs(*shapeParams, loc = loc, scale = scale)
        if lb <= value <= ub:
            count += 1
            values.append(value)
    
    pd.Series(values).to_excel('%s/random_values.xlsx' % out_dir, header = False, index = False)    
    
    

    
if __name__ == '__main__':
    
    data = read_data(DATA_FILE)
    
    fittedInfos = identify_distribution(data, DISTRIBUTIONS)
    
    plot_results(OUT_DIR, data, DATA_LABEL, fittedInfos)
    
    if NEED_RANDOM_VALUES:
        generate_random_values(OUT_DIR, data, fittedInfos, SIZE)
