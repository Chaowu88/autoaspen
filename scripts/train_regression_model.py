#!/usr/bin/env pyhton
# -*- coding: UTF-8 -*-


__author__ = 'Chao Wu'
__date__ = '12/07/2022'


r'''
This script trains and tunes four machine learning model:
1. polynomial ridge regression;
2. linear SVM;
3. random forest;
4. gradient tree boosting

python C:\Users\cwu\Desktop\Software\Aspen_automation\Scripts\AutoAspen2\train_regression_model.py
'''


OUT_DIR = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\AutoAspen_paper\case_3_FT\training'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\AutoAspen_paper\case_2_HEFA\training'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\AutoAspen_paper\case_1_cellulosic\training'
DATA_FILE = r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\AutoAspen_paper\case_3_FT\training\training_data.xlsx'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\AutoAspen_paper\case_2_HEFA\training\training_data.xlsx'
#r'C:\Users\cwu\Desktop\Software\Aspen_automation\Results\AutoAspen_paper\case_1_cellulosic\training\training_data.xlsx'
METHODS = {'poly': 'polynomial ridge regression',
           'lsvm': 'linear SVM',
           'rf': 'random forest',
           'gtb': 'gradient tree boosting'}
LABEL = 'MFSP (\$ GGE$^{-1}$)'           


import sys
import warnings
import os
if not sys.warnoptions:   
    warnings.simplefilter('ignore')
    os.environ['PYTHONWARNINGS'] = 'ignore'
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import PolynomialFeatures, StandardScaler
from sklearn.linear_model import Ridge
from sklearn.svm import LinearSVR
from sklearn.ensemble import RandomForestRegressor
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import GridSearchCV
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
#from scipy.stats import pearsonr
import matplotlib.pyplot as plt
from joblib import dump


class Model():
    '''
    Parameters:
    ----------
    method: {'poly', 'lsvm', 'rf', 'gtb'}
        machine learning models
    njobs: int
        # of jobs in parallel
    '''

    def __init__(self, method, njobs = 3):
        '''
        Parameters:
        ----------
        method: {'poly', 'lsvm', 'rf', 'gtb'}
            machine learning models
        njobs: int
            # of jobs in parallel
        '''
        
        if method == 'poly':
            self.model = Pipeline(steps = [('poly', PolynomialFeatures()), 
                                           ('ridge', Ridge(random_state = 0))])
            self.paramGrid = {'poly__degree': [1, 2, 3, 4, 5],
                              'ridge__alpha': [0.1, 1, 5, 10],
                              'ridge__fit_intercept': [True, False],
                              'ridge__normalize': [True, False]}
        elif method == 'lsvm':
            self.model = Pipeline(steps = [('scale', StandardScaler()),
                                           ('lsvm', LinearSVR(random_state = 0))])
            self.paramGrid = [{'lsvm__C': [0.01, 0.1, 1, 10], 
                               'lsvm__epsilon': [0, 0.1, 1, 10], 
                               'lsvm__fit_intercept': [True, False],
                               'lsvm__loss': ['epsilon_insensitive', 'squared_epsilon_insensitive'], 
                               'lsvm__dual': [True]},
                              {'lsvm__C': [0.01, 0.1, 1, 10], 
                               'lsvm__epsilon': [0, 0.1, 1, 10], 
                               'lsvm__fit_intercept': [True, False],
                               'lsvm__loss': ['squared_epsilon_insensitive'], 
                               'lsvm__dual': [False]}]
        elif method == 'rf':
            self.model = RandomForestRegressor(random_state = 0, n_jobs = njobs)
            self.paramGrid = [{'n_estimators': [50, 100, 200], 
                               'max_depth': [10, 20, 30], 
                               'min_samples_split': [2, 5], 
                               'min_samples_leaf': [1, 5, 10], 
                               'max_features': ['auto', 'sqrt', 'log2']}]
        elif method == 'gtb':
            self.model = GradientBoostingRegressor(random_state = 0)
            self.paramGrid = [{'n_estimators': [50, 100, 200], 
                               'learning_rate': [0.1, 0.2, 0.3], 
                               'max_depth': [3, 5, 10], 
                               'min_samples_split': [2, 5], 
                               'min_samples_leaf': [1, 5, 10], 
                               'max_features': ['auto', 'sqrt', 'log2']}]
        self.method = method


    def train_and_tune(self, features, targets, nfolds = 5, njobs = 3):
        '''
        Parameters
        ----------
        features: df
            training features
        targets: ser
            training targets
        nfolds: int
            # of cross validation folds
        njobs: int
            # of jobs in parallel
        '''

        Xtrain, Xtest, Ytrain, Ytest = train_test_split(features, targets, test_size = 0.25, random_state = 0)

        regModels = GridSearchCV(self.model, self.paramGrid, cv = nfolds, n_jobs = njobs)
        
        regModels.fit(Xtrain, Ytrain)
        
        self.bestModel = regModels.best_estimator_
        self.bestParams = regModels.best_params_
        
        predicted = self.bestModel.predict(Xtest)
        #self.r2 = pearsonr(predicted, Ytest)[0]**2
        self.mse = mean_squared_error(Ytest, predicted)
        self.mae = mean_absolute_error(Ytest, predicted)
        self.r2 = r2_score(Ytest, predicted)
        self.true_and_pred = pd.DataFrame({'True': Ytest.values, 'Predicted': predicted})
        

    def display_results(self, out_dir):
        
        print('\nbest parameters:')
        for param, value in self.bestParams.items():
            print('%s = %s' % (param, value))
            
        # print('\ntrue vs predicted')
        # for _, (vpred, vtrue) in self.true_and_pred.iterrows():
        #     print(round(vpred, 4), round(vtrue, 4))
        
        print('\nMSE: %.4f' % self.mse)
        print('MAE: %.4f' % self.mae)
        print('R2: %.4f' % self.r2)

        metrics = pd.Series({'MES': self.mse, 'MAE': self.mae, 'R2': self.r2})
        metrics.to_excel('%s/metrics.xlsx' % out_dir, header = False)


    def save_model(self, out_dir):
        
        os.makedirs(out_dir, exist_ok = True)
        dump(self.bestModel, '%s/%s.mod' % (out_dir, self.method))


    def plot_true_vs_predicted(self, out_dir, label):
        
        fig, ax = plt.subplots()
        ax.scatter(self.true_and_pred['True'], self.true_and_pred['Predicted'], color = 'lightsalmon', 
                   edgecolors = 'gray', s = 120)
        ax.set_xlabel('True value of %s' % label, fontsize = 20)
        ax.set_ylabel('Predicted %s' % label, fontsize = 20)
        ax.tick_params(labelsize = 15)
        ax.text(0.1, 0.7, 'MSE = %.3f\nMAE = %.3f\n$R^2$ = %.3f' % (self.mse, self.mae, self.r2), 
                fontsize = 15, transform = ax.transAxes)
        
        lineLB, lineUB = self.true_and_pred.min().min(), self.true_and_pred.max().max()
        ax.plot([lineLB, lineUB], [lineLB, lineUB], linestyle = '--', linewidth = 4, 
                color = 'steelblue', zorder = 0)

        os.makedirs(out_dir, exist_ok = True)
        fig.savefig('%s/true_vs_predicted.jpg' % out_dir, dpi = 300, bbox_inches = 'tight')
        self.true_and_pred.to_excel('%s/true_vs_predicted.xlsx' % out_dir, header = True, index = False)


def read_data(data_file):
    '''
    Parameters
    ----------
    data_file: str
        data file
    
    Returns
    -------
    features: df
    targets: ser
    '''
    
    dataInfo = pd.read_excel(data_file, sheet_name = ['Inputs', 'Output'])
    inputInfo = dataInfo['Inputs']
    outputInfo = dataInfo['Output'].squeeze()
    
    inputValues = inputInfo['Values'].str.split(',')
    features = pd.DataFrame(dict(zip(inputInfo['Input variable'], inputValues)), dtype = float)
    
    outputValues = outputInfo['Values'].split(',')
    targets = pd.Series(outputValues, name = outputInfo['Output variable'], dtype = float)
    
    features = features.iloc[:targets.size, :]
    
    return features, targets




if __name__ == '__main__':
    
    features, targets = read_data(DATA_FILE)
    
    for method in METHODS:
        print('\ntuning %s\n' % METHODS[method])
        model = Model(method)
        model.train_and_tune(features, targets)
        model.save_model(OUT_DIR+'\\'+method)
        model.display_results(OUT_DIR+'\\'+method)
        model.plot_true_vs_predicted(OUT_DIR+'\\'+method, LABEL)

  