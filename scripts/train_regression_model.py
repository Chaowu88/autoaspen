r'''
Trains and fine-tunes four machine learning models:
1. polynomial ridge regression
2. linear SVM
3. random forest
4. gradient tree boosting

Usage:
python path\to\train_regression_model.py
'''


__author__ = 'Chao Wu'


import sys
import warnings
import os
if not sys.warnoptions:   
    warnings.simplefilter('ignore')
    os.environ['PYTHONWARNINGS'] = 'ignore'
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import PolynomialFeatures, MinMaxScaler, StandardScaler
from sklearn.linear_model import Ridge
from sklearn.svm import LinearSVR
from sklearn.ensemble import RandomForestRegressor
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import GridSearchCV
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
import matplotlib.pyplot as plt
from joblib import dump


OUT_DIR = 'path\to\output'
DATA_FILE = 'path\to\training_data.xlsx'
METHODS = {
    'poly': 'polynomial ridge regression',
    'lsvm': 'linear SVM',
    'rf': 'random forest',
    'gtb': 'gradient tree boosting'
}
LABEL = 'MFSP (\$ GGE$^{-1}$)'           


class Model():

    def __init__(self, method, scale = False, njobs = 3):
        '''
        Parameters:
        ----------
        method: str
            The type of machine learning model to train. Supported methods are "poly", "lsvm", 
            "rf", and "gtb".
        scale: bool
            Whether to scale the features before training the model.
        njobs: int
            The number of parallel jobs to use when training the model.
        '''
        
        self.scale = scale

        if method == 'poly':
            self.model = Pipeline(
                steps = [('poly', PolynomialFeatures()), ('ridge', Ridge(random_state = 0))]
            )
            self.paramGrid = {
                'poly__degree': [1, 2, 3, 4, 5],
                'ridge__alpha': [0.1, 1, 5, 10],
                'ridge__fit_intercept': [True, False],
                'ridge__normalize': [True, False]
            }
        
        elif method == 'lsvm':
            if self.scale:
                self.model = LinearSVR(random_state = 0)
                self.paramGrid = [
                    {
                        'C': [0.01, 0.1, 1, 10], 
                        'epsilon': [0, 0.1, 1, 10], 
                        'fit_intercept': [True, False],
                        'loss': ['epsilon_insensitive', 'squared_epsilon_insensitive'], 
                        'dual': [True]
                    },
                    {
                        'C': [0.01, 0.1, 1, 10], 
                        'epsilon': [0, 0.1, 1, 10], 
                        'fit_intercept': [True, False],
                        'loss': ['squared_epsilon_insensitive'], 
                        'dual': [False]
                    }
                ]
            else:
                self.model = Pipeline(
                    steps = [('scale', StandardScaler()), ('lsvm', LinearSVR(random_state = 0))]
                )
                self.paramGrid = [
                    {
                        'lsvm__C': [0.01, 0.1, 1, 10], 
                        'lsvm__epsilon': [0, 0.1, 1, 10], 
                        'lsvm__fit_intercept': [True, False],
                        'lsvm__loss': ['epsilon_insensitive', 'squared_epsilon_insensitive'], 
                        'lsvm__dual': [True]
                    },
                    {
                        'lsvm__C': [0.01, 0.1, 1, 10], 
                        'lsvm__epsilon': [0, 0.1, 1, 10], 
                        'lsvm__fit_intercept': [True, False],
                        'lsvm__loss': ['squared_epsilon_insensitive'], 
                        'lsvm__dual': [False]
                    }
                ]

        elif method == 'rf':
            self.model = RandomForestRegressor(random_state = 0, n_jobs = njobs)
            self.paramGrid = [
                {
                    'n_estimators': [50, 100, 200], 
                    'max_depth': [10, 20, 30], 
                    'min_samples_split': [2, 5], 
                    'min_samples_leaf': [1, 5, 10], 
                    'max_features': ['auto', 'sqrt', 'log2']
                }
            ]
            
        elif method == 'gtb':
            self.model = GradientBoostingRegressor(random_state = 0)
            self.paramGrid = [
                {
                    'n_estimators': [50, 100, 200], 
                    'learning_rate': [0.1, 0.2, 0.3], 
                    'max_depth': [3, 5, 10], 
                    'min_samples_split': [2, 5], 
                    'min_samples_leaf': [1, 5, 10], 
                    'max_features': ['auto', 'sqrt', 'log2']
                }
            ]
        
        self.method = method


    def fine_tune(self, features, targets, nfolds = 5, njobs = 3):
        '''
        Parameters
        ----------
        features: pd.DataFrame
            training features
        targets: pd.Series
            The training targets.
        nfolds: int
            The number of cross-validation folds.
        njobs: int
            The number of parallel jobs to use when tuning the model.
        '''

        datasets = train_test_split(features, targets, test_size = 0.25, random_state = 0)
        Xtrain, Xtest, Ytrain, Ytest = datasets

        if self.scale:
            self.scaler = MinMaxScaler()
            self.scaler.fit(Xtrain)
            Xtrain = self.scaler.transform(Xtrain)
            Xtest = self.scaler.transform(Xtest)

        regModels = GridSearchCV(self.model, self.paramGrid, cv = nfolds, n_jobs = njobs)
        
        regModels.fit(Xtrain, Ytrain)
        
        self.bestModel = regModels.best_estimator_
        self.bestParams = regModels.best_params_
        
        predicted = self.bestModel.predict(Xtest)
        self.mse = mean_squared_error(Ytest, predicted)
        self.mae = mean_absolute_error(Ytest, predicted)
        self.r2 = r2_score(Ytest, predicted)
        self.true_and_pred = pd.DataFrame({'True': Ytest.values, 'Predicted': predicted})

        

    def display_results(self, out_dir):
        
        print('best parameters:')
        for param, value in self.bestParams.items():
            print(f'{param} = {value}')
        
        print(f'\nMSE: {self.mse:.4f}')
        print(f'MAE: {self.mae:.4f}')
        print(f'R2: {self.r2:.4f}')

        metrics = pd.Series({'MES': self.mse, 'MAE': self.mae, 'R2': self.r2})
        metrics.to_excel(f'{out_dir}/metrics.xlsx', header = False)


    def save_model(self, out_dir):
        
        os.makedirs(out_dir, exist_ok = True)
        dump(self.bestModel, f'{out_dir}/{self.method}.mod')

        if self.scale:
            dump(self.scaler, f'{out_dir}/{self.method}.scaler')


    def plot_true_vs_predicted(self, out_dir, label):
        
        fig, ax = plt.subplots()
        
        ax.scatter(
            self.true_and_pred['True'], 
            self.true_and_pred['Predicted'], 
            color = 'lightsalmon', 
            edgecolors = 'gray', 
            s = 120
        )
        ax.set_xlabel(f'True value of {label}', fontsize = 20)
        ax.set_ylabel(f'Predicted {label}', fontsize = 20)
        ax.tick_params(labelsize = 15)
        ax.text(
            0.1, 
            0.7, 
            f'MSE = {self.mse:.3f}\nMAE = {self.mae:.3f}\n$R^2$ = {self.r2:.3f}', 
            fontsize = 15, 
            transform = ax.transAxes
        )
        
        lineLB, lineUB = self.true_and_pred.min().min(), self.true_and_pred.max().max()
        ax.plot(
            [lineLB, lineUB], 
            [lineLB, lineUB], 
            linestyle = '--', 
            linewidth = 4, 
            color = 'steelblue', 
            zorder = 0
        )

        os.makedirs(out_dir, exist_ok = True)
        fig.savefig(f'{out_dir}/true_vs_predicted.jpg', dpi = 300, bbox_inches = 'tight')
        self.true_and_pred.to_excel(
            f'{out_dir}/true_vs_predicted.xlsx', 
            header = True, 
            index = False
        )


def read_data(data_file):
    '''
    Parameters
    ----------
    data_file: str
        Path to the data file.
    
    Returns
    -------
    features: pd.DataFrame
        A DataFrame containing the training features.
    targets: pd.Series
        A Series containing the training targets.
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
        print(f'\ntuning {METHODS[method]}\n')
        model = Model(method)
        model.fine_tune(features, targets)
        model.save_model(f'{OUT_DIR}/{method}')
        model.display_results(f'{OUT_DIR}/{method}')
        model.plot_true_vs_predicted(f'{OUT_DIR}/{method}', LABEL)

  