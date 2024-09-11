# AutoAspen
AutoAspen is a comprehensive tool for conducting stochastic techno-economic analysis (TEA) using machine learning techniques. It has been developed to facilitate dataset creation, model training, parameter optimization, and Monte Carlo sampling for estimating the distributions of critical parameters, such as the minimum selling price (MFSP), in the context of chemical engineering pathways. With the power of trained models, AutoAspen is capable of:

1. Univariate Uncertainty Analysis that evaluates the impact of individual variables on the MFSP while holding all other factors constant.

2. Bivariate Uncertainty Analysis that provides insights into how the MFSP responds to variations in pairs of variables, which are visually represented in contour plots.

3. Multivariate Uncertainty Analysis that allows users to modify a set of input variables simultaneously during a simulation.

## Prerequisites
The AP V10 compatible version was developed and tested on Python 3.8 with the following dependencies: numpy 1.20.2, pandas 1.0.5, scipy 1.5.0, scikit-learn 0.23.1, matplotlib 3.2.2, seaborn 0.12.0, pillow 7.2.0, xlrd 1.2.0, openpyxl, pywin32.

The AP V12 compatible version was developed and tested on Python 3.10 with the following dependencies: numpy 1.26.0, pandas 2.2.2, scipy 1.14.0, scikit-learn 1.5.1, matplotlib 3.9.2, seaborn 0.13.2, pillow 7.2.0, xlrd 1.2.0, openpyxl, pywin32.

## Usage
1. `generate_dataset_template.py`: This script creates a dataset template for training by generating random values for input variables based on specified distributions defined in the `var_info` file. Supported distributions include: normal, alpha, beta, gamma, triangular, pareto, and bernoulli. A sample `var_info` file is provided [here](https://github.com/Chaowu88/autoaspen/blob/main/ATJ_pathway/var_info.xlsx).

2. `generate_dataset.py`: It generates a training dataset by invoking the Aspen model and the .xslm calculator. This step automates the traditional stochastic TEA, which usually involves iterative calls to Aspen Plus software.

3. `train_regression_model.py`: This script trains and fine-tunes four machine learning models. Users can select the best-performing model for Monte Carlo simulation. Supported machine learning models include: polynomial ridge regression, linear SVM, random forest, and gradient tree boosting.

4. `predict_and_simulate.py`: This script predicts the MFSP using a trained regression model and conducts univariate, bivariate, or multivariate uncertainty analysis. The distributions and baseline values of input variables are specified in the [`config`](https://github.com/Chaowu88/autoaspen/blob/main/ATJ_pathway/config.xlsx) file.
