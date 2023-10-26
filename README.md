# AutoAspen
AutoAspen is a tool for stochastic techno-economic analysis (TEA) using machine learning methods. It is developed to streamline dataset creation, model training, parameter optimization, and Monte Carlo sampling for generating distributions of the minimum selling price (e.g., MFSP) for specified chemical engineering pathway. With trained models, it can perform:
    
__1. Univariate uncertainty analysis__ that assesses how individual variables affect the MFSP while keeping all other variables constant.
    
__2. Bivariate uncertainty analysis__ that provides insights into how the MFSP value responds to variations in paired viables, illustrated in contour plot.

__3. Multivariate uncertainty analysis__ that offers users the flexibility to specify a set of input variables that will be simultaneously modified during a simulation.
## Dependencies
AutoAspen was developed and test underPython 3.8 with the dependencies:

numpy 1.20.2, pandas 1.0.5, scipy 1.5.0, scikit-learn 0.23.1, matplotlib 3.2.2, seaborn 0.12.0, pillow 7.2.0, xlrd 1.2.0, pythoncom and win32com
## Usage
__generate_dataset_template.py__ creates a dataset template for training by generating random values for input variables according to specified distributions defined in var_info. Supported distributions include: normal, alpha, beta, gamma, triangular, pareto and bernoulli.
A demo var_info file can be found [here](https://github.com/Chaowu88/autoaspen/blob/main/ATJ_pathway/var_info.xlsx).

__generate_dataset.py__ generates a training dataset by calling the Aspen model and the .xslm calculator. This step serves as the automation of traditional stochastic TEA reliant on iterative calls to Aspen software.

__train_regression_model__ trains and fine-tunes four machine learning models. User can select the one with best performance for Monte Carlo simulation. Support machine learning models include: polynomial ridge regression, linear SVM, random forest and gradient tree boosting.

__predict_and_simulate__ predicts the output minimum selling price using a trained regression model and conducts univariate, bivariate or multivariate uncertainty analysis. The distributions and baseline values of input variables are specified in [config](https://github.com/Chaowu88/autoaspen/blob/main/ATJ_pathway/config.xlsx) file.

# AutoAspen
AutoAspen is a comprehensive tool for conducting stochastic techno-economic analysis (TEA) using machine learning techniques. It has been developed to facilitate dataset creation, model training, parameter optimization, and Monte Carlo sampling for estimating the distributions of critical parameters, such as the minimum selling price (MFSP), in the context of chemical engineering pathways. With the power of trained models, AutoAspen is capable of:

1. Univariate Uncertainty Analysis that evaluates the impact of individual variables on the MFSP while holding all other factors constant.

2. Bivariate Uncertainty Analysis that provides insights into how the MFSP responds to variations in pairs of variables, which are visually represented in contour plots.

3. Multivariate Uncertainty Analysis that allows users to modify a set of input variables simultaneously during a simulation.

Dependencies:
AutoAspen has been developed and tested on Python 3.8 with the following dependencies: numpy 1.20.2, pandas 1.0.5, scipy 1.5.0, scikit-learn 0.23.1, matplotlib 3.2.2, seaborn 0.12.0, pillow 7.2.0, xlrd 1.2.0, pythoncom, and win32com.

Usage:
1. `generate_dataset_template.py`: This script creates a dataset template for training by generating random values for input variables based on specified distributions defined in the `var_info` file. Supported distributions include: normal, alpha, beta, gamma, triangular, pareto, and bernoulli. A sample `var_info` file is provided [here](https://github.com/Chaowu88/autoaspen/blob/main/ATJ_pathway/var_info.xlsx).

2. `generate_dataset.py`: It generates a training dataset by invoking the Aspen model and the .xslm calculator. This step automates the traditional stochastic TEA, which usually involves iterative calls to Aspen Plus software.

3. `train_regression_model.py`: This script trains and fine-tunes four machine learning models. Users can select the best-performing model for Monte Carlo simulation. Supported machine learning models include: polynomial ridge regression, linear SVM, random forest, and gradient tree boosting.

4. `predict_and_simulate.py`: This script predicts the MFSP using a trained regression model and conducts univariate, bivariate, or multivariate uncertainty analysis. The distributions and baseline values of input variables are specified in the [`config`](https://github.com/Chaowu88/autoaspen/blob/main/ATJ_pathway/config.xlsx) file.
