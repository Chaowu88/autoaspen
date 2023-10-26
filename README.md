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
