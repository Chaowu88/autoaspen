# AutoAspen
AutoAspen is a tool for stochastic techno-economic analysis (TEA) using machine learning methods. It is developed to streamline dataset creation, model training, parameter optimization, and Monte Carlo sampling for generating distributions of the minimum selling price (e.g., MFSP) for specified chemical engineering pathway. With trained models, it can perform:
    
__1. Univariate uncertainty analysis__ that assesses how individual variables affect the MFSP while keeping all other variables constant.
    
__2. Bivariate uncertainty analysis__ that provides insights into how the MFSP value responds to variations in paired viables, illustrated in contour plot.

__3. Multivariate uncertainty analysis__ that offers users the flexibility to specify a set of input variables that will be simultaneously modified during a simulation.
## Dependencies
AutoAspen was developed and test underPython 3.8 with the dependencies:

numpy 1.20.2, pandas 1.0.5, scipy 1.5.0, scikit-learn 0.23.1, matplotlib 3.2.2, seaborn 0.12.0, pillow 7.2.0, xlrd 1.2.0, pythoncom and win32com
## Usage
