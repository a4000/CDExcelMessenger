<img src="cimcb_logo.png" alt="drawing" width="400"/>

# CDExcelMessenger
CDExcelMessenger.py has functions that allow passing data between an Excel file and a Compound Discoverer (CD) results file. There is also a function to convert data exported from CD into the TidyData format. CDExcelNotebook.ipynb is a Jupyter Notebook that is designed to make it easy to use the functions in CDExcelMessenger.py

## Steps to use

1. Download Anaconda from the [Anaconda website](https://anaconda.com/products/distribution)
2. Run Anaconda Prompt
3. Run these lines in Anaconda Prompt one line at a time
```console
conda install git
git clone https://github.com/a4000/CDExcelMessenger.git
cd CDExcelMessenger
conda env create -f environment.yml
conda activate CDExcelMessenger
```
4. Copy your CD results file and Excel file to the folder of the CDExcelMessenger environment. You can find the path to the environment folder in the Anaconda Prompt window in the format '(CDExcelMessenger) path to environment'
5. Then run this line in Anaconda Prompt and choose CDExcelNotebook.ipynb
```console
jupyter notebook
```

### Authors
- [Adam Bennett](https://github.com/a4000)
- [David Broadhurst](https://scholar.google.ca/citations?user=M3_zZwUAAAAJ&hl=en)
