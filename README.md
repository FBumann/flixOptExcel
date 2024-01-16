# README
## Purpose of this Package

## General Structure
This package is an Extension of the flixOpt-package from GitHub. This package gets installed with instalation of this package
1. This package consist of 2 parts:
   1. Initiation: For creating a Optimzation Model based on input via excel
   2. Evaluation: For Processing the results for efficient Analysation of the results

## Useage
1. Create a new Python project in your IDE (PyCharm, Spyder, ...) and activate it
2. Install this package via pip in to your environment: `pip install git+https://github.com/FBumann/flixOptExcel.git`
3. Copy the "main.py" file from this package into your project.
4. Make a local copy of the Template_DataInput.xlsx and save it somewhere on your Computer (for. ex. Desktop)
5. Copy the path of the new file into the "main.py" file
6. **Edit the Excel-file to initialize your Model.**
   1. Specify the path, where the results should be saved to.
   2. Specify CO2-Limits
   3. Specify Costs for Electricity, Gas, H2, ...
   4. Specify the Heat Demand
   5. If needed, specify the Temperature of the Heating Network, the surrounding Air and other Heat sources for Heat Pumps
   6. Specify all existing and optional Heat-Generators
      1. KWK
      2. Kessel
      3. EHK
      4. Wärmepumpen
      5. Abwärmequellen
      6. ...
7. Run the main.py file
8. Analyse the results of your Model. It's saved under the path you specified in the input-excel-flie

## Update
To update the repository to the newest version, simply do a clean uninstall and reinstall:
`pip uninstall flixOptExcel `
`pip install --upgrade git+https://github.com/FBumann/flixOptExcel.git@main`