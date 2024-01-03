from setuptools import setup, find_packages

def read_requirements(file_path):
    with open(file_path, 'r') as file:
        return [line.strip() for line in file if line.strip()]

setup(
    name='flixOptExcel',
    version='0.1.0',
    packages=find_packages(),
    include_package_data=True,  # Include non-code files specified in MANIFEST.in
    package_data={
        'flixOptExcel': ['flixOpt_excel/resources/ExcelTemplates/*.xlsx'],
    },
    #install_requires=read_requirements('requirements.txt')
    install_requires=['matplotlib==3.8.2', 'networkx==3.2.1', 'openpyxl==3.1.2', 'plotly==5.18.0', 'pprintpp==0.4.0', 'git+https://github.com/FBumann/flixOpt.git']
)
