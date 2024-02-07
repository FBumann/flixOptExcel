from flixOpt_excel.Initiation.Modules import ExcelModel



# Specify paths and solver_name
excel_file_path = '/Absolute/path/to/the/excel/file/with/the/Inputdata.xlsx'


solver_name = "cbc"         #  open source solver
#solver_name = "gurobi"     # commercial solver (Free academic licences). Much faster for large Models and storages


def main():
    excel_model = ExcelModel(excel_file_path=excel_file_path)
    excel_model.solve_model(solver_name=solver_name, gap_frac=0.01, timelimit=3600)
    excel_model.visualize_results()
    # calculation_results_for_further_inspection = excel_model.load_results()










if __name__ == '__main__':
    main()

# optional: change values for gap_frac and timelimit
'''
:param gap_frac:
    0...1 ; gap to relaxed solution. Higher values for faster solving. 0...1
:param timelimit:
    timelimit in seconds. After this time limit is exceeded, the solution process is stopped and the best yet found result is used. 
    If no result is found yet, There Process is aborted
'''