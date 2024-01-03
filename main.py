from flixOpt_excel.Initiation.Excel_Modell import run_excel_model





# Specify paths and solver_name
excel_file_path = '/Users/felix/Documents/Dokumente - eigene/Arbeit_SE_Trainee/DataInput mehrere Jahre.xlsx'


solver_name = "cbc"         #  open source solver
#solver_name = "gurobi"     # commercial solver (Free academic licences). Much faster for large Models and storages










if __name__ == '__main__':
    run_excel_model(excel_file_path, solver_name, gap_frac=0.001, timelimit= 3600)

# optional: change values for gap_frac and timelimit
'''
:param gap_frac:
    0...1 ; gap to relaxed solution. Higher values for faster solving. 0...1
:param timelimit:
    timelimit in seconds. After this time limit is exceeded, the solution process is stopped
'''