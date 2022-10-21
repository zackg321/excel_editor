import openpyxl
from openpyxl import Workbook, load_workbook

import excel_editor
import new_strain
import gen_reports 

def main():
    more_data = True
    while more_data:
        print(
        "Choose what you would like to do:\n")
        user_choice = int(input("""1. Edit an existing sheet.\n2. Create a new strain.\n3. Create strain reports.\n4. Exit\n"""))
        if user_choice == 1:
            excel_editor.run()
        if user_choice == 2:
            new_strain.run()
        if user_choice == 3:
            gen_reports.run()
        if user_choice == 4:
            more_data = False
            break      

if __name__ == "__main__":
    main()

