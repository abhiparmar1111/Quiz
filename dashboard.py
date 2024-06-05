from tkinter import *
from openpyxl import *
import signup
import login

wb=load_workbook("C:\\Users\\Abhayraj sinh parmar\\python_files\\student_info.xlsx")

sheet=wb.active

if __name__ == "__main__":
    root=Tk()
