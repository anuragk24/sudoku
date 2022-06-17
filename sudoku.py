import numpy as np
import pandas as pd
import shutil as sh
import openpyxl
from openpyxl.styles import PatternFill

class sudoku:
    def __init__(self, maze=np.array(None)):
        self.items = dict()

        for i in range(9):
            for j in range(9):
                self.items[(i, j)] = set(range(1, 10))

        for row, item in enumerate(maze):
            for col, val in enumerate(item):
                if maze[row, col]:
                    self.items[(row, col)] = {int(val)}

    def valcheck(self, row, column):
        maxrow, maxcol = 9, 9
        temp = set()
        if row < maxrow:
            for j in range(maxcol):
                if len(self.items[(row, j)]) == 1:
                    temp = temp.union(self.items[(row, j)])

        if column < maxcol:
            for i in range(maxrow):
                if len(self.items[(i, column)]) == 1:
                    temp = temp.union(self.items[(i, column)])

        for i in range(3*(row//3), 3*(row//3 + 1)):
            for j in range(3*(column//3), 3 * (column//3 + 1)):
                if len(self.items[(i, j)]) == 1:
                    temp = temp.union(self.items[(i, j)])

        self.items[(row, column)] -= temp

    def solve_puzzle(self):
        sum = 0
        for i in range(9):
            for j in range(9):
                if len(self.items[(i, j)]) != 1:
                    self.valcheck(i, j)
                else:
                    sum += 1
                    if sum == 81:
                        return

        self.solve_puzzle()

    def __iter__(self):
        for i in range(9):
            for j in range(9):
                yield self.items[(i, j)]

    def show(self):
        puz = np.empty([9, 9], dtype=set)
        for i in range(9):
            for j in range(9):
                puz[i, j] = self.items[(i, j)]

        print(puz)
        return puz

def set_to_val(set_val):
    set_shape = np.shape(set_val)
    return np.array([[list(set_val[i, j])[0] for j in range(set_shape[1])] for i in range(set_shape[0])])

inp = pd.read_excel('given.xlsx')
given = inp.to_numpy()
a = sudoku(given)
a.solve_puzzle()
result = set_to_val(a.show())

org = r'given.xlsx'
out = r'result.xlsx'
sh.copy(org, out)

df = pd.DataFrame(result, columns=range(1, 10), index=range(1, 10))
writer = pd.ExcelWriter(out, engine='xlsxwriter')
df.to_excel(writer)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Add some cell formats.
format1 = workbook.add_format({'size': 28, 'align': 'center', 'bold':'True'})
# cell_format = workbook.add_format()
# cell_format.set_bg_color('green')
# worksheet.write('B1', None, cell_format)
# Set the column width and format.
worksheet.set_column(1, 10, None, format1)

# Close the Pandas Excel writer and output the Excel file.
writer.save()




