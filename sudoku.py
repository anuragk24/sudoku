import numpy as np
import pandas as pd
import shutil as sh

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


inp = pd.read_excel('given.xlsx')
given = inp.to_numpy()
a = sudoku(given)
a.solve_puzzle()
a.show()

org = r'given.xlsx'
out = r'result.xlsx'
sh.copy(org, out)


