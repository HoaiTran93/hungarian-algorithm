from scipy.optimize import linear_sum_assignment
import numpy as np
import xlrd #reading xls/xlsx files

class Hungarian():
    def __init__(self,pathData):
        self.pathData = pathData
        # To open Workbook
        self.wb = xlrd.open_workbook(self.pathData)
        self.matrixPrioritySheet = self.wb.sheet_by_index(0)
    
    def execute(self):
        self.getMatrixPriority()
        self.row_ind, self.col_ind = linear_sum_assignment(self.matrixPriority)
        self.sum_matrix = self.matrixPriority[self.row_ind, self.col_ind].sum()

    def getMatrixPriority(self):
        self.matrixPrio = []
        for id_row in range(1, self.matrixPrioritySheet.nrows):
            teacherPrio = []
            for id_col in range(1, self.matrixPrioritySheet.ncols):
                teacherPrio.append(self.matrixPrioritySheet.cell_value(id_row, id_col))
            self.matrixPrio.append(teacherPrio)  
        self.matrixPriority = np.array(self.matrixPrio)

    def toString(self):
        self.columns = []
        for i in range(1, self.matrixPrioritySheet.ncols):
            self.columns.append(self.matrixPrioritySheet.cell_value(0, i))
        
        self.rows = []
        for i in range(1, self.matrixPrioritySheet.nrows):
            self.rows.append(self.matrixPrioritySheet.cell_value(i, 0))
        
        desc = '#' * 80 + '\n'
        for index in range(len(self.row_ind)):
            desc += 'TeacherID: {:s} <===> ClassID: {:s} <===> Cost: {:.1f}'.format(self.rows[index], self.columns[index],
                                                                                    self.matrixPriority[int(self.row_ind[index]),int(self.col_ind[index])])
            desc += '\n'
        desc += '#' * 80 + '\n'
        desc += 'Sum priority: {:.1f}'.format(self.sum_matrix)
        return desc






