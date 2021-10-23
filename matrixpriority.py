import xlrd #reading xls/xlsx files
import xlwt #writing xls files

class MatrixPriority():
    def __init__(self,path):
        self.path = path
        # To open Workbook
        self.wbo = xlrd.open_workbook(self.path)
        self.CoursesSheet = self.wbo.sheet_by_index(0)
        self.TeacherSheet = self.wbo.sheet_by_index(1)
        self.RegistrationSheet = self.wbo.sheet_by_index(2)
        #to write Workbook
        self.wbw = xlwt.Workbook()
        self.priorityMatrixSheet = self.wbw.add_sheet('Priority_Matrix', cell_overwrite_ok=True)
        self.output = './output/matrix_priority.xls'

    def generate(self):
        self.get_courses()
        self.get_courses_classes()
        self.priorityMatrixSheet.write(0, 0, "Teacher ID \ Class ID")

        for index_column in range(len(self.listCouseID)):
            self.priorityMatrixSheet.write(0, index_column + 1, self.listCouseID[index_column])

        for index_row in range(1, self.TeacherSheet.nrows):
            self.priorityMatrixSheet.write(index_row, 0, self.TeacherSheet.cell_value(index_row, 0))

        for id_row in range(1, self.TeacherSheet.nrows):
            for id_column in range(len(self.listCouseID)):
                classid = self.listCouseID[id_column]
                classindex = int(classid[1:classid.find("-")])

                for col in range(1, self.RegistrationSheet.ncols):
                    if self.TeacherSheet.cell_value(id_row, 0) == self.RegistrationSheet.cell_value(0, col):
                        if self.RegistrationSheet.cell_value(classindex, col) == '':
                            self.priorityMatrixSheet.write(id_row, id_column + 1, 999)
                        else:
                            self.priorityMatrixSheet.write(id_row, id_column + 1, self.RegistrationSheet.cell_value(classindex, col))
        
        self.wbw.save(self.output)
        return self.output
    
    def get_courses(self):
        self.courseDict = {}
        for i in range(1,self.CoursesSheet.nrows):
            classes = self.CoursesSheet.cell_value(i, 3)
            if classes != '':
                CourseID = "C" + str(i)
                numClasses = {CourseID:classes}
                self.courseDict.update(numClasses)
    
    def get_courses_classes(self):
        self.listCouseID = []
        for item in self.courseDict.items():
            for num in range(int(item[1])):
                nameCourseID = item[0] + "-" + str(num + 1)
                self.listCouseID.append(nameCourseID)
