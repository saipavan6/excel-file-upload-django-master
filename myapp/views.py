from django.shortcuts import render
import openpyxl
from . import dal

def CheckExceldata(request):
    if "GET" == request.method:
        return render(request, 'myapp/ExelReadOnly.html', {})
    else:
        excel_file = request.FILES["excel_file"]
        wb = openpyxl.load_workbook(excel_file)
        worksheet = wb["4th new "]
        row_data2 = list()
        for i,row in enumerate(worksheet.iter_rows()):
            # if i == 0:
            #     continue
            t = tuple()
            for cell in row:
                t = t + (cell.value,)
            row_data2.append(t)

        return render(request, 'myapp/ExelReadOnly.html', {"excel_data": row_data2})




def index(request):
    if "GET" == request.method:
        return render(request, 'myapp/index.html', {})
    else:
        excel_file = request.FILES["excel_file"]

        # you may put validations here to check extension or file size

        wb = openpyxl.load_workbook(excel_file)

        # getting all sheets
        #sheets = wb.sheetnames
        #print(sheets)

        # getting a particular sheet
        worksheet = wb["GRADE III"]

        # getting active sheet
        #active_sheet = wb.active
        #print(active_sheet)

        # reading a cell
        #print(worksheet["A1"].value)

        #excel_data = list()
        # iterating over the rows and
        # getting value from each cell in row
        row_data2 = list()
        for i,row in enumerate(worksheet.iter_rows()):
            if i == 0:
                continue
            #row_data = list()
            t = tuple()
            for cell in row:
                t = t + (cell.value,)

            row_data2.append(t)
        data = dal.dbStudentList(row_data2)
        return render(request, 'myapp/index.html', {"excel_data": row_data2})









        #print(row_data2)

        # import pyodbc
        # server = 'SUVARNAMUKHI\SQLEXPRESS'
        # database_name = 'CampusMgmtNew'
        # user_name = 'sa'
        # password = 'sadguru'
        # conn = pyodbc.connect(
        #             'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database_name + ';UID=' + user_name + ';PWD=' + password)
        #
        # cursor = conn.cursor()
        # for i in row_data2:
        #     print(i)
        #     if i[0] == None:
        #         AdmissionNo = 'null'
        #     else:
        #         AdmissionNo = f"'{i[0]}'"
        #
        #     if i[1] == None:
        #         HallTicketNo = 'null'
        #     else:
        #         HallTicketNo = f"'{i[1]}'"
        #
        #     if i[5] == None:
        #         dob = 'null'
        #     else:
        #         dob = f"'{i[5]}'"
        #     if i[6] == None:
        #         FatherName = 'null'
        #     else:
        #         FatherName = f"'{i[6]}'"
        #
        #     if i[7] == None:
        #         MobileNo = 'null'
        #     else:
        #         MobileNo = f"'{i[7]}'"
        #     if i[8] == None:
        #         MotherName = 'null'
        #     else:
        #         MotherName = f"'{i[8]}'"
        #
        #     if i[9] == None:
        #         Email = 'null'
        #     else:
        #         Email = f"'{i[9]}'"
        #
        #     if i[2] == None:
        #         conn.commit()
        #         print("Data Successfully Inserted")
        #         conn.close()
        #         return render(request, 'myapp/index.html', {"excel_data":row_data2})
        #
        #     SQLCommand = f"""DECLARE @Regid int,@SuperId int,@BranchId int,@Academicid int,@Sectionsid int,@Standardsid int
        #     SET @SuperId = 20019
        #     select @BranchId = id from [CONFIG].[Branches] where SuperId = @SuperId and Orgtype =1
        #     select @Academicid = id from [CONFIG].[AcademicYears] where SuperId = @SuperId and CurrentActive = 1 and IsActive = 1
        #     select @Sectionsid = id from [CONFIG].[Sections] where SuperId = @SuperId and BranchId = @BranchId
        #     select @Standardsid = s.id from [CONFIG].[Programs] p
        #         inner join [CONFIG].[Courses] c on c.Programid = P.ID
        #         INNER JOIN [CONFIG].[Standards] S on s.CourseId = c.Id
        #         where p.SuperId = 20019 and s.Name = '{i[3]}' and
        #         p.BranchId = @BranchId
        #
        #     insert into [Campus].Registrations(Superid,AdmissionNo,HallTicketNo,FirstName,Gender,Dob,BranchId)
        #     values(@SuperId,{AdmissionNo},{HallTicketNo},'{i[2]}',{i[4]},{dob},@BranchId)
        #
        #     SET @Regid = SCOPE_IDENTITY()
        #
        #     insert into [Campus].[StudentTransactions](Academicid,Regid,Standardid,sectionid,ispromoted)
        #     values(@Academicid,@Regid,@Standardsid,@Sectionsid,0)
        #
        #     insert into [Campus].[GuardianDetails](Regid,FatherName,MobileNo,MotherName,Email )
        #     values(@Regid,{FatherName},{MobileNo},{MotherName},{Email})"""
        #
        #     # cursor = conn.cursor()
        #     # quary_execution = cursor.execute(SQLCommand)
        #
        #     # Processing Query
        #     cursor.execute(SQLCommand)
        #
        # conn.commit()
        # print("Data Successfully Inserted")
        # conn.close()
        #
        #
        # return render(request, 'myapp/index.html', {"excel_data":row_data2})









