import pyodbc

server = 'SUVARNAMUKHI\SQLEXPRESS'
database_name = 'CampusMgmt_05'
user_name = 'sa'
password = 'sadguru'
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database_name + ';UID=' + user_name + ';PWD=' + password)

def dbStudentList(row_data2):
    cursor = conn.cursor()
    for i in row_data2:
        print(i)
        if i[0] == None:
            AdmissionNo = f"' '"
        else:
            AdmissionNo = f"'{i[0]}'"

        if i[1] == None:
            HallTicketNo = f"'HallTicketNo'"
        else:
            HallTicketNo = f"'{i[1]}'"

        if i[5] == None:
            dob = 'null'
        else:
            dob = f"'{i[5]}'"
        if i[6] == None:
            FatherName = f"'FatherName'"
        else:
            FatherName = f"'{i[6]}'"

        if i[7] == None:
            MobileNo = 'null'
        else:
            MobileNo = f"'{i[7]}'"
        if i[8] == None:
            MotherName = f"'MotherName'"
        else:
            MotherName = f"'{i[8]}'"

        if i[9] == None:
            Email = f"'xxx@xx.com'"
        else:
            Email = f"'{i[9]}'"
        if i[10] == None:
            PhoneNo = 'null'
        else:
            PhoneNo = f"'{i[10]}'"

        if i[2] == None:
            conn.commit()
            print("Data Successfully Inserted")
            conn.close()
            return row_data2

        SQLCommand = f"""DECLARE @Regid int,@SuperId int,@BranchId int,@Academicid int,@Sectionsid int,@Standardsid int,@StudentTransactionsid int,@GuardianDetailsid int
                SET @SuperId = 20019
                select @BranchId = id from [CONFIG].[Branches] where SuperId = @SuperId and Orgtype =1
                select @Academicid = id from [CONFIG].[AcademicYears] where SuperId = @SuperId and CurrentActive = 1 and IsActive = 1
                select @Sectionsid = id from [CONFIG].[Sections] where SuperId = @SuperId and BranchId = @BranchId AND name = '{i[11]}'
                select @Standardsid = s.id from [CONFIG].[Programs] p
                    inner join [CONFIG].[Courses] c on c.Programid = P.ID
                    INNER JOIN [CONFIG].[Standards] S on s.CourseId = c.Id
                    where p.SuperId = @SuperId and s.Name = '{i[3]}' and 
                    p.BranchId = @BranchId

                insert into [Campus].Registrations(Superid,AdmissionNo,HallTicketNo,FirstName,Gender,Dob,BranchId)  
                values(@SuperId,{AdmissionNo},{HallTicketNo},'{i[2]}',{i[4]},{dob},@BranchId)

                SET @Regid = SCOPE_IDENTITY()
                
                insert into [Campus].[StudentTransactions](Academicid,Regid,Standardid,sectionid,ispromoted) 
                values(@Academicid,@Regid,@Standardsid,@Sectionsid,0)
                
                SET @StudentTransactionsid = SCOPE_IDENTITY()
                
                
                insert into [Campus].[GuardianDetails](Regid,FatherName,MobileNo,MotherName,Email,PhoneNo ) 
                values(@Regid,{FatherName},{MobileNo},{MotherName},{Email},{PhoneNo})
        
                SET @GuardianDetailsid = SCOPE_IDENTITY()
                
                
                insert into StudentsIds(Academicid,Regid,StudentTransactionsid,GuardianDetailsid,Standardid,sectionid,FirstName,AdmissionNo)
                values(@Academicid,@Regid,@StudentTransactionsid,@GuardianDetailsid,@Standardsid,@Sectionsid,'{i[2]}',{AdmissionNo} )
                """

        # cursor = conn.cursor()
        # quary_execution = cursor.execute(SQLCommand)

        # Processing Query
        cursor.execute(SQLCommand)
        #print("1 record inserted, ID:", cursor.lastrowid)

    conn.commit()
    print("Data Successfully Inserted")
    conn.close()

    return row_data2