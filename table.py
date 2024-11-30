import pymysql



connection=pymysql.connect(host="localhost",user="root",passwd="",database="student")
cursor=connection.cursor()

RegTableSql = """CREATE TABLE Registration(
Date CHAR(100),
Name  CHAR(100),
Mobile_Number CHAR(100),
Alternate_Number CHAR(100),
Email_Id  CHAR(100),
Address  CHAR(100),
Course_Interested  CHAR(100),
Batch_Preferred CHAR(100),
How_You_Came_To_Know_us CHAR(100),
Experience_Fresher CHAR(100),
Contact_Person_From_Besant  CHAR(100),
Counselor  CHAR(100),
Fees  CHAR(100),
Comments  CHAR(100))"""
cursor.execute(RegTableSql)
connection.close()
print("table created success")


enq=""" create table Enquiry(
Date CHAR(100),
Name  CHAR(100),
Mobile_Number CHAR(100),
Alternate_Number CHAR(100),
Email_Id  CHAR(100),
Address  CHAR(100),
Course_Interested  CHAR(100),
Batch_Preferred CHAR(100),
How_You_Came_To_Know_us CHAR(100),
Experience_Fresher CHAR(100),
Contact_Person_From_Besant  CHAR(100),
Counselor  CHAR(100),
Fees  CHAR(100),
Comments  CHAR(100))"""
cursor.execute(enq)
connection.close()
print("ok")
