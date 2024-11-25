import pymysql

#database connection
# connection = pymysql.connect(host="localhost", user="root", passwd="", database="Student_details")
# cursor = connection.cursor()
# # Query for creating table
# StuTableSql = """CREATE TABLE Stu_Registration(
# Date CHAR(100),
# Name  CHAR(100),
# Mobile_Number CHAR(100),
# Alternate_Number CHAR(100),
# Email_Id  CHAR(100),
# Address  CHAR(100),
# Course_Interested  CHAR(100),
# Batch_Preferred CHAR(100),
# How_You_Came_To_Know_us CHAR(100),
# Experience_Fresher CHAR(100),
# Contact_Person_From_Besant  CHAR(100),
# Counselor  CHAR(100),
# Fees  CHAR(100))"""
# cursor.execute(StuTableSql)
# connection.close()

connection = pymysql.connect(host="localhost", user="root", passwd="", database="Student_details")
cursor = connection.cursor()
# Query for creating table
StuenqTableSql = """CREATE TABLE Stu_Enquiry(
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
Fees  CHAR(100))"""
cursor.execute(StuenqTableSql)
connection.close()