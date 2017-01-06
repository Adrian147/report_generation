import os.path
import json
import xlstojson
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import datetime

##TODO

##Regex Check on the date string

##need of normalization column beween excel file and the template generated.
##Write check function to see diffence in normalised headers.+

#are dictionaries Call by Value or references

##if a course is not taken.



def main():
    if prelim_check():
        exit()
        
    filename = raw_input("Enter the path to the filename -> ") or "marks.xls"
    filejson = xlstojson.main(filename)
    
    if(filejson == ""):
        print "data extraction process failed.\n"
        exit()
    
    with open(filejson, "r") as file_data:
        json_string = file_data.read()
        json_data = json.loads(json_string)
        
        #verify the grade and attendance column.
        if verify_data_columns(json_data):
            exit()
        
        #taking input from user for meta details
        (date, semester, duration) = date_sem_entry()
        
        #Printing out each student data.
        for student_data in generate_student(json_data):
            #print "\n"
            student_data = insert_meta(student_data, date, semester, duration)
            file_print(student_data)

#function to insert meta details for date, semester and duration.
def insert_meta(student_data, date, semester, duration):
    student_data["date"] = date
    student_data["semester"] = "%s, %s"%(semester, duration)
    
    
    return student_data
        
#generator function for extraction data for each student.
#need to start on a marks detail extractor.
def generate_student(json_data):
    courses = gen_course_lookup(json_data["Courses"])
    
    #print courses
    for i in range(0, len(json_data["Grade"])):
        student = student_detail(json_data["Grade"][i])    
        student["marks"] = student_marks(json_data, i, courses)
        
        yield student
        
#Generates a dictionary for the course details.
def gen_course_lookup(raw_courses):
    courses = {}
    
    for item in raw_courses:
        courses[item["Course Code"]] = item
            
    return courses

    
#populates the marks and attendance for all courses
def student_marks(json_data, i, courses):
    marks = []
    
    for course_code in json_data["Grade"][i].keys():
        if course_code in courses:
            course_data = courses[course_code]
            course_data["Grade"] = json_data["Grade"][i][course_code]
            course_data["Attendance"] = json_data["Attendance"][i][course_code]
            
            marks.append(course_data)
    
    return marks
    

#student details
def student_detail(row):
    student = {}
    student['name'] = row["Name"]
    student['roll_no'] = row["Roll no"]
    student['branch'] = row["Branch"]
    student['credit'] = row["Total Credit"]
    student['gpa'] = row["GPA"]
    student['cgpa'] = row["CGPA"]
    
    return student   
        
        
#entering the date and semester details.
##incomplete
def date_sem_entry():
    print "\n###\n\n"
    today = datetime.datetime.now().strftime("%d/%m/%Y")
    date = raw_input("Date on pdf[%s]: " %today) or today
    
    print "\n###\n\n"
    semester = raw_input("Semester No: ") or "I"
    
    print "\n###\n\n"
    month_dur = raw_input("Months Duration? (Spring or Fall): ") or "Spring"
    
    if(month_dur == "Spring"):
        duration = "January - April"
    else:
        duration = "August - November"
    
    return (date, semester, duration)


#Each row of the Grades section needs to tally with the attendance per roll no.
def verify_data_columns(json_data):
    print "no of students : %d" %len(json_data["Grade"])
    
    for i in range(0, len(json_data["Grade"])):
        if json_data["Grade"][i]["Roll no"] != \
            json_data["Attendance"][i]["Roll no"]:
            print "Grades do not tally at row %d" %(i + 1)
            return 1
    
    print "json verified"
    return 0
    
       
#Prelim Check
def prelim_check():
    if not os.path.isfile("template.html"):
        print "pdf template file not present"
        return 1
    
    if not os.path.isdir("generated"):
        os.mkdir("generated")
        if not os.path.isdir("generated"):
            print "directory creation failed!"
            return 1
    return 0
    
           
#function to print individual reports.
def file_print(student_data):
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template("template_new.html")
    html_out = template.render(student_data)
    #file1 = open("%d.html" %student_data["roll_no"], "w")
    #file1.write(html_out)
    #file1.close()
    HTML(string=html_out).write_pdf("%d.pdf" %student_data["roll_no"],\
        stylesheets=["stylesheet.css"])

    
if __name__ == '__main__':
    main()

