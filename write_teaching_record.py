#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jul 20 20:24:57 2021

@author: panda
"""
from docx import Document
import openpyxl
import docx2txt

import re
import os

faculty_names= []
doc= Document('Teaching Record Temp.docx')


            #Search column I for name to get row 
                #use row to get columns:
                    #"G" - Subject Course num (take off section Num)
                    # "H" - Course ttile
                    #"K" - enrollment
                    # "L" - response rate (x 100% to get percentage)
                    # "M" - Instructor Average 
                    # "P" - Course Average
                    # "S" - Dept Instructor Avg
                    # "V" - Dept Course Average

class evaluationScores:
    def __init__(self, faculty):
        self.faculty= faculty
        self.all_scores={}
        self.name= faculty.split(',')
        self.last_name= self.name[0]
        
        self.course_id_column=''
        
        
    def save_scores(self, rundown, q):
        self.courses=[]
        for row in rundown["L"]:
            upper_name=self.last_name.upper()
            
            if upper_name in row.value:
                
                r= row.row
                self.course_id= rundown.cell(column=get_quarter_columns(rundown, "Subject Course Section"), row=r).value
                self.title= rundown.cell(column=get_quarter_columns(rundown, "Course Title"), row=r).value
                self.enrollment= rundown.cell(column=get_quarter_columns(rundown, "Enrollment"), row=r).value
                self.response= rundown.cell(column=get_quarter_columns(rundown, "Response Rate"), row=r).value
                self.ins_avg= rundown.cell(column=get_quarter_columns(rundown, "Inst AVG"), row=r).value
                self.crs_avg=rundown.cell(column=get_quarter_columns(rundown, "Crs AVG"), row=r).value
                self.dept_ins_avg=rundown.cell(column=get_quarter_columns(rundown, "Dept Inst AVG"), row=r).value
                self.dept_crs_avg=rundown.cell(column=get_quarter_columns(rundown, "Dept Crs AVG"), row=r).value
                self.courses.append([self.course_id] )
                              #self.title, self.enrollment, self.response, self.ins_avg, 
                              #self.crs_avg, self.dept_ins_avg, self.dept_crs_avg]
            self.all_scores[q]= self.courses
        return self.all_scores
        
def get_quarter_columns(rundown, column_name):

        for column in rundown[1]:
            if re.search(column_name, column.value)!= None:
                return column.column
def get_quarters_years():
    years= ['19','20','21']
    sessions=["W", "S", "F"]
    quarter_list=[]
    for y in years:
        for q in sessions:
            quarter_list.append(y+q)
    return quarter_list
 
def open_rundown_file(yq):
    path="ENGL Evaluations/"+yq+" ENGL/Rundown Reports"
    location= os.path.abspath(path)
    rundown_file=os.path.join(location, yq+"_ENGL_DeptRundown_Instructor_- A_Instruct_- B_TA_Eval.xlsx")
    try:
        rundown_report= openpyxl.load_workbook(rundown_file)
    except:
        return None
    rundown= rundown_report.worksheets[0]
    return rundown



def main():
    #template= docx2txt.process("Teaching Record Temp.docx")
    quarters= get_quarters_years()
    #iterate through each faculty member
    for name in faculty_names:
        faculty=evaluationScores(name)
        
        
        #iterate through each quarter
        for q in quarters:
            #open rundown file per quarter
            rundown=open_rundown_file(q)
            if rundown!=None:
                #faculty.get_quarter_columns(rundown)
                for name in faculty_names:
        
                    faculty.save_scores(rundown,q)
            
        print(faculty.all_scores)
#File Name variations            
#19S_ENGL_DeptRundown - A_Instruct, B_TA_Eval.xlsx            
#19F_ENGL_DeptRundown_Instructor_- A_Instruct_- B_TA_Eval.xlsx   
#20W_ENGL_DeptRundown_Instructor_- A_Instruct_- B_TA_Eval.xlsx   
#20S_ENGL_DeptRundown_Instructor_- A_Instruct_- B_TA_Eval.xlsx  
main()
        #template= re.sub(name, "<NAME>", template)
        # insert each quarter in table and course info if necessary
    
    #doc.save(name+"Teaching Record.docx")
    
file= "WEAVER, E.M." 
last_name="Weaver"
