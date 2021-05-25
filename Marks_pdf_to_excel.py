# Using PyPdf2 to work with the pdf 

import PyPDF2
import pandas as pd

pdfFileObj = open('11860_1.pdf','rb')  

# creating a reader object so that various methods can be used over it.
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
  
# printing 1st page to get the idea about type of data in pdf
first_page = pdfReader.getPage(1)
print(first_page.extractText())

# combining the text on all pages at same place so it can be processed.

for page_no in range(pdfReader.getNumPages()):
  pageObj_single = pdfReader.getPage(page_no)
  pageone_text_single = pageObj_single.extractText()
  all_in_one = all_in_one + pageone_text_single

print(pdfReader.getNumPages())

# Have to give the the start and end seat number as the data is about the results 
# As needed it should convert the all years of same of that department so making provision for other sem/year

input_start_seat_no = 20830
input_end_seat_no = 20926

marks_of_year = 1

# have to pass the subjects list so provided for 1st year and making provision for 2nd and 3rd
if marks_of_year == 1:
    list_of_subjects = [' 101 ',' 102 ',' 103 ',' 104 ',' 105 ',' 106 ',' 107 ',' 108 ',' 109 ',' 110 ',' 111 ',' 191 ',' 192 ',' 201 ',' 202 ',' 203 ',' 204 ',' 205 ',' 206 ',' 207 ',' 208 ',' 209 ',' 210 ',' 211 ',' 291 ',' 292 ']
elif marks_of_year == 2:
    pass
elif marks_of_year == 3:
    pass

# Making the list of students seat no and converting the list elements to str as needed later
list_of_seat_no = list(range(input_start_seat_no,input_end_seat_no+1))
list_of_seat_no_str = []
for item_losn in list_of_seat_no:
    list_of_seat_no_str.append(str(item_losn))

# function to extract the name 
def name_extraction(single_seat_text,name_start_index):              
        single_name = ''
        alpha_one = single_seat_text[name_start_index]
        while alpha_one != ' ':  
            if alpha_one ==' ':
                break
            single_name = single_name + alpha_one
            name_start_index = name_start_index+1
            alpha_one = single_seat_text[name_start_index]
        return single_name,name_start_index+1
# Main code for extaction

all_students_marks = [] # final list to use 
count = 1

for seat_no_index in range(0,len(list_of_seat_no_str)):     #looping through the list of seat numbers
    start_seat_no = list_of_seat_no_str[seat_no_index]      # taking the seat number index
    if start_seat_no== list_of_seat_no_str[-1]:             # for handing the condition of last seat number of list
        upto_seat_no = ''
    else:
        upto_seat_no = list_of_seat_no_str[seat_no_index+1]
    
    
    start_seat_no_index = all_in_one.find(start_seat_no)
    if upto_seat_no=='':                                    # for handling the last seat number index i.e upto where to look
        end_seat_no_index = None
    else:
        end_seat_no_index = all_in_one.find(upto_seat_no)

# needed some looping cases above below code will be same

    single_seat_text = all_in_one[start_seat_no_index:end_seat_no_index]     # taking the block of marks of specific student

    list_of_individual = []    
    
    seat_no_individual = ''
    seat_no_individual = single_seat_text[:5]
    list_of_individual.append([seat_no_individual])
    
    full_name_individual = ''

    name_start_index = 7
    
    for i in range(3):                                        # looping to get the full name
        name,name_start_index = name_extraction(single_seat_text,name_start_index)
        full_name_individual += ' '+name
    striped_full_name = full_name_individual.strip()
    list_of_individual.append([striped_full_name])
    
    for i in range(0,len(list_of_subjects)):               # looping through the subject list for collectiong the marks
        sub_code_start_index = single_seat_text.find(list_of_subjects[i])
        if sub_code_start_index==(single_seat_text.find(list_of_subjects[-1])):
            sub_code_end_index = sub_code_start_index+29
        else:
            sub_code_end_index = single_seat_text.find(list_of_subjects[i+1])
        sub_marks_string = single_seat_text[(sub_code_start_index+6):sub_code_end_index]
        list_of_marks = sub_marks_string.split(' ')
        cleaned_marks = []
        for list_item in list_of_marks:
            if list_item.strip() and  not('---' in list_item) and not('!' in list_item):
                cleaned_marks.append(list_item)
        list_of_individual.append(cleaned_marks)
    all_students_marks.append(list_of_individual)
    count += 1
        
# Preparing the list to be added to the df
for individual_student in all_students_marks:
    individual_student_prepared = []
    for individual_student_list_item in individual_student:
        if len(individual_student_list_item) > 1 and len(individual_student_list_item)<=5 :
            if '*' not in individual_student_list_item and len(individual_student_list_item)==4:
                individual_student_list_item.insert(1,' ')
                individual_student_list_item.insert(1,' ')
            if '*' not in individual_student_list_item and len(individual_student_list_item)==5:
                individual_student_list_item.insert(2,' ')
            if '*' in individual_student_list_item and len(individual_student_list_item)== 5:
                individual_student_list_item.insert(1,' ')
        individual_student_prepared.append(individual_student_list_item)

 # creating empty data frame with name column

marks_df = pd.DataFrame(columns=['seat_no','Name_of_Student']) 

# list of dataframe columns   ' 101_INT ' ,' 101_EXT ','101_CUR_SUB','101_TOTAL','101_GRADE','101_CRD'
list_of_columns = []
for each_sub in list_of_subjects:
    list_sub_addons =['_INT','_EXT','_CUR_SUB','_TOTAL','_GRADE','_CRD']
    for list_sub_addons_item in list_sub_addons:
        marks_type = each_sub + list_sub_addons_item
        list_of_columns.append(marks_type)
        
# Adding the columns to the dataframe with empty

for sub in list_of_columns: 
    marks_df[sub] = ''

# To handle some Conditions 
i=0
for individual_student in all_students_marks:
    individual_student_single_prepared = []
    for individual_student_item in individual_student:
        for individual_student_item_1 in individual_student_item:
            individual_student_single_prepared.append(individual_student_item_1)
    if i ==10:
        print(individual_student_single_prepared)
    marks_df.loc[len(marks_df)] = individual_student_single_prepared
    i = i + 1
    
 print(marks_df.info())

# Finally Exporting ths DataFrame to excel

marks_df.to_excel('Students_marks.xlsx')

# pdf and excel both are added to the repository.
