"""
Different PDFs to make


Main Groups 
- Just PNG analysts
- All PNG members
- Everyone 

Sub Groups
- Software / Data
- Quant
- Fundamentals

Sub Sub Groups
- School
- Sponsor



Groupings
PNG Analysts | All Schools | Group -> 3
PNG Members | All Schools | Group  -> 3
Teaser (All | Top 30) -> 1
"""



import pandas as pd
import os
from pypdf import PdfMerger



# Below are the inputs for excel file with information and directory path containing hte resumes

excel_file = 'resumes.xlsx'
directory = 'resumesfolder'

main_groups = ["analysts", "allpng", "everyone"]
schools = ['University of Chicago', 'University of Pennsylvania', 'New York University']
groups = ['software', 'quant',' fundamentals']

# subcategories = {
#     "everyone": [], 
#     "memberswe": [], 
#     "memberquant": [], 
#     "memberfund": [], 
#     "analystswe": [],
#     "analystquant": [],
#     "analystfund": [],
#     "upennswe": [],
#     "upennquant": [],
#     "upennfund": [],
#     "uchicagoswe": [],
#     "uchicagoquant": [],
#     "uchicagofund": [],
#     "nyuquant": [],
#     "nyuswe":[]
                 
#                  }

subcategories = {
    "everyoneswe": [], 
    "everyonequant": [],
    "everyonefund": [],
    "memberswe": [], 
    "memberquant": [], 
    "memberfund": [], 
    "analystswe": [],
    "analystquant": [],
    "analystfund": [],
    "allpngswe": [],
    "allpngquant": [],
    "allpngfund": [],
                 }

# Code for getting and parsing info
df = pd.read_excel(excel_file)
df.set_index('Timestamp', inplace=True)


for i in range(len(df)):
    # Get key variables 
    first_name = df['First Name'][i]
    middle_initial = df['Middle Initial (put NA if you don\'t have one)'][i]
    last_name = df['Last Name'][i]
    school = df['What school do you attend?'][i]
    school_email = df['School Email'][i]
    grad_year = df['Graduation Year'][i]
    interest_group = df['Interest group'][i]
    interest_group_list = interest_group.split(',')
    relation = df['What is your relation to PNG?'][i]

    new_school = ""

    if school == "University of Chicago":
        new_school = "uchicago"
    elif school == "University of Pennsylvania":
        new_school = "upenn"
    else:
        new_school  = "nyu"

    #format in resume format

    resume_format = ""
    if middle_initial.upper() == "NA":
        resume_format = first_name.lower() + "." + last_name.lower() + "_" + new_school.lower() + "_" + str(int(grad_year)) + "_" + "resume"
        print(resume_format)
    else:
         resume_format = first_name.lower() + "." + middle_initial.lower() + "." + last_name.lower() + "_" + new_school.lower() + "_" + str(int(grad_year)) + "_" + "resume"
         print(resume_format)
    print(relation)
    print(interest_group)

    # PNG Analyst
    if relation == "Investment Analyst / Quantitative Analyst":
        if "Fundamentals" in interest_group:
            subcategories["analystfund"].append(resume_format)
            subcategories["allpngfund"].append(resume_format)
        if "Quantitative Trading/Research" in interest_group:
            subcategories["analystquant"].append(resume_format)
            subcategories["allpngquant"].append(resume_format)
        if "Software Development / Data Science" in interest_group:
            subcategories["analystswe"].append(resume_format)
            subcategories["allpngswe"].append(resume_format)

    # PNG Members too
    if relation == "Education Group Member":
        if "Fundamentals" in interest_group:
            subcategories["memberfund"].append(resume_format)
            subcategories["allpngfund"].append(resume_format)
        if "Quantitative Trading/Research" in interest_group:
            subcategories["memberquant"].append(resume_format)
            subcategories["allpngquant"].append(resume_format)
        if "Software Development / Data Science" in interest_group:
            subcategories["memberswe"].append(resume_format)
            subcategories["allpngswe"].append(resume_format)


    # Everyone
    subcategories["everyonefund"].append(resume_format)
    subcategories["everyoneswe"].append(resume_format)
    subcategories["everyonequant"].append(resume_format)



# Output PDFs
output_directory_path = "resumesoutput"
input_directory_path = 'resumesfolder'
for category in subcategories:
    merger = PdfMerger()
    for resume in subcategories[category]:
        try:
            curr = input_directory_path + "/" + resume + ".pdf"
            merger.append(curr)
        except:
            continue


    merger.write(output_directory_path + "/" + category + ".pdf")
    merger.close()