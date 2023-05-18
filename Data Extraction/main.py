import re
import pdfplumber
import openpyxl

def extract_data_from_pdf(path):
    # Initialize an empty dictionary for the colleges
    college_dict = {}

    # Open the PDF file
    with pdfplumber.open(path) as pdf_file:

        # Iterate through the pages of the PDF
        for page in pdf_file.pages:

            # Extract the text from the page
            text = page.extract_text()

            # Use regular expressions to find the college name pattern
            college_pattern = r"\d{4}\s-\s\w+[^\n]*"
            college_match = re.search(college_pattern, text)

            # If a college match is found, use it as the key in the college dictionary
            if college_match:
                college_key = college_match.group().strip()

                # Check if the college key already exists in the dictionary
                if college_key not in college_dict:
                    # If the college key does not exist, add it to the dictionary with an empty list for courses and grades
                    college_dict[college_key] = {'grade':[], 'course':[]}

                # Use regular expressions to find all the course patterns
                course_pattern = r"\d{9}\s-\s\w+[^\n]*"
                course_matches = re.finditer(course_pattern, text)
                
                # Iterate through each course match
                for course_match in course_matches:
                    course_key = course_match.group().strip()

                    # Added courses to the list of courses in respective colleges
                    college_dict[college_key]['course'].append(course_key)

                # Use pdfplumber to extract tables from the page
                tables = page.extract_tables()

                # If tables are found, check if they are unique and append them to the list of grades in respective colleges
                if tables:
                    for table in tables:
                        college_dict[college_key]['grade'].append(table)

    # Return the dictionary of colleges and courses
    return college_dict



path = "./data/raw data/mht-cet-2022-round-3-cutoff-mh.pdf"
college_dict = extract_data_from_pdf(path)
header = ['College Name', 'Course'] 
indexing = [] 


# Create a new workbook and select the active worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active

# For header row values (header of the table)
for college in college_dict:
    for course, grade in zip(college_dict[college]['course'], college_dict[college]['grade']):
        for cast, rank in zip(grade[0], grade[1]):
           if cast and cast not in indexing:
            header.append('Percentage ' + cast)
            header.append('Rank ' + cast)
            indexing.append(cast)
worksheet.append(header)

# For row values
for college in college_dict:
    for course, grade in zip(college_dict[college]['course'], college_dict[college]['grade']):
        row = ['-'] * (len(header)+2)
        row[0] = college
        row[1] = course
        for cast, rank in zip(grade[0], grade[1]):
            rank = rank.split('(')
            if cast and cast in indexing:
                i = indexing.index(cast)
                row[(2*i)+2] = rank[-1].replace(')', '')
                row[(2*i)+3] = rank[0]
        worksheet.append(row)

workbook.save('./data/processed data/mht-cet-2022-round-3-cutoff-mh.xlsx')
