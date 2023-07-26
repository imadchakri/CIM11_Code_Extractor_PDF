# Import required libraries
import pandas as pd
from PyPDF2 import PdfReader
import re

# Read the PDF file and get the total number of pages
reader = PdfReader("Desktop/cim11.pdf")
number_of_pages = len(reader.pages)

# Print the total number of pages in the PDF
print(number_of_pages)

# Text variable will contain all texts from pages 46 to page 833
# Starting from page 833, the format of the code changes, beginning with a letter instead of a number.
text = ""

# Start fetching pages from Chapitre 1 = page 46
for i in range(47, 833):
    page = reader.pages[i]
    text += page.extract_text()

# Initialize a list to store data that will be exported to EXCEL
dataExcel = []

# Regular expressions to extract relevant information from the text
pattern = r'(\d+[A-Z]\d*)(\s*\.\s*\d*[A-Za-z]?)?\s*([^-\s].*)'
second_pattern = r'^\b[A-Z](?:\.\d+)?\b(?!\.\d)'
third_pattern = r'\.\s*([A-Z])'


# Find all matches based on the pattern in the extracted text
matches = re.findall(pattern, text)

# Loop through the matches and process the data
for match in matches:
    code = match[0].strip() + match[1].strip()
    
    # Check if the description meets certain conditions to be considered valid
    if (
        len(match[2].strip()) > 4
        and not match[2].replace(" ", "").startswith("‑")
        and not match[2].startswith("Ce chapitre")
        and not match[2][0].isdigit()
    ):
        match_second = re.findall(second_pattern, match[2])
        match_third = re.findall(third_pattern, match[2].replace(" ", ""))
        
        if match_second:
            second += 1
            if match_third:
                # Handle the case where there is a third part in the code
                description_final = match[2].replace(match_second[0] + "." + match_third[0], "")
                code_final = (code + match_second[0] + "." + match_third[0]).replace(" ", "")
                third += 1
                dataExcel.append({"CIM11": code_final, "Description_fr": description_final})
            else:
                # Handle the case where there is no third part in the code
                out += 1
                description_final = match[2].replace(match_second[0], "")
                code_final = code + match_second[0]
                dataExcel.append({"CIM11": code_final, "Description_fr": description_final})
        else:
            # Handle the case where the code starts with a letter instead of a number
            first += 1
            description_final = match[2]
            code_final = code
            dataExcel.append({"CIM11": code_final, "Description_fr": description_final})

# This part extracts text from page 834 to the end, where the codes start with TWO uppercase letters
text2 = ""
# Start fetching pages from Chapitre 1 = page 46
for i in range(834, 1952):
    page2 = reader.pages[i]
    text2 += page2.extract_text()

# Regular expression to extract codes and descriptions from the second part of the text
pattern2 = r'([A-Z]{2}[A-Z0-9\.]+)\s(.+)'
matches2 = re.findall(pattern2, text2)

# Loop through the matches and process the data
for match in matches2:
    code = match[0]
    description = match[1]
    if code != "MMS" and code != "CHAPITRE":
        if not description.startswith("-") or not description.startswith("‑"):
            dataExcel.append({"CIM11": code, "Description_fr": description})

# Create a pandas DataFrame from the collected data
df = pd.DataFrame(dataExcel)

# Print the DataFrame
print(df)

# Export the DataFrame to an Excel file
output_path = "Desktop/cim11database.xlsx"
df.to_excel(output_path, index=False)

# Clear the dataExcel list and print a success message
dataExcel = []
print("Data exported to", output_path)
