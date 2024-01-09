import pandas as pd
from bs4 import BeautifulSoup
import os
import re
import openpyxl
import codecs

# Loading the CSV data into a DataFrame
df = pd.read_csv('input_email_details.csv')

# Defining the path to the "mail" folder which contains all the html files
mail_folder_path = 'mail'  

# Function to extract the subject from email content
def extract_subject(email_content):
    # Defining a regular expression pattern to match the subject line
    subject_pattern = re.compile(r"Subject:\s*([^\n]+).*?Mail Archive", re.DOTALL)
    match = subject_pattern.search(email_content)

    if match:
        subject = match.group(1)
        return subject.strip()
    else:
        return None  # Returning None if no subject is found

# Function to extract email content from HTML file
def extract_email_content(html_filename):
    try:
        if isinstance(html_filename, str):
            # Extracting the file name from the path
            filename = os.path.basename(html_filename)

            # Constructing the full path to the HTML file
            full_path = os.path.join(mail_folder_path, filename)
            print('Extracting File :', full_path)
            # Checking if the file exists before attempting to open it
            if os.path.exists(full_path):
                with open(full_path, 'r', encoding='utf-8') as file:
                    email_html = file.read()
                soup = BeautifulSoup(email_html, 'html.parser')
                email_content = soup.get_text()
                # Extracting and returning the email subject
                email_subject = extract_subject(email_content)
                print("Email subject of the file :", email_subject)
                return email_content.strip(), email_subject
            else:
                # Handling the case when the file doesn't exist
                return "File does not exist: " + filename, None
        else:
            # Handling the case when the filename is not a string
            return "Invalid format: " + str(html_filename), None
    except Exception as e:
        # Handling exceptions and printing error details
        print(f"Error processing {html_filename}: {str(e)}")
        return "Error", None

# Applying the function to update the "Email Body" and "Email subject" columns, but skipping NaN or empty values
# Applying the function to create a DataFrame with two columns
result_df = df['mail_filename'].apply(
    lambda x: extract_email_content(x) if pd.notna(x) and x.strip() else (None, None)
)
result_df = pd.DataFrame(result_df.tolist(), columns=['Email Body', 'Email subject'])

def clean_text(text):
    if isinstance(text, str):
        # Defining a regular expression pattern to match illegal characters
        # Excluding both 'b' and 'B' from the characters to be replaced
        illegal_char_pattern = re.compile(r'[|　bB]')  # Removing 'b' and 'B' from the pattern
        cleaned_text = re.sub(illegal_char_pattern, ' ', text)  # Replacing with spaces

        # Preserving line breaks and multiple spaces
        cleaned_text = re.sub(r'\n+', '\n', cleaned_text)
        cleaned_text = re.sub(r' +', ' ', cleaned_text)

        return cleaned_text
    return text

# Updating the original DataFrame with the new data
df[['mail_body', 'mail_subject']] = result_df

df['mail_body'] = df['mail_body'].apply(clean_text)
df['mail_subject'] = df['mail_subject'].apply(clean_text)

# Saving the updated DataFrame back to the CSV file with proper encoding
file_path = 'output.xlsx'
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)

# Manually setting the encoding of the XLSX file to UTF-8
with open(file_path, 'rb') as f:
    data = f.read()
with open(file_path, 'wb') as f:
    f.write(data)
