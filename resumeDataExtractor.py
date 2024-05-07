from PyPDF2 import PdfReader
import os, re
# print(type(text))
# print(text)

import win32com.client as win32

def read_doc_file(filepath):
  """
  Reads the text content from a .doc file using pywin32.

  Args:
      filepath (str): Path to the .doc file.

  Returns:
      str: The extracted text content from the file or None if there's an error.
  """
  try:
    # Create a Word application instance (invisible)
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    # Open the document
    doc = word.Documents.Open(filepath)

    # Get the text content from the whole document
  
    text = doc.Range().Text

    # Close the document and quit Word
    doc.Close()
    word.Quit()

    return text
  except Exception as e:
    print(f"Error reading file: {filepath} - {e}")
    return None

folderLoc = r"D:\\OST_Assignment\\Sample2"
# Example usage
filepath = os.listdir(folderLoc)

def extract_text(text):
  phone_regex = r"\+?\d{1,3}[- \.]?\d{3}[- \.]?\d{4,5}"
  email_regex = r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z]{2,}"#r"\b[A-Za-z0-9+_. -]+@[A-Za-z.]\b"
  phone_numbers = re.findall(phone_regex, text)
  emails = re.findall(email_regex, text)
  emailIDs = set()
  number = set()
  if emails:
    for email in emails:
      emailIDs.add(email)
  if phone_numbers:
    for num in phone_numbers:
      num = num.replace(" ","").replace("-","")
      if len(num) == 10:
        number.add(num)
  else:
    print("ERROR: Invalid phone number or Email IDs")
  return emailIDs, number
# for cv in filepath:
#   path = f"{folderLoc}\\{cv}"
#   resume_text = None
#   if not cv.endswith(".pdf"):
#     resume_text = read_doc_file(path)
#   else:
#     reader = PdfReader(path)
#     number_of_pages = len(reader.pages)
#     page = reader.pages[0]
#     resume_text = page.extract_text()
#   if resume_text:
#     emailID, phnN = extract_text(resume_text)
  
  
#   print()


open('extracted.txt','w').write()