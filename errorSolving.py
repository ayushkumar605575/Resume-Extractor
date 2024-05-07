from PyPDF2 import PdfReader
import os, re, shutil
# from docx import Document
# print(type(text))
# print(text)
from spire.doc import Document
# from spire.doc.common import *

# Create a Document object
# Load a Word document


from openpyxl import Workbook

def write_text_to_xls(filepath, data):
  """
  Writes text data to an XLSX (.xlsx) file using openpyxl.

  Args:
      filepath (str): Path to the output XLSX file.
      data (list): A list of lists containing the text data to be written.
  """
  try:
    # Create a new workbook
    workbook = Workbook()

    # Add a worksheet
    worksheet = workbook.active
    worksheet.title = "Extracted Resume's Data"  # Customize worksheet title

    # Write data to rows and columns (starting from row 1, column 1)
    for row_index, row_data in enumerate(data, start=1):
      for col_index, cell_value in enumerate(row_data, start=1):
        worksheet.cell(row=row_index, column=col_index).value = cell_value

    # Save the workbook as an XLSX file
    workbook.save(filepath)
    print(f"Text data written to XLSX file: {filepath}")
  except Exception as e:
    print(f"Error writing data to file: {filepath} - {e}")

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
    # word = win32.gencache.EnsureDispatch('Word.Application')
    # word.Visible = False
    document = Document()
    document.LoadFromFile(filepath)

    # Extract the text of the document
    document_text = document.GetText()
    document_text = document_text.replace("Evaluation Warning: The document was created with Spire.Doc for Python.","")
    document.Close()
    return document_text
    # Open the document
    # doc = word.Documents.Open(filepath)

    # # Get the text content from the whole document
  
    # text = doc.Range().Text

    # # Close the document and quit Word
    # doc.Close()
    # word.Quit()

    # return full_text
  except Exception as e:
    print(f"Error reading file: {filepath} - {e}")
    return None

folderLocc = r"Sample2"
# Example usage

def extract_text(text):
  phone_regex = r"\+?\d{1,3}[- \.]?\d{3,5}[- \.]?\d{4,5}"
  email_regex = r"[a-z0-9_.+-]+\s?+@[a-z0-9-]+\s?+\.[a-z]{2,}"#r"\b[A-Za-z0-9+_. -]+@[A-Za-z.]\b"
  phone_numbers = re.findall(phone_regex, text)
  emails = re.findall(email_regex, text)
  emailIDs = list()
  number = list()
  if emails:
    for email in emails:
      email = email.replace(" ", "")
      emailIDs.append(email)
  if phone_numbers:
    for num in phone_numbers:
      num = num.replace(" ","").replace("-","")
      if len(num) >= 10:
        number.append(num)
  else:
    print("ERROR: Invalid phone number or Email IDs")
  return emailIDs, number

def findCVs(folderPath, cvPaths):
  for file in os.listdir(folderPath):
    if os.path.isdir(f"{folderPath}\\{file}"):
      findCVs(folderPath=f"{folderPath}\\{file}", cvPaths=cvPaths)
    if os.path.isfile(f"{folderPath}\\{file}") and not f"{folderPath}\\{file}".endswith(".zip"):
      cvPaths.append(f"{folderPath}\\{file}")

def main(folderLoc):
  cvPaths = []
  findCVs(folderPath=folderLoc, cvPaths=cvPaths)
  # for file in os.listdir(folderLoc):
  #   if os.path.isdir(f"{folderLoc}\\{file}"):
  #     filepath = os.listdir(f"{folderLoc}\\{file}")
  #     for f in filepath:
  #       if os.path.isfile(f"{folderLoc}\\{file}\\{f}"):
  #         cvPaths.append(f"{folderLoc}\\{file}\\{f}")

  text_data = [
    ["Email", "Phone", "Resume Overall Content"],
  ]
  c = 0
  for path in cvPaths:
    tmp = []
    # path = f"{folderLoc}\\{cv}"
    # print(path)
    resume_text = None
    if not path.endswith(".pdf"):
      resume_text = read_doc_file(path)
    else:
      reader = PdfReader(path)
      # number_of_pages = len(reader.pages)
      page = reader.pages[0]
      resume_text = page.extract_text()
    if resume_text:
      emailID, phnN = extract_text(resume_text)
      if len(emailID) >= 1:
        tmp.append(emailID[0])
      else:
        tmp.append("")
      if len(phnN) >= 1:
        tmp.append(phnN[0])
      else:
        tmp.append("")
      tmp.append(resume_text)
    text_data.append(tmp)
  print(c)

  write_text_to_xls(r"downloads\\extractedData.xlsx", text_data)

  try:
    shutil.rmtree(folderLoc, ignore_errors=True)  # Set ignore_errors to True
    print("Directory removed successfully (if it was empty)!")
  except OSError as e:
    print(f"Error removing directory: {e}")

# main("uploads\\tmp")