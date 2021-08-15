import pandas as pd
import pdfkit
import os
import shutil
from datetime import date

TEMP_HTML_FILENAME = 'placeholder.html'
WORKING_DIRECTORY = os.getcwd()

def make_directories():
  '''This creates the directories needed to organize the files'''

  today = date.today()
  directory_name = today.strftime("%m-%d-%y")

  main_directory_path = WORKING_DIRECTORY + '/' + directory_name
  if not os.path.isdir(main_directory_path):
    os.mkdir(main_directory_path)

  pdf_directory_path = main_directory_path + '/pdf'
  if not os.path.isdir(pdf_directory_path):
    os.mkdir(pdf_directory_path)

  xlsx_directory_path = main_directory_path + '/xlsx'
  if not os.path.isdir(xlsx_directory_path):
    os.mkdir(xlsx_directory_path)

  return main_directory_path, xlsx_directory_path, pdf_directory_path

def move_file(filename, current_directory, target_directory):
  '''This moves a file from one directory to another'''

  current_file = current_directory + '/' + filename
  target_file = target_directory + '/' + filename
  #os.rename(current_file, target_file)
  shutil.move(current_file, target_file)
  #os.replace(current_file, target_file)

def convert_file(filename, xlsx_directory, pdf_directory):
  '''This converts xlsx files into an html file, then it
      makes changes to the HTML, adding a style to the
      'email' column, and finally converts it to a PDF.'''

  file_location = WORKING_DIRECTORY + '/' + filename + '.xlsx'
  df = pd.read_excel(file_location)
  df.to_html(TEMP_HTML_FILENAME)
  s = df.style.format('{}')
  s.set_table_styles({
      'email': [{'selector': '', 'props': 'word-wrap: break-word'}]
  }, overwrite=False)

  s.to_html(TEMP_HTML_FILENAME)
  print(TEMP_HTML_FILENAME)
  print(filename + '.pdf')
  pdfkit.from_file(TEMP_HTML_FILENAME, pdf_directory + '/' + filename + '.pdf')
  move_file(filename + '.xlsx', WORKING_DIRECTORY, xlsx_directory)

def find_files():
  '''This goes through all xlsx files in the current directory
      to convert them to PDFs.'''

  main_directory, xlsx_directory, pdf_directory = make_directories()
  for candidate_file in os.listdir('.'):
    if candidate_file.endswith('.xlsx'):
      excel_filename = candidate_file.split('.xlsx',1)[0]
      print(excel_filename)
      convert_file(excel_filename, xlsx_directory, pdf_directory)
  os.remove(WORKING_DIRECTORY + '/' + TEMP_HTML_FILENAME)

find_files()
