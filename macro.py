#import all the libraries
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
from selenium import webdriver
from selenium.webdriver.common.by import By
import io
from dotenv import load_dotenv
import os
import openpyxl
from tempfile import NamedTemporaryFile

load_dotenv()

#target url taken from sharepoint and credentials
sharepoint_site = os.getenv("SHAREPOINT_SITE")
relative_url = os.getenv("SHEET_RELATIVE_URL")
url = os.getenv("SHEET_URL")
username = os.getenv("SHAREPOINT_USER")
password = os.getenv("SHAREPOINT_PASSWORD")

ctx_auth = AuthenticationContext(sharepoint_site)
if ctx_auth.acquire_token_for_user(username, password):
  ctx = ClientContext(sharepoint_site, ctx_auth)
  web = ctx.web
  ctx.load(web)
  #file will be used to update later
  file = ctx.web.get_file_by_server_relative_url(relative_url)
  ctx.load(file)
  ctx.execute_query()
  print("Authentication successful")

response = File.open_binary(ctx, relative_url)

#save data to BytesIO stream
bytes_file_obj = io.BytesIO()
bytes_file_obj.write(response.content)
bytes_file_obj.seek(0) #set file object to start

#get data from webpage
driver = webdriver.Chrome()
driver.get("https://pt.wikipedia.org/wiki/Segunda_Guerra_Mundial")
text = driver.find_element(By.CLASS_NAME, "mw-page-title-main").text

#manipulate worksheet with openpyxl
workbook = openpyxl.load_workbook(bytes_file_obj)
worksheet = workbook[os.getenv("WORKSHEET_NAME")]

worksheet.append((text, 132, "olha", 4132))

#save updated excel file to memory
with NamedTemporaryFile(delete=False) as tmp:
  workbook.save(tmp.name)
  tmp.seek(0)
  updated_excel = tmp.read()
  tmp.close()
  os.remove(tmp.name)

# Upload updated Excel file to SharePoint and overwrite existing file
if ctx_auth.acquire_token_for_user(username, password):
  file.save_binary_stream(updated_excel)
  ctx.execute_query()
  print("Updated successfully")