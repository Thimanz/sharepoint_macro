#import all the libraries
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
import io
from dotenv import load_dotenv
import os
import pandas as pd

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
  file = ctx.web.get_file_by_server_relative_url(relative_url)
  ctx.load(file)
  ctx.execute_query()
  print("Authentication successful")

response = File.open_binary(ctx, relative_url)

#save data to BytesIO stream
bytes_file_obj = io.BytesIO()
bytes_file_obj.write(response.content)
bytes_file_obj.seek(0) #set file object to start

#read excel file and each sheet into pandas dataframe 
df = pd.read_excel(bytes_file_obj, sheet_name = None)['Planilha1']

#update df with new row
new_row = pd.Series({'aaa':'Hyperion', 'bbb':24000, 'ccc':'55days', 'ddd':1800})
df = pd.concat([df, new_row.to_frame().T], ignore_index=True)

# Save updated pandas dataframe to memory as Excel file
updated_excel = io.BytesIO()
with pd.ExcelWriter(updated_excel) as writer:
    df.to_excel(writer, index=False, sheet_name='Planilha1')
updated_excel.seek(0)

# Upload updated Excel file to SharePoint and overwrite existing file
if ctx_auth.acquire_token_for_user(username, password):
  file.save_binary_stream(updated_excel)
  ctx.execute_query()
  print("Updated successfully")