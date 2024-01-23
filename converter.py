import os
import win32com.client as client

excel = client.Dispatch("Excel.Application")

for file in os.listdir(os.getcwd() + "/certificados.xls/"):
  filename, fileextension = os.path.splitext(file)
  wb = excel.Workbooks.Open(os.getcwd() + "/certificados.xls/" + file)
  output = os.getcwd() + "/certificados.xlsx/" + filename
  wb.SaveAs(output, FileFormat = 51)
  wb.Close()