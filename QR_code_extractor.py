from pdf2docx import Converter
import zipfile
import shutil
import os
from PIL import Image
from pyzbar import pyzbar
import pandas as pd


#Enter File to convert
Input_file = input("Enter File: ")


def folder_deletion():
    folders = ["_rels","customXml","docProps","word"]
    os.remove('[Content_Types].xml')
    for folder in folders:
        shutil.rmtree(folder, ignore_errors=True)

#Check for QR code and export url found
def QR_Code_read():
    files = os.listdir(".//word//media//")
    for file in files:
        img = Image.open(f'.//word//media//{file}')
        output = pyzbar.decode(img)
        df = pd.DataFrame(output)
        if df.empty:
            print(f'No QR code found in {file}')
        else:
            QR_Code = df['data'][0]
            print(f"QR Code found in {file}:")
            print(QR_Code)

#Convert File
def convert(file):
    if file.endswith('.pdf'):
        docx = file.strip('.pdf')
        docx_file = f'{docx}.docx'
        print('Converting File')
        cv = Converter(file)
        cv.convert(docx_file)
        cv.close()
        print('Extracting docx')
        with zipfile.ZipFile(docx_file, 'r') as zip_ref:
            zip_ref.extractall()
        QR_Code_read()
        folder_deletion()
    elif file.endswith('.docx') :
        print('Extracting docx')
        with zipfile.ZipFile(file, 'r') as zip_ref:
            zip_ref.extractall()
        QR_Code_read()
        folder_deletion()
    else :
        print('File Type not supported')


convert(Input_file)