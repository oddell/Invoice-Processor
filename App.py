from ast import Return
from operator import truth
import sys
import os
import shutil
from tkinter.constants import FALSE
from openpyxl import Workbook, load_workbook
from os import error, listdir, remove
from os.path import isfile, join
from google.cloud import documentai
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from PyPDF2 import PdfFileWriter, PdfFileReader
from PIL import Image,ImageTk
from pdf2image import convert_from_path

#Globals
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'Application\\bright-velocity-333609-6081a4699c11.json'
excelpath = 'Application\\Global Invoice Reconcilliation.xlsx'
invoicespath = "Application/PDF Input"
progress = 0

#CLOUD INFO
project_id = "bright-velocity-333609"
location = "eu"  # Format is 'us' or 'eu'
processor_id = "fab48bf677935aeb"  # Create processor in Cloud Console

def DoProcurementAI(project_id: str, location: str, processor_id: str, file_path: str):


    # You must set the api_endpoint if you use a location other than 'us', e.g.:
    opts = {}
    if location == "eu":
        opts = {"api_endpoint": "eu-documentai.googleapis.com"}

    client = documentai.DocumentProcessorServiceClient(client_options=opts)

    # The full resource name of the processor, e.g.:
    # projects/project-id/locations/location/processor/processor-id

    name = f"projects/{project_id}/locations/{location}/processors/{processor_id}"

    # Read the file into memory
    with open(file_path, "rb") as image:
        image_content = image.read()

    document = {"content": image_content, "mime_type": "application/pdf"}

    # Configure the process request
    request = {"name": name, "raw_document": document}

    result = client.process_document(request=request)
    
    document = result.document
    entities = document.entities

    # For a full list of Document object attributes, please reference this page: https://googleapis.dev/python/documentai/latest/_modules/google/cloud/documentai_v1beta3/types/document.html#Document
    types = []
    values = []
    confidence = []
    normalizedValues = []
    # Read the text recognition output from the processor

    for entity in entities:
        types.append(entity.type_)
        values.append(entity.mention_text)
        confidence.append(round(entity.confidence,4))
        normalizedValues.append(entity.normalized_value.text)
    
    invoiceInfo = dict(zip(types, values))
    invoiceNormalizedInfo = dict(zip(types, normalizedValues))

    for key in invoiceNormalizedInfo:
        if invoiceNormalizedInfo[key] != "":
            invoiceInfo[key] = invoiceNormalizedInfo[key]

    return(invoiceInfo)

def FindProjectNo(invoiceInfo):
    with open('Application\ProjectList.txt') as f:
        projectlist = f.read().splitlines()
    for project in projectlist:
        for information in invoiceInfo.values():
            if project in information:
                invoiceInfo["project_no"] = project
                break
        else:
            continue
        break

    return invoiceInfo

def showpdf(PDF_PATH, key):

    reject = False
    root = tk.Tk()
    pdf_frame = tk.Frame(root,width=1400).pack(fill=tk.BOTH,expand=1)
    scrol_y = tk.Scrollbar(pdf_frame,orient=tk.VERTICAL)

    pdf = tk.Text(pdf_frame,yscrollcommand=scrol_y.set,bg="grey")

    scrol_y.pack(side=tk.RIGHT,fill=tk.Y)
    scrol_y.config(command=pdf.yview)

    pdf.pack(fill=tk.BOTH,expand=1)

    pages = convert_from_path(PDF_PATH, poppler_path= r"Application\poppler-21.11.0\library\bin")  #size=(1200,1500)


    photos = []

    for i in range(len(pages)):
        w, h = pages[i].size
        if w > h:
           pages[i] = pages[i].resize((1168, 927))
           pages[i] = pages[i].rotate(90)
           
        photos.append(ImageTk.PhotoImage(pages[i]))

    # Adding all the images to the text widget
    for photo in photos:
        pdf.image_create(tk.END,image=photo)
    
    # For Seperating the pages
        pdf.insert(tk.END,'\n\n')

    #Inputbox

    userinput = tk.StringVar()


    def submit_clicked():
        """ callback when the login button clicked
        """
        Reject.reject = False
        userinput.get()
        root.destroy()
    
    def Reject():
        Reject.reject = True
        root.destroy()
        
        

    userinput_label = tk.Label(pdf_frame, text="Input "+key+":")
    userinput_label.pack(fill='x', expand=True)

    userinput_entry = tk.Entry(pdf_frame, textvariable=userinput)
    userinput_entry.pack(fill='x', expand=True)
    userinput_entry.focus()

    # login button
    login_button = tk.Button(pdf_frame, text="Submit", command=submit_clicked)
    login_button.pack(fill='x', expand=True, pady=10)

    login_button = tk.Button(pdf_frame, text="Reject", command=Reject)
    login_button.pack(fill='x', expand=True, pady=10)

    root.mainloop()
    enteredvalue = userinput.get()
    print(enteredvalue+" Manually Entered")
    if Reject.reject == True:
        showpdf.reject = True
    else:
        return(enteredvalue)

def TroubleshootInfo(invoiceInfo, PDF_PATH):
    keys = ("invoice_date", "supplier_name", "invoice_id", "line_item", "net_amount","project_no")
    invoiceInfo["IsHire"] = "Purchase"
    chars = "£Ł,"
    pdfopen = False
    for key in keys:
        try:
            invoiceInfo[key]
        except KeyError:
            found = False
            if key == "net_amount":
                try:
                    for c in chars:
                        invoiceInfo['total_amount'] = invoiceInfo['total_amount'].replace(c, "")
                        invoiceInfo['total_tax_amount'] = invoiceInfo['total_tax_amount'].replace(c,"")
                    invoiceInfo['net_amount'] = float(invoiceInfo['total_amount']) - float(invoiceInfo['total_tax_amount'])
                    try:
                        print("Calculated Net of "+str(invoiceInfo['total_amount'])+" - "+str(invoiceInfo['total_tax_amount'])+" = "+str(invoiceInfo['net_amount'])+" for invoice "+invoiceInfo["invoice_id"])
                        found = True
                    except KeyError:
                        found = False
                        
                except:
                    continue
            if  key == "invoice_id":
                try:
                    invoiceInfo[key] = invoiceInfo["purchase_order"]
                    print("Invoice ID Altered to "+ invoiceInfo[key])
                    found = True
                except KeyError:
                    continue
            if found == False:
                print("Error: "+key+" not found!")
                invoiceInfo[key]=showpdf(PDF_PATH, key)
                if showpdf.reject == True:
                    return
               
    
    if "mildren" in str(invoiceInfo["supplier_name"]).lower():
        invoiceInfo["supplier_name"]=showpdf(PDF_PATH, key="supplier_name")

    try:
        if invoiceInfo['net_amount']:   
            try:
                for c in chars:
                    invoiceInfo['net_amount'] = invoiceInfo['net_amount'].replace(c, "")
            except:
                ""
    except KeyError:
        invoiceInfo['net_amount']=showpdf(PDF_PATH, key="net_amount")

    #Standardise Date
    try:
        invoiceInfo['excel_date'] = (pd.to_datetime(invoiceInfo['invoice_date'], dayfirst=True))
    except:
        invoiceInfo["excel_date"]= (pd.to_datetime(showpdf(PDF_PATH, key="invoice_date"), dayfirst=True))

    invoiceInfo['invoice_date'] = (pd.to_datetime(invoiceInfo['excel_date'], dayfirst=True)).strftime("%b %Y")
     
    invoiceInfo["invoice_id"]=invoiceInfo["invoice_id"].replace("\n","")

    #Check if hire invoice
    for information in invoiceInfo.values():
        if "hire" in str(information).lower():
            invoiceInfo["IsHire"] = "Hire"
    return(invoiceInfo)

def updateexcel(invoiceInfo, PDF_OUTPUT_PATH):

    wb = load_workbook(excelpath)
    ws = wb.active
    ws.append([invoiceInfo["excel_date"], invoiceInfo["project_no"] , invoiceInfo["supplier_name"] , '=HYPERLINK("{}", "{}")'.format(PDF_OUTPUT_PATH[12:], invoiceInfo["invoice_id"]), invoiceInfo["IsHire"], invoiceInfo["line_item"] ,float(invoiceInfo["net_amount"])])
    AttemptSave= True
    #while AttemptSave == True:
        #try:
    wb.save(excelpath)

def movepdf(invoiceInfo, PDF_OUTPUT_PATH, PDF_PATH):
    os.makedirs(os.path.dirname(PDF_OUTPUT_PATH), exist_ok=True)
    try:
        shutil.move(PDF_PATH,PDF_OUTPUT_PATH)
    except error:
        print(PDF_OUTPUT_PATH+" has failed! (File Path Error)")
        invoiceInfo["invoice_id"] = input("Input filename: ")
        PDF_OUTPUT_PATH = "Application/PDF Output/"+invoiceInfo["project_no"]+"/"+invoiceInfo['invoice_date']+"/"+invoiceInfo['invoice_id']+".pdf"
        shutil.move(PDF_PATH,PDF_OUTPUT_PATH)

def splitpdfs(invoicespath):
    PDF_PATHS = [invoicespath+"/"+f for f in listdir(invoicespath) if isfile(join(invoicespath,f))]

    for PDF_PATH in PDF_PATHS:
        with open(PDF_PATH, mode='rb') as r:
            inputpdf = PdfFileReader(r)
            print(PDF_PATH)
            for i in range(inputpdf.numPages):
                output = PdfFileWriter()
                output.addPage(inputpdf.getPage(i))
                with open(invoicespath+"/"+PDF_PATH[len(invoicespath)+1:]+str(i)+".pdf", "wb") as outputStream:
                    output.write(outputStream)
            
        remove(PDF_PATH)

def SplitInvoicesQuery():
    box = tk.Tk()
    box.geometry("100x100") 
    res = messagebox.askquestion('PDF Splitter', 'Do any of the PDFs have more than one invoice?')
    if res == 'yes':
        splitpdfs(invoicespath)

    box.destroy()

def Main():
    SplitInvoicesQuery()

    PDF_PATHS = [invoicespath+"/"+f for f in listdir(invoicespath) if isfile(join(invoicespath,f))]
    progress = 0
    print(str(len(PDF_PATHS))+" PDFs found.")

    for PDF_PATH in PDF_PATHS:   
        showpdf.reject = False
        invoiceInfo = DoProcurementAI(project_id, location, processor_id, PDF_PATH)

        invoiceInfo = FindProjectNo(invoiceInfo)
        invoiceInfo = TroubleshootInfo(invoiceInfo, PDF_PATH)
        progress = progress + 1  
        if showpdf.reject == False:
            PDF_OUTPUT_PATH = "Application/PDF Output/"+invoiceInfo["project_no"]+"/"+invoiceInfo['invoice_date']+"/"+invoiceInfo['invoice_id'].replace("/"," ")+".pdf"
            updateexcel(invoiceInfo, PDF_OUTPUT_PATH)
            movepdf(invoiceInfo, PDF_OUTPUT_PATH, PDF_PATH)
            print("Document "+PDF_PATH[len(invoicespath)+1:]+" Processed "+str(progress)+"/"+str(len(PDF_PATHS)))
            invoiceInfo.clear()               
        else:
            pdfRejectPath = "Application/PDF Output/Rejects/"+PDF_PATH[len(invoicespath)+1:]+".pdf"
            movepdf(invoiceInfo, pdfRejectPath, PDF_PATH)
            print("Document "+PDF_PATH[len(invoicespath)+1:]+" Rejected "+str(progress)+"/"+str(len(PDF_PATHS)))

if __name__ == '__main__':
    sys.exit(Main())