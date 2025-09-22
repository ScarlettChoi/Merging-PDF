import os
import pypdf
from datetime import datetime
import win32com.client
import pythoncom
from config import *

pathfile='setting.properties'
name,email=loadvar(pathfile)

def merge_pdfs(folder_path, output_path):
    merger = pypdf.PdfWriter()
    for root, _, files in os.walk(folder_path):
        pdf_files = [os.path.join(root, file) for file in files if file.lower().endswith('.pdf')]
        for pdf in sorted(pdf_files):
            merger.append(pdf)
    merger.write(output_path)
    merger.close()

def send_email(subject="", address="" ,att=""):
    time=datetime.now()
    print(f"send email module : {datetime.now()}")
    try: 
        outlook=win32com.client.Dispatch("Outlook.Application")
        print(f"Connect Success : {datetime.now()}")
    except pythoncom.com_error as e:
        print(e)
        outlook=win32com.client.DispatchEx("Outlook.Application")
    print("writting Email....")    
    Txoutlook = outlook.CreateItem(0)
    print("Opening Email....")  
    Txoutlook.To = address
    print("Entering Address....")  
    Txoutlook.Subject = subject
    print("Entering Subject....")
    Txoutlook.BodyFormat = 1
    print("Entering BodyFormat....") 
    Txoutlook.Body = subject
    print(f"Entering Body....{datetime.now()}")
    Txoutlook.Attachments.Add(att)
    print("Writing Email Done...")
    #Txoutlook.Display()
    #Txoutlook.Save()
    Txoutlook.Send()
    print(f"Sending Email Done...{datetime.now()}")    

if __name__ == "__main__":
    path=input("Folder path for merging :")
    folder_path = rf'{path}'
    today = datetime.today().strftime('%Y%m%d')
    file_type_s=input("File type :\n1. General Expense\n2. Personal Claim\n")
    if file_type_s=="1": file_type="[GENEX]"
    elif file_type_s=="2": file_type="[Claim]"
    else :
        print("Type error.Try again")
        exit()
    email_yn=input("Do you want to send email? (Y/N)")
    if email_yn=="Y":
        email_address=input("Please type email address (Press Enter : vfskvat@volvo.com):")
        if email_address=="":
            email_address=email
        print(f"Email address : {email_address}")
    elif email_yn=="N":
        pass
    else :
        print("Type Y/N in uppercase letter. Try again")
        exit()
    output_path = f'{folder_path}/{file_type}{name}_{today}.pdf'
    email_subject=f'{file_type}{name}_{today}'
    merge_pdfs(folder_path, output_path)
    print(f'Merged PDF saved as {output_path}')
    if email_yn=="Y":
        try :
            send_email(email_subject,email_address, output_path)
        except Exception as e :
            print(e)
    print("All Jobs Successfully Done...")

