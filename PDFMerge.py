import argparse
from glob import glob
import os
import datetime
import PyPDF2
import win32com.client
 

def send_email(subject="", att=""):
    outlook=win32com.client.Dispatch("Outlook.Application")
    Txoutlook = outlook.CreateItem(0)
    Txoutlook.To = "" 
    Txoutlook.Subject = subject
    Txoutlook.BodyFormat = 1 
    Txoutlook.Body = subject
    Txoutlook.Attachments.Add(os.path.join(os.getcwd(),att))
    Txoutlook.Display()
    Txoutlook.Save()
    Txoutlook.Send()    


def main(type="1", name="", directory=".", remarks=""):
 
    today = datetime.datetime.today() 
    today_date = today.strftime("%Y%m%d")   
   
    
    print(today_date)
    print(type)
    
    if type == "1" :
        filename="Claim_"+name+"_"+today_date+"_"+remarks+".pdf"
    elif type == "2": 
        filename="Genex_"+name+"_"+today_date+"_"+remarks+".pdf"
        
        
    
    merger = PyPDF2.PdfMerger()

    for f in glob(f"{directory}/*.pdf"):
        merger.append(f)


    # os.chdir(directory)
    # if not os.path.isdir(sub_dir):
    #     os.mkdir(sub_dir)

    # merger.write(f"{directory}/{sub_dir}/{bookname}.pdf")
    merger.write(f"{directory}\\{filename}")
    merger.close()
    #send_email(filename, directory+"\\"+filename)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-t", "--type", help="Claim:1 / Genex :2")
    parser.add_argument("-n", "--name", help="Name")
    parser.add_argument("-d", "--directory", help="directory where files to be merged live")    
    parser.add_argument("-r", "--remarks", default="merged", help="remarks")  
    args = parser.parse_args()

  
    # print(args.type)
    # print(args.name)
    # print(args.directory)
    
    main(args.type, args.name,args.directory, args.remarks)
    