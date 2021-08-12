import os
from shutil import copyfile
import re
import PyPDF2
from pathlib import Path
from colorama import init,Fore

# ██████╗░███████╗██╗░░░██╗███████╗██╗░░░░░░█████╗░██████╗░███████╗██████╗░  ██████╗░██╗░░░██╗
# ██╔══██╗██╔════╝██║░░░██║██╔════╝██║░░░░░██╔══██╗██╔══██╗██╔════╝██╔══██╗  ██╔══██╗╚██╗░██╔╝
# ██║░░██║█████╗░░╚██╗░██╔╝█████╗░░██║░░░░░██║░░██║██████╔╝█████╗░░██║░░██║  ██████╦╝░╚████╔╝░
# ██║░░██║██╔══╝░░░╚████╔╝░██╔══╝░░██║░░░░░██║░░██║██╔═══╝░██╔══╝░░██║░░██║  ██╔══██╗░░╚██╔╝░░
# ██████╔╝███████╗░░╚██╔╝░░███████╗███████╗╚█████╔╝██║░░░░░███████╗██████╔╝  ██████╦╝░░░██║░░░
# ╚═════╝░╚══════╝░░░╚═╝░░░╚══════╝╚══════╝░╚════╝░╚═╝░░░░░╚══════╝╚═════╝░  ╚═════╝░░░░╚═╝░░░

# ░░░░░██╗░█████╗░░██████╗██████╗░██╗███╗░░██╗  ██╗░░██╗░█████╗░██████╗░██╗░░██╗██╗
# ░░░░░██║██╔══██╗██╔════╝██╔══██╗██║████╗░██║  ██║░██╔╝██╔══██╗██╔══██╗██║░██╔╝██║
# ░░░░░██║███████║╚█████╗░██████╦╝██║██╔██╗██║  █████═╝░███████║██████╔╝█████═╝░██║
# ██╗░░██║██╔══██║░╚═══██╗██╔══██╗██║██║╚████║  ██╔═██╗░██╔══██║██╔══██╗██╔═██╗░██║
# ╚█████╔╝██║░░██║██████╔╝██████╦╝██║██║░╚███║  ██║░╚██╗██║░░██║██║░░██║██║░╚██╗██║
# ░╚════╝░╚═╝░░╚═╝╚═════╝░╚═════╝░╚═╝╚═╝░░╚══╝  ╚═╝░░╚═╝╚═╝░░╚═╝╚═╝░░╚═╝╚═╝░░╚═╝╚═╝

# Local Folder or Full directory where all the PDF answer files are saved
directory = r'submitted_answers'
# Local Folder or Full directory where all the formatted files are need to be saved
destination = r'corrected_file_name_answers'

init(autoreset=True)

# Create necessary folder if not exist
Path(directory).mkdir(parents=True, exist_ok=True)
Path(destination).mkdir(parents=True, exist_ok=True)

# File Extension Here. Example: PDF
fileExtension="pdf"
totalSubmission=0
totalFileRangeFound=0
noSymbolNoCount=0
totalValidFiles=0
totalCorruptedFiles=0
print(Fore.LIGHTCYAN_EX+"File Formatter v3 \n----------------- \n"+Fore.RESET)
print(Fore.LIGHTCYAN_EX+"Available Choice \n1.Format File \n2.Format File with symbol no range \n3.Corrupt File Checker \nEnter your choice:"+Fore.RESET)
choice=int(input())
print("*****************************************************************\n")

if choice==1:
    # Exam Center Code Here.
    examCenterCode= input("Enter Exam Center Code \n")
    # Subject Name Here.
    subjectName = input("Enter the Subject Name \n")
    # Faculty Name Here.
    faculty = input("Enter Faculty. Example: BE-IT \n")
    # File Format Here. Example: 195-BE-IT_Subject Name_SymbolName
    fileFormat=examCenterCode+"_"+faculty+"_"+subjectName+"_"
    try:
        for filename in os.listdir(directory):
            if filename.endswith(".pdf") or filename.endswith(".PDF"):
                symbolNumber = re.search("(?<!\d)\d{8,10}(?!\d)", filename)

                if symbolNumber:
                    symbolNo=symbolNumber.group()
                    try:
                        copyfile(os.path.join(directory,filename),os.path.join(destination,fileFormat+symbolNo+"."+fileExtension))
                    except:
                        print(Fore.RED+"System is Unable to Copy "+filename+" may be because of some system error, consider coping it manually."+Fore.RESET)
                        pass
                    print(Fore.LIGHTGREEN_EX+filename+" >>> checked, formatted and saved successfully!"+Fore.RESET+"\n--------------------------------------------------------------------------------\n")
                    totalFileRangeFound=totalFileRangeFound+1
                else:
                    print(Fore.LIGHTYELLOW_EX+"Symbol number is NOT PRESENT or INCORRECT LENGTH in this file!: "+filename+Fore.RESET+"\n--------------------------------------------------------------------------------\n")
                    symbolNo="SymbolNoHere"
                    noSymbolNoCount=noSymbolNoCount+1
                    try:
                        copyfile(os.path.join(directory,filename),os.path.join(destination,fileFormat+symbolNo+str(noSymbolNoCount)+"."+fileExtension))
                    except:
                        print(Fore.LIGHTRED_EX+"System is Unable to Copy "+filename+" may be because of some system error, consider coping it manually."+Fore.RESET)
                        pass
                totalSubmission=totalSubmission+1
            else:
                continue
        print(Fore.LIGHTMAGENTA_EX+"Total Files = "+str(totalSubmission)+Fore.RESET+Fore.LIGHTGREEN_EX+"\nTotal Files Formatted = "+str(totalFileRangeFound)+Fore.RESET+Fore.LIGHTYELLOW_EX+"\nTotal Files without Symbol No = "+str(noSymbolNoCount)+Fore.RESET)
    except OSError:
        print("Could not open/read file, may be this directory does not exist \n---------------------------------------\n"+directory+"\n---------------------------------------\nso check the file path again")
elif choice==2:
    # Exam Center Code Here.
    examCenterCode= input("Enter Exam Center Code \n")
    # Subject Name Here.
    subjectName = input("Enter the Subject Name \n")
    # Faculty Name Here.
    faculty = input("Enter Faculty. Example: BE-IT \n")
    # Range of symbol number
    symbolNoRangeStart=int(input("Enter symbol number range: starting number\n"))
    symbolNoRangeEnd=int(input("Enter symbol number range: ending number\n"))
    # File Format Here. Example: 195-BE-IT_Subject Name_SymbolName
    fileFormat=examCenterCode+"_"+faculty+"_"+subjectName+"_"

    try:
        for filename in os.listdir(directory):
            if filename.endswith(".pdf") or filename.endswith(".PDF"):
                #copyfile(os.path.join(directory,filename),os.path.join(destination,str(counter)+".pdf"))
                symbolNumber = re.search("(?<!\d)\d{8,10}(?!\d)", filename)
                Path(destination+'/'+subjectName).mkdir(parents=True, exist_ok=True)
                if symbolNumber:
                    symbolNo=int(symbolNumber.group())
                    if symbolNo >= symbolNoRangeStart and symbolNo <= symbolNoRangeEnd:
                        try:
                            copyfile(os.path.join(directory,filename),os.path.join(destination,subjectName+'/'+fileFormat+str(symbolNo)+"."+fileExtension))
                        except:
                            print(Fore.LIGHTRED_EX+"System is Unable to Copy "+filename+" may be because of some system error, consider coping it manually."+Fore.RESET)
                            pass
                        print(Fore.LIGHTGREEN_EX+filename+" >>> checked, formatted and saved successfully!"+Fore.RESET+" \n--------------------------------------------------------------------------------\n")
                        totalFileRangeFound=totalFileRangeFound+1
                else:
                    print(Fore.LIGHTYELLOW_EX+"Symbol number is NOT PRESENT or INCORRECT LENGTH in this file!: "+filename+Fore.RESET+"\n--------------------------------------------------------------------------------\n")
                    symbolNo="SymbolNoHere"
                    noSymbolNoCount=noSymbolNoCount+1
                    try:
                        copyfile(os.path.join(directory,filename),os.path.join(destination,subjectName+'/'+fileFormat+symbolNo+str(noSymbolNoCount)+"."+fileExtension))
                    except:
                        print(Fore.LIGHTRED_EX+"System is Unable to Copy "+filename+" may be because of some system error, consider coping it manually. \n"+Fore.RESET)
                        pass
                totalSubmission=totalSubmission+1
            else:
                continue
        print(Fore.LIGHTMAGENTA_EX+"Total Files = "+str(totalSubmission)+Fore.RESET+Fore.LIGHTGREEN_EX+"\nTotal Files with the range = "+str(totalFileRangeFound)+Fore.RESET+Fore.LIGHTYELLOW_EX+"\nTotal Files without Symbol No = "+str(noSymbolNoCount)+Fore.RESET+Fore.LIGHTGREEN_EX+"\nSuccessfully!! saved to folder: "+subjectName+Fore.RESET)
    except OSError:
        print(Fore.LIGHTYELLOW_EX+"Could not open/read file, may be this directory does not exist"+Fore.RESET+" \n---------------------------------------\n"+directory+"\n---------------------------------------\nso check the file path again")
elif choice==3:
    print(Fore.LIGHTCYAN_EX+"Do you want to check corrupted files from \n1."+directory+"\n2."+destination+":"+Fore.RESET)
    folderChoice = int(input())
    if folderChoice==1:
        correctFormatDirectory=directory
    elif folderChoice==2:
        correctFormatDirectory=destination
    else:
        print("invalid choice")
        input()
    for root, dirs, files in os.walk(correctFormatDirectory):
        for filename in files:
            #print("file name:"+filename+" in "+os.path.join(root, filename))
            try :
                if filename.endswith((".pdf", ".PDF")):
                    print(Fore.LIGHTYELLOW_EX+"checking..."+Fore.RESET)
                    sample_pdf = open(os.path.join(root,filename), mode='rb')
                    pdfdoc = PyPDF2.PdfFileReader(sample_pdf)
                    print(Fore.LIGHTGREEN_EX+filename+" is valid PDF file"+Fore.RESET)
                    totalValidFiles=totalValidFiles+1
            except:
                print(Fore.LIGHTCYAN_EX+"Oops! "+Fore.RESET+Fore.LIGHTRED_EX+filename+" is corrupted PDF file"+Fore.RESET)
                print(Fore.LIGHTCYAN_EX+"File Location: "+Fore.RESET+os.path.join(root,filename))
                totalCorruptedFiles=totalCorruptedFiles+1
    print(Fore.LIGHTGREEN_EX+"\nTotal Valid Files = "+str(totalValidFiles)+Fore.RESET+Fore.LIGHTRED_EX+"\nTotal Corrupted Files = "+str(totalCorruptedFiles)+Fore.RESET)
else:
    print("invalid choice")
    
#just to hold the console
input()