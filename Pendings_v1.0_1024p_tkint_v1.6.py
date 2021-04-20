# -*- coding: utf-8 -*-
"""  
This script utilizes pyautogui to print various pendings(via mnemonics) in MediTech.
Files are collected as numerous text files in a folder on the Desktop and then text
files are collated into 1 pdf file. Key logging and mouse pointer clicking 
functionality from pyautogui are used. Current parameters are set ONLY for a 
1280x1024 resolution screen.
"""
import os, sys 
from PyPDF2 import PdfFileMerger,PdfFileReader
from datetime import datetime
from pyautogui import time, press, typewrite, hotkey, position, click, moveTo, keyDown, keyUp
from collections import OrderedDict
import datetime
import glob
import shutil
import gc
import tkinter as tk
import tkinter.messagebox
#current_date = datetime.strftime(datetime.now(),'%Y-%m-%d-%H-%M-%S')
"""
user_initials = input("Enter your initials:\n")
  


print("\nPlease select the number for the pendings you would like to print:\n\n1.Biofire\n2.Urgents\n3.Ampliprep\n4.Genetics\n5.GXP\n6.Invitae\n7.NGS\n8.Specialised_Gen\n9.TB")

pend_choice = input('Enter your choice here:\n')

print("\nThank you. Please click on the'LAB Outstanding Test Report'screen(52->12) on MediTech.")
"""


Pendings_dict = {'Biofire':["ZPMRESPSS","ZPMGICAMPY","ZPMH1","ZPMGIADENO"], 
                 'Urgents':["ZPPCRVZV","ZPHSVPCR","ZPMMALID","ZPPCRRIC",
                            "ZPARVO","ZPPCRGBS","ZD_P.jirove.RSL","ZPCMVPCRQ",
                            "ZPCMVPCR","ZPPCREBVCPT","ZPMEBVQL","ZPPCRBRUC","ZPPCRMEAS"], 
                "Routine Genetics": ["ZPPCRCYS","ZPHLAB27","ZPHHC282Y"],
                "GeneXpert": ["ZPENTPCR","ZPCDT","ZPHCVCOUNT", "PPCRHC","ZPPCRMRSA","ZPRSVAB","HPCRFVL","HPCRPT20210",
                              "ZPMBCRABLULT","PHIVPCR"],
                "Invitae": ["ZPINVRESULT","ZPINVNIPSRES","ZPMMGPOTHR","ZPINVCSC","ZPINVCSS","ZPMMGPBGCA","PWAYBILLINV","ZPINVRESULT"], 
                "NGS": ["ZPMOFASS","ZPMOMAHA","ZPMOMANASS"],
                "Specialised Genetics": ["ZPM16SAMP","ZPMFUNGIDNA","ZPPCRCAD","ZPPCRHLAB5701","PPCRDPD","ZPPCRPOR", "ZPMGJB2C26",
                      "ZPPCRBR1","ZPPCRBR185","ZPMBRCAX","PPCRBRCAKF","ZPPCRKV","POSCONKVF",
                      "ZPMIDH1C132","ZPPCRGISTKIT","ZPPCRKD816V","ZPESOT","ZPMMSISTATUS","ZPPCRMLPA1"],
                "Molecular TB":["ZPINH", "ZPINHC", "ZPFLQ", "MTBCSID"],
                "Abbott": ["ZPHPV","ZPHPVG","ZPPCRCHLAM.RSLT","ZPHBVCOUNT","ZPQNTHIVCNT","ZPQNTHIVCNTINS"],
                "HIV Resistance":["ZPHIVRES", "PHIVIG"],
                "Weekend Techs": ["ZPMRESPSS","ZPMGICAMPY","ZPMH1","ZPMGIADENO", "ZPPCRMTB1","ZPENTPCR","ZPCDT","ZPPCRMTB1HISTO","ZPHCVCOUNT","PPCRHC",
                        "ZPPCRMRSA","ZPRSVAB","HPCRFVL","HPCRPT20210","ZPMBCRABLULT","PHIVPCR","ZPPCRNDM1","ZPQNTHIVCNT","ZPQNTHIVCNTINS","ZPHBVCOUNT"],
                "Mid Techs": ["ZPMRESPSS","ZPMGICAMPY","ZPMH1","ZPMGIADENO","ZPENTPCR","ZPCDT","ZPHCVCOUNT", "PPCRHC",
                              "ZPPCRMRSA","ZPRSVAB","HPCRFVL","HPCRPT20210","ZPMBCRABLULT","PHIVPCR", "ZPQNTHIVCNT","ZPQNTHIVCNTINS"],
                 "COVID-2019" : ["ZPMSARS2"], "16S & Fungi ONLY" : ["ZPM16SAMP","ZPMFUNGIDNA"]}

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')   
newpath = os.path.join(desktop,'Pend_Temp') #This filepath is present on all PCs with MediTech installed.
#print(newpath)
if not os.path.exists(newpath): os.makedirs(newpath)

os.chdir(newpath)

def repeat_chunk(pending_code, root, file_count):   
    
    time.sleep(3)   
    
    press(['f6','enter', 'enter','enter'])
    typewrite(pending_code)
    press('enter')
    typewrite(pending_code)

    
    press(['enter','enter','enter','enter','enter', 'enter'])
   
    typewrite('6500')
    press('enter')
    typewrite('I')
    press(['enter','enter','enter','enter'])
    typewrite('PREVIEW')
    press('enter')
    time.sleep(4)


    click(281,53,1)#These are the save to text file pixel co-ordinates for the MEDITECH PDF PRINTER on MECER 1280*1024 screen.
    time.sleep(2)
    press(['tab','tab', 'tab', 'tab', 'space', 'tab', 'enter'])
    time.sleep(2)
    pending_txt = os.path.join(newpath,'{date:%H%M%S}'.format( date=datetime.datetime.now()))
    time.sleep(2)
    typewrite(pending_txt) #Creates first instance of text folder on Desktop with first text filed in mnemonic list.
    time.sleep(1)
    press('enter')
    time.sleep(2)

    click(587,556,1)#This click is to save the text file in MediTech document manager(for spooling).
    time.sleep(2)
    click(1262,9,1)#This click is to close MediTech document manager(for spooling).
    click(938,441,1)#This click is to return MediTech back to the active window.
    press('enter')  

    typewrite('12')#Returns to MediTech Outstanding pendings report page.
    press('enter')
    

input("WELCOME TO YOUR PENDINGS_GENERATOR.\n\nPlease remember to turn off CAPS-LOCK and press ENTER to continue.\n")    
pendings = OrderedDict([(key, value) for key, value in Pendings_dict.items()])

pendings = OrderedDict([(key, value) for key, value in Pendings_dict.items()])

print("\nPlease select the number for the pendings you would like to print:\n")

for i, key in enumerate (pendings.keys()):
    print("{}.{}".format(i, key))
pend_choice = int(input('\nEnter your choice here:\n'))
#print(pyautogui.position())# Used to locate mouse pointer pixel co-ordinates in (x,y) format
time.sleep(5)
pend_count = 1
for value in list(pendings.items())[pend_choice][1]:
    pend_count +=1
    #print(value, pend_count)
    repeat_chunk(value, newpath, int(pend_count))

#window = tkinter.Tk()

root = tk.Tk()
root.title("Pendings Generation Complete!")
label = tk.Label(root, text="Please Return to Pendings Program")
label.pack(side="top", fill="both", expand=True, padx=20, pady=20)
button = tk.Button(root, text="OK", command=lambda: root.destroy())
button.pack(side="bottom", fill="none", expand=True)
root.attributes("-topmost", True) 
root.mainloop()


# creating a simple alert box
#tkinter.messagebox.showinfo("Alert Message", "Pendings Complete! Return to Program!")

#window.mainloop()
#userfilename= ('What do you want to name the pdf?(Please note that ~!*%&^$#/\'? characters are not allowed)')

userfilename= input('Enter PDF name:\n(~!*%&^$#/\'? characters are not allowed)')

x = [a for a in os.listdir() if a.endswith('.pdf')]

merger = PdfFileMerger()

for pdf in x:
    merger.append(open(pdf, 'rb'))
    
with open(userfilename + '.pdf', 'wb') as fout:
    merger.write(fout)

merger.close()  


shutil.copy(userfilename + '.pdf', desktop)

del merger



os.chdir(desktop)

gc.collect()

shutil.rmtree(newpath)

root = tk.Tk()
root.title("Close Program!")
label = tk.Label(root, text="You can close the program now!")
label.pack(side="top", fill="both", expand=True, padx=20, pady=20)
button = tk.Button(root, text="OK", command=lambda: root.destroy())
button.pack(side="bottom", fill="none", expand=True)
root.mainloop()
