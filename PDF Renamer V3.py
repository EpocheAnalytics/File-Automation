#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# Importing required modules
from tkinter import Tk
from tkinter.filedialog import askdirectory
from datetime import datetime
import glob
import PyPDF2 
import re
import io 
import pandas as pd
import os

#iterator 
i = 0

#Loop to allow user to pick naming convention
while i < 1:
    try:
        decision = input("Press 1 and enter to rename PDFs with only an invoice number. Press 2 and enter to rename pdfs with a buisness name included. ")
        int(decision)
        it_is = True
        if it_is == True:
            decision = int(decision)
            if decision == 1:
                print('You have chosen to rename PDFs with only an invoice number.')
                i = i+1
            elif decision == 2:
                print('You have chosen to rename PDFs with a company name and invoice number.')
                i = i+1
            else:
                print('You have entered an incorrect value, please try again.')
    except ValueError:
        print('You have entered an incorrect value, please try again.')
    
print('Please choose the location of the PDFs you wish to rename.')

#Selecting File Path
path = askdirectory(title='Select Folder') # shows dialog box and return the path

string_input_with_date = "12/08/2022"
past = datetime.strptime(string_input_with_date, "%d/%m/%Y")
present = datetime.now()
if past.date() > present.date():
    print('Starting PDF Processing')
    #Grabbing all of the files and looping through them
    for name in glob.glob((path+'/*.pdf')):
        
        #setting locations
        location = name
        location2 = (path+'/')

        #Only activate print if the script fails and you want to see which file it failed on
        #print(name)

        # creating a pdf file object  
        pdfFileObj = open(location, 'rb') 

        # creating a pdf reader object 
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 

        # printing number of pages in pdf file 
        #print(pdfReader.numPages) 

        # creating a page object 
        pageObj = pdfReader.getPage(0) 

        # extracting text from page 

        ###
        #print(pageObj.extractText()) 
        textboi = pageObj.extractText()
        ###

        # closing the pdf file object 
        pdfFileObj.close() 

        #Turning data into dataframe
        data = io.StringIO(textboi)
        df = pd.read_csv(data, sep= '\n', header = 0)
        df.columns =['Info']

        #Finding the location of Invoice
        index = textboi.find('Invoice')
        index = (index+12)
        index2 = (index+6)
        try1 = textboi[index:index2]
        
        # The location of the business name changes constantly between buisnesses
        # The location of the Invoice number does not
        # This is a way to try to calculate where it will be off of information that doesn't move
        hi = df.index[df['Info'] == 'Invoice No:'].tolist()
        
        #Checking to make sure the file has the correct structure to be worked on
        if len(hi)>0:
            # Formatting the list into a single int
            hi = str(hi)
            hi = hi.replace('[','')
            hi = hi.replace(']','')
            hi = int(hi)

            #Two indexes needed for looping
            business_location = hi-4
            contact = hi-7
            hi = (hi-6)

            #If loop preparation for buninesses with specific location
            #This is customer specific
            business = str(df.loc[business_location])
            business = business.split(',')[0]
            business = business.replace('Info','')
            business = business.replace(' ','')

            #Business Name
            test = str(df.iloc[hi])
            test = test.replace('Info','')
            test = test[4:(len(test)-23)]
            
            #Checking to make sure the business name pulled isn't their address
            value = test[0:3]
            if True in [char.isdigit() for char in value]:
                if test != '3 Rooms Communications':
                    newname = hi-1
                    test = str(df.iloc[newname])
                    test = test.replace('Info','')
                    test = test[4:(len(test)-23)]
                    if test == 'Attn: Accounts Payable':
                        newname = newname-1
                        test = str(df.iloc[newname])
                        test = test.replace('Info','')
                        test = test[4:(len(test)-23)]
            
            
            ####
            #Test the business name
            #Business is the business location

            ####
            #try1 is the invoice number 
            ####
            if decision == 1:
                Filename = ('Invoice ' + try1 +'.pdf')
            elif decision ==2:
                Filename = (test + ' - Invoice ' + try1 +'.pdf')
            print(Filename)
            
            
            #Company Specific Naming Convention #1
            if test == 'Charter Communications Inc'and business == 'Coppell':
                    Filename = ('NTX ' + Filename)
            if test == 'Charter Communications Inc'and business == 'Dallas':
                    Filename = ('NTX ' + Filename)
            if test == 'Charter Communications Inc'and business == 'Austin':
                    Filename = ('CTX ' + Filename)
            ######################################
            #Company Specific Naming Convention #2
            
            #Setting the contact name for emailing purposes
            deg = str(df.iloc[contact])

            #list of others = Jeff Bow, Carlos Vega, 
            SH = 'Sandy Homan'
            CV = 'Carlos Vega'
            JB = 'Jeff Bow'
            ML = 'Montez Luis'
            DC = 'Dean Cole'

            if test == 'Suddenlink Communications':
                if SH in deg:
                    Filename = (test + ' SH' + ' - Invoice ' + try1 + '.pdf')
                if CV in deg:
                    Filename = (test + ' CV' + ' - Invoice ' + try1 + '.pdf')
                if JB in deg:
                    Filename = (test + ' JB' + ' - Invoice ' + try1 + '.pdf')
                if ML in deg:
                    Filename = (test + ' ML' + ' - Invoice ' + try1 + '.pdf')
                if DC in deg:
                    Filename = (test + ' DC' + ' - Invoice ' + try1 + '.pdf')
                #elif:
                    #Filename = (test + ' SH' + ' - Invoice' + try1 + '.pdf')
                    
            Filename = (location2 + Filename)
            Filename = Filename.replace('\n','')
            os.rename(location, Filename)
            del df
            del textboi
            
        else:
            print('Missing Invoice Number, Skippng File.')
    print('PDF Processing is Finished')
else:
    print('Script Needs Updating, Contact James Emmett')


# In[ ]:




