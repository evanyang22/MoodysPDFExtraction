#!/usr/bin/env python
# coding: utf-8

# In[1]:


#How to use
#Change the number marked below to the number of folders you want to process
# Ex: want to process 10 folders, change the "2607" to 10 marked below
#Change the PDFTOTEXT_PATH variable to your pdftotext command line


# In[2]:


#country list
#list from here: https://pytutorial.com/python-country-list
countryList=['Afghanistan', 'Aland Islands', 'Albania', 'Algeria', 'American Samoa', 'Andorra', 'Angola', 'Anguilla', 'Antarctica', 'Antigua and Barbuda', 'Argentina', 'Armenia', 'Aruba', 'Australia', 'Austria', 'Azerbaijan', 'Bahamas', 'Bahrain', 'Bangladesh', 'Barbados', 'Belarus', 'Belgium', 'Belize', 'Benin', 'Bermuda', 'Bhutan', 'Bolivia, Plurinational State of', 'Bonaire, Sint Eustatius and Saba', 'Bosnia and Herzegovina', 'Botswana', 'Bouvet Island', 'Brazil', 'British Indian Ocean Territory', 'Brunei Darussalam', 'Bulgaria', 'Burkina Faso', 'Burundi', 'Cambodia', 'Cameroon', 'Canada', 'Cape Verde', 'Cayman Islands', 'Central African Republic', 'Chad', 'Chile', 'China', 'Christmas Island', 'Cocos (Keeling) Islands', 'Colombia', 'Comoros', 'Congo', 'Congo, The Democratic Republic of the', 'Cook Islands', 'Costa Rica', "Côte d'Ivoire", 'Croatia', 'Cuba', 'Curaçao', 'Cyprus', 'Czech Republic', 'Denmark', 'Djibouti', 'Dominica', 'Dominican Republic', 'Ecuador', 'Egypt', 'El Salvador', 'Equatorial Guinea', 'Eritrea', 'Estonia', 'Ethiopia', 'Falkland Islands (Malvinas)', 'Faroe Islands', 'Fiji', 'Finland', 'France', 'French Guiana', 'French Polynesia', 'French Southern Territories', 'Gabon', 'Gambia', 'Georgia', 'Germany', 'Ghana', 'Gibraltar', 'Greece', 'Greenland', 'Grenada', 'Guadeloupe', 'Guam', 'Guatemala', 'Guernsey', 'Guinea', 'Guinea-Bissau', 'Guyana', 'Haiti', 'Heard Island and McDonald Islands', 'Holy See (Vatican City State)', 'Honduras', 'Hong Kong', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Iran, Islamic Republic of', 'Iraq', 'Ireland', 'Isle of Man', 'Israel', 'Italy', 'Jamaica', 'Japan', 'Jersey', 'Jordan', 'Kazakhstan', 'Kenya', 'Kiribati', "Korea, Democratic People's Republic of", 'Korea, Republic of', 'Kuwait', 'Kyrgyzstan', "Lao People's Democratic Republic", 'Latvia', 'Lebanon', 'Lesotho', 'Liberia', 'Libya', 'Liechtenstein', 'Lithuania', 'Luxembourg', 'Macao', 'Macedonia, Republic of', 'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Mali', 'Malta', 'Marshall Islands', 'Martinique', 'Mauritania', 'Mauritius', 'Mayotte', 'Mexico', 'Micronesia, Federated States of', 'Moldova, Republic of', 'Monaco', 'Mongolia', 'Montenegro', 'Montserrat', 'Morocco', 'Mozambique', 'Myanmar', 'Namibia', 'Nauru', 'Nepal', 'Netherlands', 'New Caledonia', 'New Zealand', 'Nicaragua', 'Niger', 'Nigeria', 'Niue', 'Norfolk Island', 'Northern Mariana Islands', 'Norway', 'Oman', 'Pakistan', 'Palau', 'Palestinian Territory, Occupied', 'Panama', 'Papua New Guinea', 'Paraguay', 'Peru', 'Philippines', 'Pitcairn', 'Poland', 'Portugal', 'Puerto Rico', 'Qatar', 'Réunion', 'Romania', 'Russian Federation', 'Rwanda', 'Saint Barthélemy', 'Saint Helena, Ascension and Tristan da Cunha', 'Saint Kitts and Nevis', 'Saint Lucia', 'Saint Martin (French part)', 'Saint Pierre and Miquelon', 'Saint Vincent and the Grenadines', 'Samoa', 'San Marino', 'Sao Tome and Principe', 'Saudi Arabia', 'Senegal', 'Serbia', 'Seychelles', 'Sierra Leone', 'Singapore', 'Sint Maarten (Dutch part)', 'Slovakia', 'Slovenia', 'Solomon Islands', 'Somalia', 'South Africa', 'South Georgia and the South Sandwich Islands', 'Spain', 'Sri Lanka', 'Sudan', 'Suriname', 'South Sudan', 'Svalbard and Jan Mayen', 'Swaziland', 'Sweden', 'Switzerland', 'Syrian Arab Republic', 'Taiwan, Province of China', 'Tajikistan', 'Tanzania, United Republic of', 'Thailand', 'Timor-Leste', 'Togo', 'Tokelau', 'Tonga', 'Trinidad and Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'Turks and Caicos Islands', 'Tuvalu', 'Uganda', 'Ukraine', 'United Arab Emirates', 'United Kingdom', 'United States', 'United States Minor Outlying Islands', 'Uruguay', 'Uzbekistan', 'Vanuatu', 'Venezuela, Bolivarian Republic of', 'Viet Nam', 'Virgin Islands, British', 'Virgin Islands, U.S.', 'Wallis and Futuna', 'Yemen', 'Zambia', 'Zimbabwe']
countryList.append("U.S.A.")


# In[3]:


#generating random sample 
import random, os
random.seed(1111) #sets seed to 22 so its replicable

folderArray=[] #contains all folders
totalFolders=0
for folderName in os.scandir():
    if folderName.is_dir():
        totalFolders+=1
        folderArray.append(folderName)

# IMPORTANT- CHANGE 2607 to desired sample size, if you want to process every single folder then set it equal to number of folders   
sample=random.sample(range(0,len(folderArray)),2607)
selectedSample=[]
for ranNum in sample:
    selectedSample.append(folderArray[ranNum])

print(selectedSample)
for item in selectedSample:
    print(item)


# In[4]:


import os,sys,subprocess
import random
command= "C:\\Users\\evany_cdhq038\\OneDrive\\Desktop\\Economics_Research\\xpdf-tools-win-4.03\\bin64\\pdftotext.exe -layout C:\\Users\\evany_cdhq038\\OneDrive\\Desktop\\CompanyResearch\\"


def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start

masterArray=[]
logArray=[]

errorCounter=0
totalProcessedCounter=0
humanCounter=0
otherCounter=0

print(selectedSample)
for folderName in selectedSample:
    print(folderName)
    for fileName in os.scandir(folderName):
            if fileName.is_file() and fileName.name.endswith(".pdf"):
            #to investigate a specific PDF, uncomment the line below and comment line above
            #if fileName.is_file() and fileName.name.endswith(".pdf") and fileName.name=="PR_329110_185500_KO_20151022.pdf":
                testPath=fileName.path
                pdfPath=testPath
                #IMPORTANT- Change PDFTOTEXT_PATH variable below to the location of your pdftotext.exe
                PDFTOTEXT_PATH="C:\\Users\\evany_cdhq038\\OneDrive\\Desktop\\Economics_Research\\xpdf-tools-win-4.03\\bin64\\pdftotext.exe"
                pdfName=pdfPath[find_nth(pdfPath,'\\',2)+1:]
                try:
                    q = subprocess.Popen([PDFTOTEXT_PATH,'-layout',pdfPath,"-"], stdout=subprocess.PIPE, stderr=subprocess.PIPE) 
                    pdfText, err = q.communicate()
                except:
                    print('pdftotext failed')
                #PDFTOTEXT_PATH = '/usr/local/bin/pdftotext' 
                #processing decoded text
                decodedText=pdfText.decode("Latin1")
                decodedText=decodedText.replace("\x0c","")
                #decodedText=decodedText.replace("\x0c","\r\n") #replaces page break character with \r\n
                #creates challenges because sometimes \x0c needs to be treated as a \r\n and sometimes it needs to be not treated as one
                pdfArray=decodedText.split("\r\n") #splits using \r\n into array
                storageArray=[]

                counter=0
                containsRO= False 
                while counter<len(pdfArray):
                    if "Releasing Office" in pdfArray[counter]:
                        containsRO=True
                    if "JOURNALISTS" in pdfArray[counter]:
                        if "SUBSCRIBERS" in pdfArray[counter+1] or 'Client Service' in pdfArray[counter+1]:
                            newCounter=counter
                            newString=""
                            while newCounter>0:
                                if pdfArray[newCounter] == '':
                                    endCounter=newCounter
                                    break
                                newString+= pdfArray[newCounter].replace("/",":or:")+'/' # replaces / with :or: in text so we know it was digitally modified
                                newCounter-=1
                            storageArray.append(newString)
                            totalProcessedCounter+=1
                    counter+=1
                #print(storageArray)
                
                #processing things in storage array by splitting at "/Journalists"
                #catches the error where two journalists get appended to one
                for thing in storageArray:
                    if "/JOURNALISTS" in thing:
                        #print(thing)
                        storageArray.remove(thing) #removes the instance with two journalists
                        #processing instance into two strings
                        firstPart=thing[:thing.find("/JOURNALISTS")]+"/"
                        secondPart= thing[thing.find("/JOURNALISTS")+1:-1]+"/"#+1 removes the / at beginning, -1 removes the / at end
                        
                        #processing firstpart by removing subscribers
                        if "/SUBSCRIBERS" in firstPart:
                            #cutOff=firstPart.find("/SUBSCRIBERS")
                            #print(cutOff)
                            firstPart=firstPart[:firstPart.find("/SUBSCRIBERS")]
                            firstPart+="/"
                        #removing subscribers from second part
                        if "/SUBSCRIBERS" in secondPart:
                            secondPart=secondPart[:secondPart.find("/SUBSCRIBERS")]
                            secondPart+="/"


                        
                        storageArray.append(firstPart)
                        storageArray.append(secondPart)
                        totalProcessedCounter+=2 #adds two to counter because two got processed
                   # else:
                    #    print(thing)
                
                
                #post array processing
                #if contains RO, adds containsRO, it doesn't, adds "noRO"
                if containsRO:
                    for item in storageArray:
                        item=item +"containsRO/"+pdfName
                        masterArray.append(item)
                else:
                    for item in storageArray:
                        item=item +"noRO/" +pdfName
                        masterArray.append(item)
                
                #creating log file
                if not storageArray:
                    #print("This pdf is empty: "+ pdfName )
                    if containsRO:
                        pdfName="Manually Check: "+ pdfName
                        errorCounter+=1
                    logArray.append(pdfName)
    



# In[5]:


import spacy
from spacy import displacy
nlp = spacy.load("en_core_web_sm")
#processing strings
dictArray=[]
throwAway=[]

for item in masterArray:
    if "Releasing Office" in item: #skips loop if releaisng office is found in string
        continue 
    #replaces ", CFA" with""
    if ", CFA" in item:
        item=item.replace(", CFA","")
    if ",CFA" in item:
        item=item.replace(",CFA","")
    addThis={}
    containsRO= False 
    addThis["PDFName"]=item[item.rfind("/")+1:]
    item=item[:item.rfind("/")] #removes pdfname
    if "containsRO" in item:
        containsRO= True
    item=item[:item.rfind("/")] #removes containsRO or noRO
    
    #print(item[item.rfind("/")+1:])
    if len(item[item.rfind("/")+1:]) > 50: #processes string so if first line is >30 then it stops
        print("Chopped= "+ item[item.rfind("/")+1:])
        item=item[:item.rfind("/")]
        
    #chopping for each credit rating
    if item[item.rfind("/")+1:] == "for each credit rating." or item[item.rfind("/")+1:]=="please see the ratings tab on the issuer:or:entity page on www.moodys.com for the most updated credit rating":
        print("Chopped= "+ item[item.rfind("/")+1:])
        item=item[:item.rfind("/")]
    
    if len(item[item.rfind("/")+1:]) > 50: #processes string so if first line is >30 then it stops
        print("Chopped= "+ item[item.rfind("/")+1:])
        item=item[:item.rfind("/")]
    
    if len(item[item.rfind("/")+1:]) > 50: #processes string so if first line is >30 then it stops
        print("Chopped= "+ item[item.rfind("/")+1:])
        item=item[:item.rfind("/")]
    
    #containsRO
    if containsRO:
        addThis["AnalystName"]= item[item.rfind("/")+1:]
        item=item[:item.rfind("/")] #removes name
        addThis["Title"]=item[item.rfind("/")+1:]
        item=item[:item.rfind("/")] #removes title
        addThis["Group"]=item[item.rfind("/")+1:]
        item=item[:item.rfind("/")] #removes group
    else: #noRO removes city
        item=item[:item.rfind("/")] #removes city at the beginning
        addThis["AnalystName"]= item[item.rfind("/")+1:]
        item=item[:item.rfind("/")] #removes name
        addThis["Title"]=item[item.rfind("/")+1:]
        item=item[:item.rfind("/")] #removes title
        addThis["Group"]=item[item.rfind("/")+1:]
        item=item[:item.rfind("/")] #removes group

    
    #cleaning up name- replacing ? and double spaces with space
    addThis["AnalystName"]= addThis["AnalystName"].replace("?"," ")
    addThis["AnalystName"]= addThis["AnalystName"].replace("  "," ")
    addThis["AnalystName"]= addThis["AnalystName"].replace("  "," ")
    #extracts division
    
    if "Moody" in item:
        addThis["Division"]=item[item.rfind("/")+1:]
        item=item[:item.rfind("/")]
    
    
    
    #extracting address using number of / in remaining string
    #if remaining string is has more than x / in it, then process it accordingly to pattern
    #x/country/state/address/moodys investors service inc
    # london/united kingdom is different - x/uk/london/place/anotherplace/moodys service
    
    if item.count("/")>2: #if it has two or more /s
        if "United Kingdom" not in item:
            addThis["Address"]= item[item.rfind("/")+1:] +" " #extracts address
            item=item[:item.rfind("/")]
            addThis["Address"]+=item[item.rfind("/")+1:]+" " #extracts state
            item=item[:item.rfind("/")]
            addThis["Address"]+=item[item.rfind("/")+1:] #extracts country
            item=item[:item.rfind("/")]
            
        else:
            addThis["Address"]= item[item.rfind("/")+1:] +" "#extracts place1
            item=item[:item.rfind("/")]
            addThis["Address"]+=item[item.rfind("/")+1:] +" " #extracts place2
            item=item[:item.rfind("/")]
            addThis["Address"]+=item[item.rfind("/")+1:]+ " " #extracts city
            item=item[:item.rfind("/")]
            addThis["Address"]+=item[item.rfind("/")+1:]
            item=item[:item.rfind("/")]
    

    
    
    
    addThis["remainingString"]=item
    if '/' in addThis["remainingString"]:
        errorCounter+=1
        addThis["Errors"]="Potentially"
    else:
        addThis["Errors"]= "No"
    
    
    #adding Country column
    for country in countryList:
        if country in addThis["remainingString"]:
            addThis["Country"]= country
        if "Address" in addThis and country in addThis["Address"]:
            addThis["Country"]= country
    
    #labeling using spacy
    #doc.ents is a tuple
    doc = nlp(addThis["AnalystName"])
    if not doc.ents: # if tuple is empty
        addThis["NLPTag"]="Empty"
        otherCounter+=1
    for ent in doc.ents:
        addThis["NLPTag"]= ent.label_ 
        if ent.label_=="PERSON":
            humanCounter+=1
        else:
            otherCounter+=1
    
    print(addThis)
    
    #throwing away certain error of releasing office section
    if addThis["AnalystName"]=="250 Greenwich Street" and addThis["NLPTag"]=="FAC":
        # if name is that and nlp tag is fac, then throw out
        throwAway.append(addThis)
        print("Thrown Out")
 #   elif addThis["Errors"]== "Potentially":
  #      throwAway.append(addThis)
   #     print("Thrown OUt")
    elif addThis["AnalystName"]=="One Canada Square" and addThis["NLPTag"]=="GPE":
        throwAway.append(addThis)
        print("Thrown Out")
    else:
        dictArray.append(addThis)
    
    
    
    


# In[6]:


#measuring rough error rates
print("Error Counter / Total Processed: "+ str(errorCounter)+"/"+str(totalProcessedCounter))
print("NER Tagging Error Ratio (Humans to Other): " + str(humanCounter)+":"+str(otherCounter))
print("Throwaway length:"+str(len(throwAway)))
#in this case, minus 19 from NER tagging ratio because of misidentified names


# In[7]:


#creating  csv files
import csv
csv_columns=["PDFName","AnalystName","Group","Title","Division","Address","Country","remainingString","Errors","NLPTag"]

csv_file="MoodysAnalysts2.csv"
try:
    with open(csv_file,'w',newline="") as csvfile:
        writer=csv.DictWriter(csvfile,fieldnames=csv_columns)
        writer.writeheader()
        for data in dictArray:
            writer.writerow(data)
except IOError:
    print("IOError")
    


# In[8]:


#creating throwAway csv file
csv_columns=["PDFName","AnalystName","Group","Title","Division","Address","Country","remainingString","Errors","NLPTag"]

csv_file="throwAway2.csv"
try:
    with open(csv_file,'w',newline="") as csvfile:
        writer=csv.DictWriter(csvfile,fieldnames=csv_columns)
        writer.writeheader()
        for data in throwAway:
            writer.writerow(data)
except IOError:
    print("IOError")


# In[9]:


#creating logfile
with open("emptyFiles2.txt", "w") as txt_file:
    for line in logArray:
        txt_file.write(line + "\n") 


# In[ ]:




