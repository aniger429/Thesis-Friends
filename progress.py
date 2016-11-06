from openpyxl import load_workbook
from openpyxl import Workbook
import json

# used openpyxl#
wb = load_workbook(filename = 'Test.xlsx') #input
sheet = wb['Sheet1']
rown = 2
column = 'A'
ti = 0
id = 10000000;
userList = dict()
type(userList)
for row in sheet.rows:
    if (sheet.max_row < rown):
        break;
    else:
        d = sheet['A' + str(rown)]
        e = sheet['B' + str(rown)]
        f = sheet['D' + str(rown)]
        userList[f.value] = d.value
        rown += 1
        ti += 1
        id += 1
        
rown = 2
for row in sheet.rows:
    if (sheet.max_row < rown):
        break;
    else:
        e = sheet['B' + str(rown)]
        aa = e.value
        rown += 1
        if "@" in aa:
            splitList = aa.split()
            for i in range(0, len(splitList)):
                if "@" in splitList[i]:
                    for key in userList:
                        if key in splitList[i]:
                            splitList[i] = "@" + str(userList[key])
            aa = ""
            for i in range(0, len(splitList)):
                aa += " " + splitList[i]
                #User Anonymization by far#
        aa = aa.lower() #to lowercase#
        splitList = aa.split()
        for i in range(0, len(splitList)):
            standard = splitList[i]
            stopF = open("C:/Python27/stopwords.txt");
            contF = open("C:/Python27/contractions.txt");
            Fcount = len(stopF.readlines())
            Fcount2 = len(contF.readlines())
            stopF.close();
            contF.close();
            stopF = open("C:/Python27/stopwords.txt");
            contF = open("C:/Python27/contractions.txt");
            if "http://" in splitList[i] or "https://" in splitList[i]: #link#
                splitList[i] = ""
            if "#" in splitList[i]: #Hashtag#
                splitList[i] = ""
            if "&" in splitList[i] and ";" in splitList[i]: #escape HTML char#
                splitList[i] = ""
            if "." in splitList[i]:
                splitList[i] = splitList[i].replace(".", "")
            if "," in splitList[i]:
                splitList[i] = splitList[i].replace(",", "")
            if "!" in splitList[i]:
                splitList[i] = splitList[i].replace("!", "")
            if "?" in splitList[i]:
                splitList[i] = splitList[i].replace("?", "") #remove punctuation marks
            for j in range(0, Fcount2):
                cont = contF.readline()
                contWords = cont.split()
                if contWords[0] == splitList[i]:
                    if (len(contWords) > 1):
                        splitList[i] = contWords[1] + " " + contWords[2]
                    else:
                        splitList[i] = contWords[1] #contraction#
            for j in range(0, Fcount):       
                stopW = stopF.readline()
                stopW = stopW.strip()
                if stopW == splitList[i]:
                    splitList[i] = "" #stopword
            for j in range(0, len(standard)-2):
                if (standard[j] == standard[j+1] == standard[j+2]):
                    splitList[i] = splitList[i].replace(standard[j+2], "", 1) #standardize#
            stopF.close()
            contF.close()
        aa = ""
        for i in range(0, len(splitList)):
            aa += " " + splitList[i]
            #Removal of hashtags and link without exporting json file#
            #What would be the range of the emoticons? Do we need something like dictionary for emoticons as well? I need help with it#     
        print(aa)

            #for removal of contraction, how do we deal with things like peoples', human's, etc?#
            #split attach words to be done..#
            

        
                            

