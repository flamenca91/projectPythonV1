import docx
import re
import xlwings

# The dictionary holds  the tuples of child tags as Values. Here Keys are all the parent tags
docRelation = {"HRD":("HRS"), "HRS":("PRS"), "PRS":("URS","RISK"), "HTR":("HTP"), "HTP":("HRD", "HRS"), \
               "SDS":("BOLUS","ACE","AID"), "ACE":("PRS", "TBV", "DER"), "BOULUS":("PRS"), "AID":("PRS","DER"), \
               "SVAL":("BOLUS", "ACE", "AID"), "SVATR":("SVAL"), "UT":("UNIT"), "INS": ("UNIT")}      # to be created by the GUI

# The dictionary holds the tags as Keys and the file where they are found as a ValueError
docFile = {"HRD":"HDS_new_pump.docx", "HRS":"HRS_new_pump.docx", "HTP":"HTP_new_pump.docx", "HTR":"HTR_new_pump.docx", \
           "PRS":"PRS_new_pump.docx", "RISK":"RiskAnalysis_Pump.docx", "SDS":"SDS_New_pump_x04.docx", \
           "ACE":"SRS_ACE_Pump_X01.docx", "BOLUS":"SRS_BolusCalc_Pump_X04.docx", "SRS":"SRS_DosingAlgorithm_X03.docx", \
           "SVAL":"SVaP_new_pump.docx", "SVATR":"SVaTR_new_pump.docx", "UT":"SVeTR_new_pump.docx", "URS":"URS_new_pump.docx"}

docFileList = list(docFile.keys())                         # extracts all the tags into a list
filePath = "C:/Users/steph/OneDrive/Desktop/Docs_Project/" #location of the directory of all documents

def GetText(filename):                                 #Extracts all lines of any document and converts it into a list format (paragraph by paragraph)
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:                            #Where the collection is a set of paragraphs extracted from the document
        fullText.append(para.text)
    fullText = [ele for ele in fullText if ele.strip()]    # returns elements, does four loop  and if statement strips white spacesand storing back to list
    return fullText                                        # returns everything back into full text into list

def GetParentTags():
    for tag in docFileList:
        textList = GetText(filePath + docFile[tag])         #runsrpevous function
        index = 0
        ind = []
        for t in textList:
            if re.search('.*:' + re.escape(tag) + ':', t): #conduct regular expression search
                ind.append(index)
                tt = t
                y = re.findall('\S*:' + re.escape(tag) + ':\S*', t)
                tt = tt.replace(y[0], '')
                tt = tt.strip()
                print(y[0])


def GetChildTags():
    for tag in docFileList:
        textList = GetText(filePath + docFile[tag])  # runsrpevous function
        index = 0
        ind = []
        for t in textList:
            if re.search('.*:' + re.escape(tag) + ':', t):  # conduct regular expression search
                ind.append(index)
                tt = t
                y = re.findall('\s\[.+\]\s', t)
                if len(y) != 0:
                    print(y[0])

GetParentTags()
GetChildTags()

'''
#def main():
txtLst = getText('C:/Users/lrey/Desktop/Docs_Project/HDS_new_pump.docx')
index = 0
ind = []

for t in txtLst:
    if re.search('.*:HRD:', t):
        ind.append(index)
        tt=t
        y = re.findall('\S*:HRD:\S*', t)
        z = re.findall('\S*:HRS:\S*', t)
        tt = tt.replace(y[0], '')
        tt = tt.replace(z[0], '')
        tt = tt.strip()

        print(tt)
        #print(t)
        print(y)
        print(z)
    index = index + 1

print(ind)
#print(txtLst)

'''
