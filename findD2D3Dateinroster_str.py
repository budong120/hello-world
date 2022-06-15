import openpyxl
from openpyxl import Workbook

dictOfMonth={'January':1, 'February':2, 'March':3, 'April':4, 'May':5, 'June':6, 'July':7, 'August':8, 'September':9, 'October':10, 'November':11, 'December':12, }


def findD2D3Date(sheet,s):
    l=str(yearMonth)
    nameofroster=sheet[0].value
    numDate=0
    for i in sheet:
        if i.value.find(s) != -1 :
            dateD2D3=i.column-2
            l=l+str(dateD2D3)+','
            numDate +=1
    n=len(l)
    return (nameofroster,s,numDate,l[:-1])

sourcePath=r'D:\\tmp\\202204April.xlsx'
wb=openpyxl.load_workbook(sourcePath)
activeSheet=wb.active
rosterScope=activeSheet['B7':'AG12']
yearOfSheet=wb['January']
yearOfActive=yearOfSheet['AH4']
monthOfAcitve=activeSheet['B4']
yearMonth=str(yearOfActive.value)+'/'+str(dictOfMonth[monthOfAcitve.value])+'/'
savename=str(yearOfActive.value)+str(dictOfMonth[monthOfAcitve.value])


savebook=Workbook()
savesheet=savebook.active

for i in rosterScope:
    print(findD2D3Date(i,"D2"))
    savesheet.append(findD2D3Date(i,"D2"))
    print(findD2D3Date(i,"D3"))
    savesheet.append(findD2D3Date(i,"D3"))

savebook.save("{0}.xlsx".format(savename+"_D2D3"))
