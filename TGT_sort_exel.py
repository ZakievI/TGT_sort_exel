import openpyxl as opxl
import numpy as np

def month_to_number(month):
    months = {
        "январь": ".01",
        "февраль": ".02",
        "март": ".03",
        "апрель": ".04",
        "май": ".05",
        "июнь": ".06",
        "июль": ".07",
        "август": ".08",
        "сентябрь": ".09",
        "октябрь": ".10",
        "ноябрь": ".11",
        "декабрь": ".12"
    }
    
    return months.get(month.lower(), "Некорректное название месяца")

def sort_exels(file,file_new):
    book=opxl.open(file,read_only=True)
    sheet=book.active
    data=[]
    info=[sheet[4][8].value,sheet[5][8].value,sheet[6][8].value,sheet[7][8].value, sheet[8][8].value,sheet[9][8].value,sheet[10][8].value,sheet[11][8].value,sheet[12][8].value,sheet[13][8].value,sheet[14][8].value,sheet[15][8].value,sheet[16][8].value,sheet[17][8].value,sheet[18][8].value]
    j=5
    while (sheet[j][1].value!=None):
        month_y=sheet[j][1].value
        size=len(month_y)
        month=month_to_number(month_y[0:size-6])
        year=month_y[size-4:]
        date=[] # создаем пустой массив 
        i=9
        while(sheet[3][i].value!=None):
            if(sheet[j-1][i].value==None):
                i=i+1
                continue
            date.append(list([(str(sheet[3][i].value)+month+"."+year), sheet[j-1][i].value, sheet[j][i].value, sheet[j+1][i].value, sheet[j+2][i].value, sheet[j+3][i].value, sheet[j+4][i].value, sheet[j+5][i].value, sheet[j+6][i].value, sheet[j+7][i].value, sheet[j+8][i].value, sheet[j+9][i].value, sheet[j+10][i].value, sheet[j+11][i].value, sheet[j+12][i].value, sheet[j+13][i].value])) # получаем данные 
            i=i+1
        data.append(list(date))
        print(j)
        j=j+15

    data.reverse()
    book_new=opxl.Workbook()
    sheet_new=book_new.active

    for i in range(0,15):
        sheet_new.cell(row=1,column=2+i).value=info[i]

    i=0
    for data_new in data:
        for date_new in data_new:
            for j in range(0,15):
                sheet_new.cell(row=2+i,column=1+j).value=date_new[j]
            i=i+1
    book_new.save(file_new)
    book_new.close()
    book.close()


sort_exels("Скв. 388 (01.11.2016-30.11.2023).xlsx","Скв. 388 (01.11.2016-30.11.2023)_new.xlsx")
