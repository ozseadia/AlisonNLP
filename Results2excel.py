# -*- coding: utf-8 -*-
"""
Created on Fri Sep 30 12:31:53 2022

@author: OzSea
"""
import re
def ResulysInitial(my_wb,R):
    my_sheet = my_wb.active
    my_sheet.cell(row = R, column = 1).value='Category'
    my_sheet.cell(row = R, column = 2).value='Absolut'
    my_sheet.cell(row = R, column = 3).value='Ratio'
    my_sheet.cell(row = R, column = 4).value='Words dictionary'
    my_sheet.cell(row = R, column = 5).value='YAP data'
    return(my_wb)

def Results(my_wb,R,GufimWords,Gufim,Total,Tens):
    my_sheet = my_wb.active
    
    for e in GufimWords:
        R+=1
        temp=str()
        for i in range(len(GufimWords[e])):
            A=(GufimWords[e][i].split(',')[1])
            temp+=((re.sub(r'[^\w]',' ',A)))
        print(e + ' ' + temp)
        my_sheet.cell(row = R, column = 1).value=Tens+''+e
        my_sheet.cell(row=R,column=2).value=Gufim[e]
        my_sheet.cell(row=R,column=3).value=round(Gufim[e]/Total,2)
        my_sheet.cell(row = R, column = 4).value=temp
        my_sheet.cell(row = R, column = 5).value=str(GufimWords[e])
    return (my_wb,R)

def Results1(my_wb,R,Tenses):
    my_sheet = my_wb.active
    for e in Tenses:
        R+=1
        my_sheet.cell(row = R, column = 1).value=e
        my_sheet.cell(row=R,column=2).value=Tenses[e]
        my_sheet.cell(row=R,column=3).value=round(Tenses[e]/sum(Tenses.values()),2)
        
    return (my_wb,R)


def Results2(my_wb,R,Score,text):
    my_sheet = my_wb.active
    for e in Score:
        R+=1
        my_sheet.cell(row = R, column = 1).value=text +' '+e
        my_sheet.cell(row=R,column=3).value=Score[e]
        
        
    return (my_wb,R)