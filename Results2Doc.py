# -*- coding: utf-8 -*-
"""
Created on Thu Sep  1 15:42:37 2022

@author: OzSea
"""
def Results(FileName,Sentiment_Score,RepeatWords3,RepeatWords,Tense,Gufim,Conjugation,NWqm):
    #Sentiment_Score,RepeatWords3,RepeatWords,Tense,Gufim,Conjugation,NWqm
    
    
    file = open(FileName+'Results.doc', "w",encoding='utf-8')
    #convert variable to string
    str = repr(Conjugation)
    file.write("Conjugation:" + str + "\n")
    str = repr(Gufim)
    file.write("Gufim:" + str + "\n")
    str = repr(Tense)
    file.write("Tense:" + str + "\n")
    
    str = repr(Sentiment_Score)
    file.write("Sentiment_Score:" + str + "\n")
    str = repr(NWqm)
    file.write("percentage of words that enclosed in quotation marks:" + str + "\n")
    
    str = repr(RepeatWords)
    file.write("RepeatWords1:" + str + "\n")
    str = repr(RepeatWords3)
    file.write("RepeatWords3:" + str + "\n")
     
    #close file
    file.close()


