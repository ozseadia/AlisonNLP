# -*- coding: utf-8 -*-
"""
Created on Fri Aug  5 17:31:42 2022

@author: OzSea
"""
import time
import numpy as np 
import requests
import json
import docx
import re
import subprocess
import threading
from transformers import AutoTokenizer, AutoModel, pipeline

class FEYAP ():
    def __init__(self):
        self.p=subprocess.Popen(['C://Users//OzSea//yapproj//src//yap//yap','api'], shell=False, stdout=subprocess.PIPE)
        time.sleep(30)
        #subprocess.run(['C://Users//OzSea//yapproj//src//yap//yap','api'])
        self.localhost_yap = "http://localhost:8000/yap/heb/joint"
        self.headers = {'content-type': 'application/json'}
        # while (self.p.poll()):
        #     time.sleep(1)
        self.sentiment_analysis = pipeline(
	    "sentiment-analysis",
	    model="avichr/heBERT_sentiment_analysis",
	    tokenizer="avichr/heBERT_sentiment_analysis",
	    return_all_scores = True)
            
    def Text2sentenc(self,text):
        data = json.dumps({'text': "{}  ".format(text)})  # input string ends with two space characters
        response = requests.get(url=self.localhost_yap, data=data, headers=self.headers)
        json_response = response.json()
        return(json_response)
    
    def RepetWord1(self,TT):
        
        for i in range(len(TT)):
            TT[i]=TT[i].split('\t')
        n=0
        for i in range(len(TT)-1):
            
            if len(TT[i+1])>=3 and len(TT[i])>=3:
                #print(i)
                try:
                    n+=(TT[i][2]==TT[i+1][2]) 
                except:
                    print(TT[i+1])
        return(n) 
    
    def RepetWord3(self,text):
        Temp=re.sub(r'[^\w]', ' ', text) # remove all symbols from string
        Temp = list(filter(None, Temp.split(' ')))
        Words_List=list(set(Temp))
        words_dict = dict.fromkeys(Words_List,0)
        for key in words_dict:
            words_dict[key]=(re.sub(r'[^\w]', ' ', text)).count(' '+key+' ')
            
        return(words_dict)
    
    def CloseYap(self):
        self.p.terminate()
    
    def Sentiment(self,text):
        S=self.sentiment_analysis(text.split('.'))
        Score=dict({'negative':0,'positive':0,'neutral':0})
        for e in S:
            for i in range(len(e)):
                Score[list(e[i].values())[0]]+=list(e[i].values())[1]/len(S)
        return(Score)
        
    
    
    def FE_Yap(self,Text_path="101 - Atlas.docx"):
        
        doc = docx.Document(Text_path)
        
        result = [p.text for p in doc.paragraphs] 
        result[0]=re.sub("\(.*?\)","",result[0])
        result[0]=re.sub("[\[\]<>()\\/]","",result[0])
        RepeatWords3=self.RepetWord3(result[0])
        Sentiment_Score=self.Sentiment(result[0])
        NumberOfwords=result[0].count(' ')#total number of words in the text
        TEXT=result[0].split('.')
        presentCount=0
        pastCount=0
        futureCount=0
        Gufim=dict({'Me':0,'YouMs':0,'YouFs':0,'Youp':0,'We':0,'They':0,'He':0,'She':0})
        Conjugation=dict({'Me':0,'YouMs':0,'YouFs':0,'Youp':0,'We':0,'They':0,'He':0,'She':0})
        
        NWqm=0
        RepeatWords=0
        #******************************************************************
        #What percentage of words are enclosed in quotation marks (markers of dialogue)
        Quotation_marks= [_.start() for _ in re.finditer('"', result[0])]
        Nqm=0
        print(Quotation_marks)
        if Quotation_marks:
            for i in range(0,len(Quotation_marks),2):
                t=result[0][Quotation_marks[i]:Quotation_marks[i+1]+1]
                Nqm+=len(t.split(' '))
            NWqm+=Nqm/NumberOfwords
        
        #******************************************************************
        
        for text in TEXT:
            data = json.dumps({'text': "{}  ".format(text)})  # input string ends with two space characters
            response = requests.get(url=self.localhost_yap, data=data, headers=self.headers)
            json_response = response.json()
            RepeatWords+=self.RepetWord1(json_response['dep_tree'].split('\n'))
            Table=json_response['dep_tree'].split('\n')
            #print(Table)
            
            
            
            #*********************************************************************
            #Feature1: What percentage of total verbs appear in past, present, future
            
            presentCount+= json_response['md_lattice'].count('tense=BEINONI')
            pastCount += json_response['md_lattice'].count('tense=PAST')  
            futureCount += json_response['md_lattice'].count('tense=FUTURE')
            temp=presentCount+pastCount+futureCount
            #temp=1
            Tense=dict({'Past':pastCount/temp,'Present':presentCount/temp,
                        'Future':futureCount/temp})
            

            
            for i in range(len(Table)):
                #print(type(Table[i]))
                if not(re.findall(r'suf_per', Table[i], re.IGNORECASE)):
                    if not (re.findall(r'S_PRN', Table[i], re.IGNORECASE)):
                        match_list = re.findall(r'num=S|per=1', Table[i], re.IGNORECASE)
                        Gufim['Me']+=len(match_list)==2
                        #if len(match_list)==2:
                        #     print(Table[i].split('\t'))
                        match_list = re.findall(r'gen=M|num=S|per=2', Table[i], re.IGNORECASE)
                        Gufim['YouMs']+=len(match_list)==3
                        match_list = re.findall(r'gen=F|num=S|per=2', Table[i], re.IGNORECASE)
                        Gufim['YouFs']+=len(match_list)==3
                        match_list = re.findall(r'num=P|per=2', Table[i], re.IGNORECASE)
                        Gufim['Youp']+=len(match_list)==2
                        
                        match_list = re.findall(r'num=P|per=1', Table[i], re.IGNORECASE)
                        Gufim['We']+=len(match_list)==2
                        match_list = re.findall(r'num=P|per=3', Table[i], re.IGNORECASE)
                        Gufim['They']+=len(match_list)==2
                        #print(match_list)
                        #if len(match_list)==2:
                        #    print(match_list)
                        #    print(Table[i].split('\t'))
                        match_list = re.findall(r'gen=M|num=S|per=3', Table[i], re.IGNORECASE)
                        Gufim['He']+=len(match_list)==3
                        match_list = re.findall(r'gen=F|num=S|per=3', Table[i], re.IGNORECASE)
                        Gufim['She']+=len(match_list)==3
                    match_list = re.findall(r'num=S|per=1|S_PRN', Table[i], re.IGNORECASE)
                    Conjugation['Me']+=len(match_list)==3
                    match_list = re.findall(r'gen=M|num=S|per=2|S_PRN', Table[i], re.IGNORECASE)
                    Conjugation['YouMs']+=len(match_list)==4
                    match_list = re.findall(r'gen=F|num=S|per=2|S_PRN', Table[i], re.IGNORECASE)
                    Conjugation['YouFs']+=len(match_list)==4
                    match_list = re.findall(r'num=P|per=2|S_PRN', Table[i], re.IGNORECASE)
                    Conjugation['Youp']+=len(match_list)==3
                    match_list = re.findall(r'num=P|per=1|S_PRN', Table[i], re.IGNORECASE)
                    Conjugation['We']+=len(match_list)==3
                    match_list = re.findall(r'per=3|S_PRN', Table[i], re.IGNORECASE)
                    Conjugation['They']+=len(match_list)==2
                    
                    match_list = re.findall(r'gen=M|num=S|per=3|S_PRN', Table[i], re.IGNORECASE)
                    Conjugation['He']+=len(match_list)==4
                    match_list = re.findall(r'gen=F|num=S|per=3|S_PRN', Table[i], re.IGNORECASE)
                    Conjugation['She']+=len(match_list)==4

        #temp=1
        temp=sum(Gufim.values())
        for id in Gufim:
            Gufim[id]=Gufim[id]/temp
        temp=sum(Conjugation.values())
        #temp=1
        for id in Conjugation:
            Conjugation[id]=Conjugation[id]/temp
        
        return (Sentiment_Score,RepeatWords3,RepeatWords,Tense,Gufim,Conjugation,NWqm)
            

