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

class FEYAP ():
    def __init__(self):
        self.p=subprocess.Popen(['C://Users//OzSea//yapproj//src//yap//yap','api'], shell=False, stdout=subprocess.PIPE)
        time.sleep(30)
        #subprocess.run(['C://Users//OzSea//yapproj//src//yap//yap','api'])
        self.localhost_yap = "http://localhost:8000/yap/heb/joint"
        self.headers = {'content-type': 'application/json'}
        # while (self.p.poll()):
        #     time.sleep(1)
            
        
        
    def FE_Yap(self,Text_path="101 - Atlas.docx"):
        poll = self.p.poll()
        doc = docx.Document(Text_path)
        result = [p.text for p in doc.paragraphs] 
        text=result[0]
        NumberOfwords=text.count(' ')#total number of words in the text
        data = json.dumps({'text': "{}  ".format(text)})  # input string ends with two space characters
        response = requests.get(url=self.localhost_yap, data=data, headers=self.headers)
        json_response = response.json()
        Table=json_response['dep_tree'].split('\n')
        
        
        #*********************************************************************
        #Feature1: What percentage of total verbs appear in past, present, future
        
        presentCount = json_response['md_lattice'].count('tense=BEINONI')
        pastCount = json_response['md_lattice'].count('tense=PAST')  
        futureCount = json_response['md_lattice'].count('tense=FUTURE')
        temp=presentCount+pastCount+futureCount
        Tense=dict({'Past':pastCount/temp,'Present':presentCount/temp,
                    'Future':futureCount/temp})
        #******************************************************************
        #What percentage of words are enclosed in quotation marks (markers of dialogue)
        Quotation_marks= [_.start() for _ in re.finditer('"', text)]
        Nqm=0
        for i in range(0,len(Quotation_marks),2):
            t=text[Quotation_marks[i]:Quotation_marks[i+1]+1]
            Nqm+=len(t.split(' '))
        NWqm=Nqm/NumberOfwords
        
        #******************************************************************
        Gufim=dict({'Me':0,'YouM':0,'YouF':0,'We':0,'They':0})
        Conjugation=dict({'Me':0,'YouM':0,'YouF':0,'We':0,'They':0})
        for i in range(len(Table)):
            
            match_list = re.findall(r'num=S|per=1', Table[i], re.IGNORECASE)
            Gufim['Me']+=len(match_list)==2
            match_list = re.findall(r'gen=M|num=S|per=2', Table[i], re.IGNORECASE)
            Gufim['YouM']+=len(match_list)==3
            match_list = re.findall(r'gen=F|num=S|per=2', Table[i], re.IGNORECASE)
            Gufim['YouF']+=len(match_list)==3
            match_list = re.findall(r'num=P|per=1', Table[i], re.IGNORECASE)
            Gufim['We']+=len(match_list)==2
            match_list = re.findall(r'per=3', Table[i], re.IGNORECASE)
            Gufim['They']+=len(match_list)==1
            
            match_list = re.findall(r'num=S|per=1|S_PRN', Table[i], re.IGNORECASE)
            Conjugation['Me']+=len(match_list)==3
            match_list = re.findall(r'gen=M|num=S|per=2|S_PRN', Table[i], re.IGNORECASE)
            Conjugation['YouM']+=len(match_list)==4
            match_list = re.findall(r'gen=F|num=S|per=2|S_PRN', Table[i], re.IGNORECASE)
            Conjugation['YouF']+=len(match_list)==4
            match_list = re.findall(r'num=P|per=1|S_PRN', Table[i], re.IGNORECASE)
            Conjugation['We']+=len(match_list)==3
            match_list = re.findall(r'per=3|S_PRN', Table[i], re.IGNORECASE)
            Conjugation['They']+=len(match_list)==2
            
        temp=sum(Gufim.values())
        for id in Gufim:
            Gufim[id]=Gufim[id]/temp
        temp=sum(Conjugation.values())
        for id in Conjugation:
            Conjugation[id]=Conjugation[id]/temp
        
        return (Tense,Gufim,Conjugation,NWqm)
            

