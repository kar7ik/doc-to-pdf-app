# -*- coding: utf-8 -*-
"""
@author: kartik

"""

import os
import comtypes.client


path = os.walk(".")
output_dir ="PDF/"
pFiles = []


for root, directories, files in path:
    for file in files:
        pFiles.append(file)
        
print(pFiles)

if not os.path.isdir(output_dir):
    os.mkdir(output_dir)
    print("PDF folder created")
        
for index,pFile in enumerate(pFiles,1):        

    wdFormatPDF = 17
    
    in_file = os.path.abspath(f"{pFile}")
    print(in_file)
    out_file = os.path.abspath(f"./PDF/{pFile}.pdf")
    print(out_file)
    
    try:
    
        if 'doc' in in_file:
            print("---------------------------------------------------------")
            print(index)
            print("START")
            word = comtypes.client.CreateObject('Word.Application')
            print("APP")
            doc = word.Documents.Open(in_file)
            print("OPEN")
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
            print("SAVE")
            doc.Close()
            print("CLOSE")
            word.Quit()
            print("QUIT")
           
    except Exception as E:
        print("------------------------------------------------------------\n")
        
        print("Error:", E)
        
    finally:
        print("Final Quit")
            
        
        
