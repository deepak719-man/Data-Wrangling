import pandas as pd 
import numpy as np
from tkinter import *

def getvals():
    global var1#variable to read input file name from gui window
    global var3#variable to read input verbatim from gui window
    var1 = inputf.get()
    var3 = verbatim.get()
    root.destroy()
    
#gui code starts   
root = Tk()
inputf = StringVar()
outputf = StringVar()
verbatim = StringVar()

root.geometry('300x200')
#To take file name as input  
Label(root,text = "Enter File path:").pack()
Entry(root,textvariable = inputf).pack()
Label(root,text="").pack()

#Name of verbatim
Label(root,text= "Enter the Verbatim For code As 1:").pack()
Entry(root,textvariable = verbatim).pack()
Label(root,text="").pack()

Button(root,text = "Process",command=getvals).pack()

root.mainloop()
#gui code ends


xl = pd.ExcelFile(var1) 
raw = xl.parse('Input') # raw to store the input DataFrame

collapsed = pd.DataFrame(data={'ORDER':[],'SERIAL':[],'MENTION':[],'VERBATIM':[]})#Collapsed DataFrame to store Output
df1 = pd.DataFrame({'ORDER':[],'SERIAL':[],'MENTION':[],'VERBATIM':[]})#temporary DataFrame 
cframe = pd.DataFrame({'CODE':[],'VERBATIM':[]})#Dataframe to store CodeFrame


i=1 #Change the value of MENTION according to columns


for col in raw[[ 'Q8C1: ', 'Q8C2: ', 'Q8C3: ', 'Q8C4: ', 'Q8C5: ', 'Q8C6: ']]:
    df1[['SERIAL','VERBATIM']]=raw[['SERIAL',col]]
    df1.loc[:,'MENTION']=i
    collapsed = pd.concat([collapsed,df1],ignore_index=True,)
    i+=1
    
collapsed['ORDER'] = range(1,len(collapsed)+1) # Assign the values to ORDER according to the size of collapsed
collapsed.set_index('ORDER')
collapsed['VERBATIM']= collapsed['VERBATIM'].fillna("")
temp={}
#Creating a dictionary temp to store the count of appearance of a 'VERBATIM'
for val in collapsed.VERBATIM:
    if val in temp.keys():
        temp[val] = temp[val]+1
    else:
        temp[val] = 1


j = 2       
codeframe = {}      
#Loop for checking if the count is greater than 3. If so then assigning the 'CODE' the value of j 
for val in temp.keys():
    if val != var3 and val != '':
        if val == 'None' or val =='DK' or val == 'Don\'t know' or val == 'Dont know' or val=='Don’t know':
            codeframe[val] = 998
        elif val == 'Null':
            codeframe[val] = 999
        elif temp[val]>=3 :
            codeframe[val] = j
            j+=1
        else:
            codeframe[val] = 997
    elif val == var3:
        codeframe[val] = 1
            

collapsed['CODE'] = collapsed['VERBATIM'].map(codeframe)# mapping values of dictionary and store into column code 
cframe = pd.DataFrame({'CODE':[],'VERBATIM':[]},)


cframe = pd.DataFrame(columns=['VERBATIM','CODE'],data=list(codeframe.items()),index = range(len(codeframe)))#Create a DataFrame 'cframe' from dictionary 'codeframe' 
cframe.sort_values(by='CODE',axis=0,inplace=True)#sorting codeframe by the value of 'CODE'

#cf_final = {i:j for i,j in collapsed[['VERBATIM','CODE']].values if j!=999}

cframe = cframe[cframe['CODE']!=997]
cframe = cframe[cframe['CODE']!=998]
cframe = cframe[cframe['CODE']!=999]    

app = pd.DataFrame(columns=['VERBATIM','CODE'],data=[['Others',997]]) 
cframe = cframe.append(app)

with pd.ExcelWriter(var1) as op:
    raw.to_excel(op, sheet_name = 'Input') 
    collapsed.to_excel(op, sheet_name = 'Output',index=False)
    cframe.to_excel(op,sheet_name='CodeFrame',index=False)