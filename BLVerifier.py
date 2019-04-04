from PyPDF2 import PdfFileReader
from glob import *
import os
import pandas as pd
from tkinter import *

#Import Packing List

def verifier():
    #data = pd.ExcelFile(r'C:\Users\admin\PycharmProjects\draftverifier\draftverifier.xlsx')
    #df = data.parse('Sheet1')
    df = pd.read_excel(r'C:\Users\admin\PycharmProjects\draftverifier\draftverifier.xlsx', sheet_name='Sheet1', dtype=str)
    containerpk = df[df.columns[0]].tolist()
    containerpk.append(df.columns[0])
    sealpk1 = df[df.columns[1]].tolist()
    sealpk1.append(df.columns[1])
    sealpk=[r.replace('/','') for r in sealpk1]
    balepk = df[df.columns[2]].tolist()
    balepk.append(str(df.columns[2]))
    weightpk = df[df.columns[3]].tolist()
    weightpk.append(str(df.columns[3]))
    dic_pk = {containerpk[i]:[balepk[i],weightpk[i]] for i in range(len(containerpk))} #dictionary for entire table
    dic_seal_pk = {containerpk[i]:[sealpk[i]] for i in range(len(containerpk))}
    #print(containerpk)
    #print(sealpk)

    #Import draft
    list_of_files=glob(r'C:\Users\admin\PycharmProjects\draftverifier\*.pdf')
    a = max(list_of_files, key=os.path.getctime)
    #a=open(r'C:\Users\admin\PycharmProjects\MEDUMT131882_DRAFT_BL.PDF','rb')
    filereader = PdfFileReader(a)
    numberofpage = filereader.getNumPages()
    b=''
    for i in range(1,numberofpage):
        pageObj = filereader.getPage(i)
        b+=pageObj.extractText()
    #pageObj=filereader.getPage(1)
    #b=pageObj.extractText()
    print(b)
    def string_find(string, sub):
        start=0
        containerN = []
        sealN = []
        while True:
            start=string.find(sub, start)
            if start==-1:
                break
            containerN.append(str(start-11))
            sealN.append(str(start+12))
            start += len(sub)
        return [containerN,sealN]

    def weight_finder(string, sub):
        start = 0
        weight_position = []
        weight = []
        while True:
            start = string.find(sub, start)
            if start == -1:
                break
            weight_position.append(start)
            start += len(sub)
        for i in range(len(weight_position)-1):
            weight.append(str(int(float(b[weight_position[i]+5:weight_position[i]+11].replace(',','')))))
        return weight

    def weight_finder2(string, sub):
        start = 0
        weight_position = []
        weight = []
        while True:
            start = string.find(sub, start)
            if start == -1:
                break
            weight_position.append(start)
            start += len(sub)
        for i in range(len(weight_position)-1):
            weight.append(str(int(float(b[weight_position[i]-11:weight_position[i]-4].replace(',','')))))
        return weight


    def bale_finder(string,sub):
        start = 0
        bale_position = []
        bale = []
        while True:
            start = string.find(sub, start)
            if start == -1:
                break
            bale_position.append(start)
            start += len(sub)
        for i in range(len(bale_position)):
            bale.append(str(int(float(b[bale_position[i]-3:bale_position[i]]))))
        return bale

    c = string_find(b, 'SEAL NUMBER:')

    container = []
    seal = []
    L = c[1]


    for i in range(len(c[0])):
        container_tem = b[int(c[0][i]):int(c[0][i]) + 11]
        container.append(container_tem)

    try:
        for i in range(len(L)):
            start = 6
            while b[int(L[i])+start] != "'":
                start += 1

            seal.append(b[int(L[i]):int(L[i]) + start-2])
        dic_seal = {container[i]: [seal[i]] for i in range(len(container))}
        print(seal)
    except:
        z.set('error occurs.')

    try:
        try:
            f = weight_finder(b, 'KGS.')
            g = bale_finder(b, 'BALE(S)')
            dic_draft = {container[i]: [g[i], f[i]] for i in range(len(container))}  # create the dictionary for draft
        except:
            f = weight_finder2(b, 'KGS.')
            g = bale_finder(b, 'BALE(S)')
            dic_draft = {container[i]: [g[i], f[i]] for i in range(len(container))}

    except:
        x.set('error occurs.')
    #print(f)
    #print(g)


    #Comparison
    d = container.copy()
    error_container = ''
    for i in range(len(containerpk)):
        if containerpk[i] in container:
            d.remove(containerpk[i])
        else:
            error_container += ' ' + containerpk[i]
    if d == []:
        print('There is no mistake in container number.')
        q.set('There is no mistake in container number.')
        p.set(None)
    else:
        print('There is discrepancy in container# between: ', d, '  Wrong number:', error_container)
        p.set(d)
        q.set(error_container)

    correct_seal = dict(dic_seal_pk)
    error_seal = {}
    for k, v in dic_seal_pk.items():
        if dic_seal.get(k) != None:
            if set(dic_seal.get(k)) != set(v):
                error_seal.update({k:dic_seal.get(k)})
            else:
                del correct_seal[k]
    if correct_seal == {}:
        z.set('There is no mistake in seal number')
        m.set(None)
    else:
        m.set(error_seal)
        z.set(correct_seal)

    correct = dict(dic_pk)
    errorSBW = {}
    for k, v in dic_pk.items():
        if dic_draft.get(k) != None:
            if set(dic_draft.get(k)) != set(v):
                errorSBW.update({k: dic_draft.get(k)})
            else:
                del correct[k]
    if correct == {}:
        print('There is no mistake in bale/weight.')
        x.set('There is no mistake in bale/weight.')
        y.set(None)
    else:
        print('There is discrepancy in seal# between: ', correct, '  Wrong number:', errorSBW)
        x.set(correct)
        y.set(errorSBW)

class APP:
    def __init__(self, master,p,q,x,y,m,z):
        self.master = master
        self.p = p
        self.q = q
        self.x = x
        self.y = y
        self.m = m
        self.z = z
    def display(self):
        self.truecontainer = Entry(self.master, textvariable = self.q, width = 140).grid(row=1, column=0)
        self.wrongcontainer = Entry(self.master, textvariable=self.p, width=140).grid(row=1, column=1)
        self.trueBALE = Entry(self.master, textvariable=self.x, width=140).grid(row=3, column=0)
        self.wrongBALE = Entry(self.master, textvariable=self.y, width=140).grid(row=3, column=1)
        self.trueseal = Entry(self.master, textvariable=self.z, width=140).grid(row=5, column=0)
        self.wrongseal = Entry(self.master, textvariable=self.m, width=140).grid(row=5, column=1)
        self.press = Button(self.master, text='RUN SEAL', command=verifier).grid(row=7)

root = Tk()
root.title('Draft Verifier')
frame = Frame(root)
p = StringVar()
q = StringVar()
x = StringVar()
y = StringVar()
m = StringVar()
z = StringVar()
Label(frame, text='CORRECT CONTAINER').grid(row=0, column=0)
Label(frame, text='WRONG CONTAINER').grid(row=0, column=1)
Label(frame, text='CORRECT BALE/WEIGHT').grid(row=2, column=0)
Label(frame, text='WRONG BALE/WEIGHT').grid(row=2, column=1)
Label(frame, text='CORRECT SEAL').grid(row=4, column=0)
Label(frame, text='WRONG SEAL').grid(row=4, column=1)
frame.pack(padx=60, pady=20)
my_gui = APP(frame,p,q,x,y,m,z)
my_gui.display()
mainloop()