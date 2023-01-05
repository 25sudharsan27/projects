from tkinter import *
from openpyxl import load_workbook
class App:
    def __init__(self,root):
        self.co1="#2B4257"
        self.root=root
        self.root.geometry("700x700")
        self.root.title("Billing Software")

        self.cna=StringVar()
        self.ca1=StringVar()
        self.ca2=StringVar()
        self.ca3=StringVar()
        self.eekn=IntVar()
        self.eekd=StringVar()
        self.eekdi=StringVar()
        self.eekp=IntVar()
        self.eekv=IntVar()
        self.da=IntVar()

        self.a1=StringVar()
        self.a2= IntVar()
        self.a3= IntVar()
        self.a4= IntVar()


        self.b1 = StringVar()
        self.b2 = IntVar()
        self.b3 = IntVar()
        self.b4 = IntVar()

        self.c1 = StringVar()
        self.c2 = IntVar()
        self.c3 = IntVar()
        self.c4 = IntVar()

        self.d1 = StringVar()
        self.d2 = IntVar()
        self.d3 = IntVar()
        self.d4 = IntVar()

        self.e1 = StringVar()
        self.e2 = IntVar()
        self.e3 = IntVar()
        self.e4 = IntVar()

        self.f1 = StringVar()
        self.f2 = IntVar()
        self.f3 = IntVar()
        self.f4 = IntVar()

        self.g1 = StringVar()
        self.g2 = IntVar()
        self.g3 = IntVar()
        self.g4 = IntVar()

        self.h1 = StringVar()
        self.h2 = IntVar()
        self.h3 = IntVar()
        self.h4 = IntVar()

        self.i1 = StringVar()
        self.i2 = IntVar()
        self.i3 = IntVar()
        self.i4 = IntVar()



        frame1=Frame(self.root).pack(fill=X)
        label1=Label(frame1,text="Billing Software",font="fiver 15 bold",bg="#2B4257",fg="white").pack(fill=X)
        self.frame2=Frame(self.root,bg=self.co1)
        self.frame2.place(x=40,y=100,height=120,width=300)

        label1=Label(self.frame2,text="(To) Details",fg="red",bg=self.co1,font="fiver 10 bold")
        label1.grid(row=1,column=1)
        label2=Label(self.frame2,text="Company name",font="verdana 8 bold",fg="white",bg=self.co1)
        label2.grid(row=2,column=1)
        ent1=Entry(self.frame2,textvariable=self.cna)
        ent1.grid(row=2,column=2)
        label3=Label(self.frame2,text="Address 1",font="verdana 8 bold",fg="white",bg=self.co1)
        label3.grid(row=3,column=1)
        ent2=Entry(self.frame2,textvariable=self.ca1)
        ent2.grid(row=3,column=2)
        label4=Label(self.frame2,text="Address 2",font="verdana 8 bold",fg="white",bg=self.co1)
        label4.grid(row=4,column=1)
        ent3=Entry(self.frame2,textvariable=self.ca2)
        ent3.grid(row=4,column=2)
        label5=Label(self.frame2,text="Address 3",font="verdana 8 bold",fg="white",bg=self.co1)
        label5.grid(row=5,column=1)
        ent4=Entry(self.frame2,textvariable=self.ca3)
        ent4.grid(row=5,column=2)
        self.frame3=Frame(self.root,bg=self.co1)
        self.frame3.place(x=350,y=100,height=150,width=300)
        lab1=Label(self.frame3,text="Eki calculation kijam 73",bg=self.co1,fg="red",font="verdana 8 bold")
        lab1.grid(row=1,column=1)
        lab2=Label(self.frame3,text="Laat Number",font="verdana 8 bold",fg="white",bg=self.co1)
        lab2.grid(row=2,column=1)
        ent1 = Entry(self.frame3, textvariable=self.eekn)
        ent1.grid(row=2, column=2)
        lab3 = Label(self.frame3, text="PettiVaravu date",font="verdana 8 bold",fg="white",bg=self.co1)
        lab3.grid(row=3, column=1)
        ent2 = Entry(self.frame3, textvariable=self.eekd)
        ent2.grid(row=3, column=2)
        lab4= Label(self.frame3, text="Diiniyar",font="verdana 8 bold",fg="white",bg=self.co1)
        lab4.grid(row=4, column=1)
        ent3 = Entry(self.frame3, textvariable=self.eekdi)
        ent3.grid(row=4, column=2)
        lab5 = Label(self.frame3, text="petti",font="verdana 8 bold",fg="white",bg=self.co1)
        lab5.grid(row=5, column=1)
        ent4 = Entry(self.frame3, textvariable=self.eekp)
        ent4.grid(row=5, column=2)
        lab6 = Label(self.frame3, text="varavu kilo",font="verdana 8 bold",fg="white",bg=self.co1)
        lab6.grid(row=6, column=1)
        ent5 = Entry(self.frame3, textvariable=self.eekv)
        ent5.grid(row=6, column=2)

        self.frame4=Frame(self.root,bg=self.co1)
        self.frame4.place(x=20,y=225,height=400,width=320)
        self.co2="#36454F"
        frame1a=Frame(self.frame4,bg=self.co2)
        frame1a.place(x=10,y=40,height=20,width=300)
        la1=Label(frame1a,text="Date",font="times 8 bold",bg=self.co2,fg="white",padx=20)
        la1.grid(row=1,column=1)
        la2 = Label(frame1a, text="kijam", font="times 8 bold", bg=self.co2, fg="white",padx=20)
        la2.grid(row=1, column=2)
        la3 = Label(frame1a, text="eeki", font="times 8 bold", bg=self.co2, fg="white",padx=20)
        la3.grid(row=1, column=3)
        la4 = Label(frame1a, text="count", font="times 8 bold", bg=self.co2, fg="white",padx=20)
        la4.grid(row=1, column=4)

        self.frame1b=Frame(self.frame4,bg=self.co1)
        self.frame1b.place(x=10,y=70,height=700,width=320)
        a1=Entry(self.frame1b,width=8,textvariable=self.a1)
        a1.grid(row=1,column=1,padx=10,pady=5)
        a2 = Entry(self.frame1b,width=9,textvariable=self.a2)
        a2.grid(row=1, column=2,padx=10,pady=5)
        a3 = Entry(self.frame1b,width=7,textvariable=self.a3)
        a3.grid(row=1, column=3,padx=10,pady=5)
        a4 = Entry(self.frame1b,width=8,textvariable=self.a4)
        a4.grid(row=1, column=4,padx=10,pady=5)

        b1 = Entry(self.frame1b, width=8,textvariable=self.b1)
        b1.grid(row=2, column=1, padx=10,pady=5)
        b2 = Entry(self.frame1b, width=9,textvariable=self.b2)
        b2.grid(row=2, column=2, padx=10,pady=5)
        b3 = Entry(self.frame1b, width=7,textvariable=self.b3)
        b3.grid(row=2, column=3, padx=10,pady=5)
        b4 = Entry(self.frame1b, width=8,textvariable=self.b4)
        b4.grid(row=2, column=4, padx=10,pady=5)

        c1 = Entry(self.frame1b, width=8,textvariable=self.c1)
        c1.grid(row=3, column=1, padx=10,pady=5)
        c2 = Entry(self.frame1b, width=9,textvariable=self.c2)
        c2.grid(row=3, column=2, padx=10,pady=5)
        c3 = Entry(self.frame1b, width=7,textvariable=self.c3)
        c3.grid(row=3, column=3, padx=10,pady=5)
        c4 = Entry(self.frame1b, width=8,textvariable=self.c4)
        c4.grid(row=3, column=4, padx=10,pady=5)

        d1 = Entry(self.frame1b, width=8,textvariable=self.d1)
        d1.grid(row=4, column=1, padx=10,pady=5)
        d2 = Entry(self.frame1b, width=9,textvariable=self.d2)
        d2.grid(row=4, column=2, padx=10,pady=5)
        d3 = Entry(self.frame1b, width=7,textvariable=self.d3)
        d3.grid(row=4, column=3, padx=10,pady=5)
        d4 = Entry(self.frame1b, width=8,textvariable=self.d4)
        d4.grid(row=4, column=4, padx=10,pady=5)

        e1 = Entry(self.frame1b, width=8,textvariable=self.e1)
        e1.grid(row=5, column=1, padx=10,pady=5)
        e2 = Entry(self.frame1b, width=9,textvariable=self.e2)
        e2.grid(row=5, column=2, padx=10,pady=5)
        e3 = Entry(self.frame1b, width=7,textvariable=self.e3)
        e3.grid(row=5, column=3, padx=10,pady=5)
        e4 = Entry(self.frame1b, width=8,textvariable=self.e4)
        e4.grid(row=5, column=4, padx=10,pady=5)

        f1 = Entry(self.frame1b, width=8,textvariable=self.f1)
        f1.grid(row=6, column=1, padx=10,pady=5)
        f2 = Entry(self.frame1b, width=9,textvariable=self.f2)
        f2.grid(row=6, column=2, padx=10,pady=5)
        f3 = Entry(self.frame1b, width=7,textvariable=self.f3)
        f3.grid(row=6, column=3, padx=10,pady=5)
        f4 = Entry(self.frame1b, width=8,textvariable=self.f4)
        f4.grid(row=6, column=4, padx=10,pady=5)

        g1 = Entry(self.frame1b, width=8,textvariable=self.g1)
        g1.grid(row=7, column=1, padx=10,pady=5)
        g2 = Entry(self.frame1b, width=9,textvariable=self.g2)
        g2.grid(row=7, column=2, padx=10,pady=5)
        g3 = Entry(self.frame1b, width=7,textvariable=self.g3)
        g3.grid(row=7, column=3, padx=10,pady=5)
        g4 = Entry(self.frame1b, width=8,textvariable=self.g4)
        g4.grid(row=7, column=4, padx=10,pady=5)

        h1 = Entry(self.frame1b, width=8,textvariable=self.h1)
        h1.grid(row=8, column=1, padx=10,pady=5)
        h2 = Entry(self.frame1b, width=9,textvariable=self.h2)
        h2.grid(row=8, column=2, padx=10,pady=5)
        h3 = Entry(self.frame1b, width=7,textvariable=self.h3)
        h3.grid(row=8, column=3, padx=10,pady=5)
        h4 = Entry(self.frame1b, width=8,textvariable=self.h4)
        h4.grid(row=8, column=4, padx=10,pady=5)

        i1 = Entry(self.frame1b, width=8,textvariable=self.i1)
        i1.grid(row=9, column=1, padx=10,pady=5)
        i2 = Entry(self.frame1b, width=9,textvariable=self.i2)
        i2.grid(row=9, column=2, padx=10,pady=5)
        i3 = Entry(self.frame1b, width=7,textvariable=self.i3)
        i3.grid(row=9, column=3, padx=10,pady=5)
        i4 = Entry(self.frame1b, width=8,textvariable=self.i4)
        i4.grid(row=9, column=4, padx=10,pady=5)

        self.frame5=Frame(self.root)
        self.frame5.place(x=400,y=300)
        subm=Button(self.frame5,text="Submit",bg="green",fg="white",command=self.wel)
        subm.grid(row=1,column=1)
        clr=Button(self.frame5,text="Clear",bg="red",fg="white",command=self.clear)
        clr.grid(row=1,column=2)

    def wel(self):
        wb = load_workbook(r'excel/Untitled spreadsheet.xlsx')
        sheet = wb.active

        sheet['A9'] = str(self.a1.get())
        sheet['B9'] = self.a2.get()
        sheet['C9'] = self.a3.get()
        sheet['D9'] = self.a4.get()


        sheet['A10'] = str(self.b1.get())
        sheet['B10'] = self.b2.get()
        sheet['C10'] = self.b3.get()



        sheet['A11'] = str(self.c1.get())
        sheet['B11'] = self.c2.get()
        sheet['C11'] = self.c3.get()
        sheet['D11'] = self.c4.get()

        sheet['A12'] = str(self.d1.get())
        sheet['B12'] = self.d2.get()
        sheet['C12'] = self.d3.get()
        sheet['D12'] = self.d4.get()

        sheet['A13'] = str(self.e1.get())
        sheet['B13'] = self.e2.get()
        sheet['C13'] = self.e3.get()
        sheet['D13'] = self.e4.get()

        sheet['A14'] = str(self.f1.get())
        sheet['B14'] = self.f2.get()
        sheet['C14'] = self.f3.get()
        sheet['D14'] = self.f4.get()

        sheet['A15'] = str(self.g1.get())
        sheet['B15'] = self.g2.get()
        sheet['C15'] = self.g3.get()
        sheet['D10'] = self.g4.get()

        sheet['A16'] = str(self.h1.get())
        sheet['B16'] = self.h2.get()
        sheet['C16'] = self.h3.get()
        sheet['D16'] = self.h4.get()

        sheet['A17'] = str(self.i1.get())
        sheet['B17'] = self.i2.get()
        sheet['C17'] = self.i3.get()
        sheet['D17'] = self.i4.get()
        sheet['H4']=str(self.ca1.get())
        sheet['H5']=str(self.ca2.get())
        sheet['H6']=str(self.ca3.get())
        sheet['I9'] = self.eekn.get()
        sheet['I10'] = str(self.eekd.get())
        sheet['I11']=str(self.eekdi.get())
        sheet['I12']=self.eekp.get()
        sheet['I13']=self.eekv.get()
        self.q="excel/"+str(str(self.eekd.get()[0])+str(self.eekd.get()[1])+str(self.eekd.get()[3])+str(self.eekd.get()[4])+str(self.eekd.get()[6])+str(self.eekd.get()[7])+str(self.cna.get())+".xlsx")
        sheet['D10']=self.b4.get()
        wb.save(self.q)






    def clear(self):

        self.cna.set("")
        self.ca1.set("")
        self.ca2.set("")
        self.ca3.set("")
        self.eekn.set(0)
        self.eekd.set("")
        self.eekdi.set("")
        self.eekp.set(0)
        self.eekv.set(0)
        self.da.set(0)

        self.a1.set("")
        self.a2.set(0)
        self.a3.set(0)
        self.a4.set(0)

        self.b1.set("")
        self.b2.set(0)
        self.b3.set(0)
        self.b4.set(0)

        self.c1.set("")
        self.c2.set(0)
        self.c3.set(0)
        self.c4.set(0)

        self.d1.set("")
        self.d2.set(0)
        self.d3.set(0)
        self.d4.set(0)

        self.e1.set("")
        self.e2.set(0)
        self.e3.set(0)
        self.e4.set(0)

        self.f1.set("")
        self.f2.set(0)
        self.f3.set(0)
        self.f4.set(0)

        self.g1.set("")
        self.g2.set(0)
        self.g3.set(0)
        self.g4.set(0)

        self.h1.set("")
        self.h2.set(0)
        self.h3.set(0)
        self.h4.set(0)

        self.i1.set("")
        self.i2.set(0)
        self.i3.set(0)
        self.i4.set(0)



root=Tk()
ob1=App(root)
root.mainloop()

"""def wel(self):
        self.f12=open("bills/"+str( str(self.eekd.get()[0])+str(self.eekd.get()[1])+str(self.eekd.get()[3])+str(self.eekd.get()[4])+str(self.eekd.get()[6])+str(self.eekd.get()[7]))+str(self.cna.get())+".txt" ,"w")
        self.data=(f"\t\t\t Chandra textile\t\t\t\n \
        \n-------------------------------------------------------------\
        \n\tFROM \
        \n\t\tBaskaran twisting\
        \n\t\t65 Sourashtra Colony\
        \n\t\tNagal Nagar, Dindigul\
        \n\t\t624003 \
        \n ------------------------------------------------ \
        \n\tTO \
        \n\t\t{self.cna.get()}\
        \n\t\t{self.ca1.get()}\
        \n\t\t{self.ca2.get()}\
        \n\t\t{self.ca3.get()}\
        \n ------------------------------------------------- \ ")
        self.f12.write(self.data)
        self.det()
        self.dates()
        self.total1()

        self.f12.close()
    def det(self):
        self.data1=(f"\n\t\teeki calculation kijam 73 \
              \n______________________________________________________________________________________ \
              \n\tlaat    number\t\t  {self.eekn.get()} \
              \n\tpetivaravu date\t\t {self.eekd.get()} \
              \n\tdeeniar        \t\t {self.eekdi.get()} \
              \n\tpetti          \t\t {self.eekp.get()} \
              \n\tvaravu kilo    \t\t {self.eekv.get()} \
              \n---------------------------------------------------------------------------------------\
            ")
        self.f12.write(self.data1)
    def dates(self):
        self.dates1=(f"\n _________________________________________________________________________________\
                     \n Date\t|\tkejam\t|\teeki\t|\tcount\t|\teeki\t|\ttotal |\
                     \n ___________________________________________________________________________________ ")
        self.f12.write(self.dates1)
        self.tot4=0
        if self.a1.get()!="":
            self.a5 = int((int(self.a2.get() * int(self.a3.get()) / 73)))
            self.tot4=int(self.tot4+(self.a5*self.a4.get()))
            self.f12.write(f"\n {self.a1.get()}\t|\t{self.a2.get()}\t|\t {self.a3.get()}\t|\t{self.a4.get()}\t|\t{self.a5}\t|\t{self.tot4}|\
                           \n _______________________________________________________________________________")
        if self.b1.get()!="":
            self.b5 = int((int(self.b2.get() * int(self.b3.get()) / 73)))
            self.tot4 =int(self.tot4 + (self.b5*self.b4.get()))
            self.f12.write(f"\n {self.b1.get()}\t|\t{self.b2.get()}\t|\t {self.b3.get()}\t|\t{self.b4.get()}\t|\t{self.b5}\t|\t{self.tot4}|\
                           \n ________________________________________________________________________________")
        if self.c1.get()!="":
            self.c5 = int((int(self.c2.get() * int(self.c3.get()) / 73)))
            self.tot4 = int(self.tot4 + (self.c5*self.c4.get()))
            self.f12.write(f"\n {self.c1.get()}\t|\t{self.c2.get()}\t|\t {self.c3.get()}\t|\t{self.c4.get()}\t|\t{self.c5}\t|\t{self.tot4}|\
                           \n _________________________________________________________________________________")
        if self.d1.get()!="":
            self.d5 = int((int(self.d2.get() * int(self.d3.get()) / 73)))
            self.tot4 = int(self.tot4 + (self.d5*self.d4.get()))
            self.f12.write(f"\n {self.d1.get()}\t|\t{self.d2.get()}\t|\t {self.d3.get()}\t|\t{self.d4.get()}\t|\t{self.d5}\t|\t{self.tot4}|\
                           \n _________________________________________________________________________________")
        if self.e1.get()!="":
            self.e5 = int((int(self.e2.get() * int(self.e3.get()) / 73)))
            self.tot4 = int(self.tot4 + (self.e5*self.e4.get()))
            self.f12.write(f"\n {self.e1.get()}\t|\t{self.e2.get()}\t|\t {self.e3.get()}\t|\t{self.e4.get()}\t|\t{self.e5}\t|\t{self.tot4}|\
                           \n __________________________________________________________________________________")
        if self.f1.get()!="":
            self.f5 = int((int(self.f2.get() * int(self.f3.get()) / 73)))
            self.tot4 = int(self.tot4 + (self.f5*self.f4.get()))
            self.f12.write(f"\n {self.f1.get()}\t|\t{self.f2.get()}\t|\t {self.f3.get()}\t|\t{self.f4.get()}\t|\t{self.c5}\t|\t{self.tot4}|\
                           \n __________________________________________________________________________________")
        if self.g1.get()!="":
            self.g5 = int((int(self.g2.get() * int(self.g3.get()) / 73)))
            self.tot4 = int(self.tot4 + (self.g5*self.f4.get()))
            self.f12.write(f"\n {self.c1.get()}\t|\t{self.c2.get()}\t|\t {self.c3.get()}\t|\t{self.c4.get()}\t|\t{self.c5}\t|\t{self.tot4}|\
                           \n _________________________________________________________________________________")
        if self.h1.get()!="":
            self.h5 = int((int(self.h2.get() * int(self.h3.get()) / 73)))
            self.tot4 = int(self.tot4 + (self.h5*self.h4.get()))
            self.f12.write(f"\n {self.h1.get()}\t|\t{self.h2.get()}\t|\t {self.h3.get()}\t|\t{self.h4.get()}\t|\t{self.c5}\t|\t{self.tot4}|\
                           \n __________________________________________________________________________________")
        if self.i1.get()!="":
            self.i5 = int((int(self.i2.get() * int(self.i3.get()) / 73)))
            self.tot4 = int(self.tot4 + (self.i5*self.i4.get()))
            self.f12.write(f"\n {self.i1.get()}\t|\t{self.i2.get()}\t|\t {self.i3.get()}\t|\t{self.i4.get()}\t|\t{self.i5}\t|\t{self.tot4}|\
                           \n ___________________________________________________________________________________")



    def total1(self):
        self.total2 = self.a4.get() + self.b4.get() + self.c4.get() + self.d4.get() + self.e4.get() + self.f4.get() + self.g4.get() + self.h4.get() + self.i4.get()
        self.f12.write(f"\nTotal count {self.total2}")
"""