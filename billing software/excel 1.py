from openpyxl import load_workbook


wb=load_workbook(r'excel/Untitled spreadsheet.xlsx')
sheet=wb.active
into="12/12/22"
sheet['A9']=into
sheet['B9']=73
sheet['C9']=210
sheet['D9']=20

sheet['A10']="13/12/22"
sheet['B10']=69
sheet['C10']=200
sheet['D10']=15


sheet['I9']=1
sheet['I10']=into
wb.save(str(str(self.eekd.get()[0])+str(self.eekd.get()[1])+str(self.eekd.get()[3])+str(self.eekd.get()[4])+str(self.eekd.get()[6])+str(self.eekd.get()[7]))+str(self.cna.get())+".txt")