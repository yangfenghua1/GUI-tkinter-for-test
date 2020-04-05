import tkinter
import xlrd
import os


root = tkinter.Tk()
root.title('NaCl-Calculation')

frame = tkinter.Frame(root)

frame.pack(padx=130, pady=80)

v1 = tkinter.StringVar()
v2 = tkinter.StringVar()
v3 = tkinter.StringVar()
v4 = tkinter.StringVar()
v5 = tkinter.StringVar()
v6 = tkinter.StringVar()
v7 = tkinter.StringVar()
v8 = tkinter.StringVar()
v9 = tkinter.StringVar()
v10 = tkinter.StringVar()
def test(content):

    return content.isdigit()


testCMD = frame.register(test)
tkinter.Label(frame, text='COD-setting:').grid(row=0, column=1)
e1 = tkinter.Entry(frame, width=60, textvariable=v1, validate='key', \
           validatecommand=(testCMD, '%P')).grid(row=0, column=2)
tkinter.Label(frame, text='COD-result:').grid(row=0, column=4)
e2 = tkinter.Entry(frame, width=60, textvariable=v6, state='readonly').grid(row=0, column=5)

 
tkinter.Label(frame, text='Si-setting:').grid(row=1, column=1)
e3 = tkinter.Entry(frame, width=60, textvariable=v2, validate='key', \
           validatecommand=(testCMD, '%P')).grid(row=1, column=2)
tkinter.Label(frame, text='Si-result:').grid(row=1, column=4)
e4 = tkinter.Entry(frame, width=60, textvariable=v7, state='readonly').grid(row=1, column=5)

tkinter.Label(frame, text='Co3-setting:').grid(row=2, column=1)
e5 = tkinter.Entry(frame, width=60, textvariable=v3, validate='key', \
           validatecommand=(testCMD, '%P')).grid(row=2, column=2)
tkinter.Label(frame, text='Co3-result:').grid(row=2, column=4)
e6 = tkinter.Entry(frame, width=60, textvariable=v8, state='readonly').grid(row=2, column=5)

 
tkinter.Label(frame, text='Na2So4-setting:').grid(row=3, column=1)
e7 = tkinter.Entry(frame, width=60, textvariable=v4, validate='key', \
           validatecommand=(testCMD, '%P')).grid(row=3, column=2)
tkinter.Label(frame, text='Na2So4-result:').grid(row=3, column=4)
e8 = tkinter.Entry(frame, width=60, textvariable=v9, state='readonly').grid(row=3, column=5)

tkinter.Label(frame, text='MinValue:').grid(row=9, column=4)
e9 = tkinter.Entry(frame, width=60, textvariable=v10, state='readonly').grid(row=9, column=5)


def calc():
	workbook = xlrd.open_workbook('NaCl.xlsx')
	sheet0 = workbook.sheet_by_index(0)
	sheet1 = workbook.sheet_by_index(1)
	Cod_mgl   = float(sheet0.cell_value(14,1))
	Si_mgl    = float(sheet0.cell_value(13,1))
	Co3_mgl   = float(sheet0.cell_value(10,1))
	So4_mgl   = float(sheet0.cell_value(5,1))
	Density   = float(sheet0.cell_value(2,1))
	Const2    = float(sheet1.cell_value(2,1))
	result_COD = float(v1.get()) / Cod_mgl * Const2
	result_SI = float(v2.get()) / Si_mgl * Const2
	result_CO3 = float(v3.get()) / Co3_mgl * Const2
	result_Na2So4 = float(v4.get()) / (So4_mgl / 96 * 142 / 10000)
	v6.set(str(result_COD))
	v7.set(str(result_SI))
	v8.set(str(result_CO3))
	v9.set(str(result_Na2So4))
	v10.set(str(min(result_COD,result_SI,result_CO3,result_Na2So4)))
    
tkinter.Button(frame, text='result', command=calc, width=20,bg = 'green').grid(row=8, column=4, pady=5)



tkinter.mainloop()
