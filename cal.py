from tkinter import *
from tkinter import ttk
import xlrd
import os

#global varieries
Entry_width = 10


root = Tk()
root.title('NaCl-Calculation')

frame = Frame(root)
frame.pack(padx=130, pady=80)

tree_result = ttk.Treeview(root)
tree_result['columns'] = ['jinshui','lvhuanamuye','lengningye','lvhuanajiejingyan','sunshishuiliang','','GB/T 5462-2015','gongyeshiyanyiji']

#tree_result.pack(side=LEFT)
tree_result.pack()


tree_result.column('jinshui',width=100, anchor="center")
tree_result.column('lvhuanamuye',width=100, anchor="center")
tree_result.column('lengningye',width=100, anchor="center")
tree_result.column('lvhuanajiejingyan',width=100, anchor="center")
tree_result.column('sunshishuiliang',width=100, anchor="center")
tree_result.column('',width=100, anchor="center")
tree_result.column('GB/T 5462-2015',width=150, anchor="center")
tree_result.column('gongyeshiyanyiji',width=100, anchor="center")

tree_result.heading('jinshui',text='进水')
tree_result.heading('lvhuanamuye',text='氯化钠母液')
tree_result.heading('lengningye',text='冷凝液')
tree_result.heading('lvhuanajiejingyan',text='氯化钠结晶盐')
tree_result.heading('sunshishuiliang',text='损失水量')
tree_result.heading('GB/T 5462-2015',text='GB/T 5462-2015')
tree_result.heading('gongyeshiyanyiji',text='工业湿盐一级')



def test(content):

    return content.isdigit()


testCMD = frame.register(test)



# paras of excel_1
excel_1_x = 2
excel_1_y = 2


Label(frame, text='进料').grid(row=excel_1_x-1, column=excel_1_y+1)
Label(frame, text='结晶母液').grid(row=excel_1_x-1, column=excel_1_y+2)

Label(frame, text='TDS').grid(row=excel_1_x-1, column=excel_1_y+3)
Label(frame, text='NaCl溶液密度').grid(row=excel_1_x-1, column=excel_1_y+4)
Label(frame, text='NaSO4溶液密度').grid(row=excel_1_x-1, column=excel_1_y+5)

v00 = StringVar()
v01 = StringVar()
v02 = StringVar()
v03 = StringVar()
v04 = StringVar()
v05 = StringVar()
v06 = StringVar()
v07 = StringVar()
v08 = StringVar()
v09 = StringVar()
v0A = StringVar()
v0B = StringVar()
v0C = StringVar()
v0D = StringVar()
v0E = StringVar()
v0F = StringVar()

v10 = StringVar()
v11 = StringVar()
v12 = StringVar()
v13 = StringVar()
v14 = StringVar()
v15 = StringVar()
v16 = StringVar()
v17 = StringVar()
v18 = StringVar()
v19 = StringVar()
v1A = StringVar()
v1B = StringVar()
v1C = StringVar()
v1D = StringVar()
v1E = StringVar()
v1F = StringVar()

v20 = StringVar()
v21 = StringVar()
v22 = StringVar()
v23 = StringVar()
v24 = StringVar()
v25 = StringVar()
v26 = StringVar()
v27 = StringVar()
v28 = StringVar()
v29 = StringVar()
v2A = StringVar()
v2B = StringVar()
v2C = StringVar()
v2D = StringVar()
v2E = StringVar()
v2F = StringVar()

v30 = StringVar()
v31 = StringVar()
v32 = StringVar()
v33 = StringVar()
v34 = StringVar()
v35 = StringVar()
v36 = StringVar()
v37 = StringVar()
v38 = StringVar()
v39 = StringVar()
v3A = StringVar()
v3B = StringVar()
v3C = StringVar()
v3D = StringVar()
v3E = StringVar()
v3F = StringVar()

v40 = StringVar()
v41 = StringVar()
v42 = StringVar()
v43 = StringVar()
v44 = StringVar()
v45 = StringVar()
v46 = StringVar()
v47 = StringVar()
v48 = StringVar()
v49 = StringVar()
v4A = StringVar()
v4B = StringVar()
v4C = StringVar()
v4D = StringVar()
v4E = StringVar()
v4F = StringVar()

v50 = StringVar()
v51 = StringVar()
v52 = StringVar()
v53 = StringVar()
v54 = StringVar()
v55 = StringVar()
v56 = StringVar()
v57 = StringVar()
v58 = StringVar()
v59 = StringVar()
v5A = StringVar()
v5B = StringVar()
v5C = StringVar()
v5D = StringVar()
v5E = StringVar()
v5F = StringVar()

v60 = StringVar()
v61 = StringVar()
v62 = StringVar()
v63 = StringVar()
v64 = StringVar()
v65 = StringVar()
v66 = StringVar()
v67 = StringVar()
v68 = StringVar()
v69 = StringVar()
v6A = StringVar()
v6B = StringVar()
v6C = StringVar()
v6D = StringVar()
v6E = StringVar()
v6F = StringVar()

a = {1:v00,2:v01,3:v02,4:v03,5:v04,6:v05,7:v06,8:v07,9:v08,10:v09,11:v0A,12:v0B,13:v0C,14:v0D,15:v0E,16:v0F,17:v10,18:v11,19:v12,20:v13,21:v14,22:v15,
23:v16,24:v17,25:v18,26:v19,27:v1A,28:v1B,29:v1C,30:v1D,31:v1E,32:v1F,33:v20,34:v21,35:v22,36:v23,37:v24,38:v25,39:v26,40:v27,41:v28,42:v29,43:v2A,44:v2B,
45:v2C,46:v2D,47:v2E,48:v2F,49:v30,50:v31,51:v32,52:v33,53:v34,54:v35,55:v36,56:v37,57:v38,58:v39,59:v3A,60:v3B,61:v3C,62:v3D,63:v3E,64:v3F,65:v40,66:v41,
67:v42,68:v43,69:v44,70:v45,71:v46,72:v47,73:v48,74:v49,75:v4A,76:v4B,77:v4C,78:v4D,79:v4E,80:v4F,81:v50,82:v51,83:v52,84:v53,85:v54,86:v55,87:v56,88:v57,
89:v58,90:v59,91:v5A,92:v5B,93:v5C,94:v5D,95:v5E,96:v5F,97:v60,98:v61,99:v62,100:v63,101:v64,102:v65,103:v66,104:v67,105:v68,106:v69,107:v6A,108:v6B,109:v6C,
110:v6D,111:v6E,112:v6F}


Label(frame, text='流量m3/h').grid(row=excel_1_x+0, column=excel_1_y+0)
e01 = Entry(frame, width=Entry_width, textvariable=v00, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 0, column=excel_1_y+1)
e02 = Entry(frame, width=Entry_width, textvariable=v01, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+0, column=excel_1_y+2)
e03 = Entry(frame, width=Entry_width, textvariable=v02, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 0, column=excel_1_y+3)
e04 = Entry(frame, width=Entry_width, textvariable=v03, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+0, column=excel_1_y+4)
e05 = Entry(frame, width=Entry_width, textvariable=v04, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+0, column=excel_1_y+5)

Label(frame, text='密度kg/m3').grid(row=excel_1_x+1, column=excel_1_y+0)
e11 = Entry(frame, width=Entry_width, textvariable=v05, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+1, column=excel_1_y+1)
e12 = Entry(frame, width=Entry_width, textvariable=v06, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+1, column=excel_1_y+2)
e13 = Entry(frame, width=Entry_width, textvariable=v07, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 1, column=excel_1_y+3)
e14 = Entry(frame, width=Entry_width, textvariable=v08, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+1, column=excel_1_y+4)
e15 = Entry(frame, width=Entry_width, textvariable=v09, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+1, column=excel_1_y+5)

Label(frame, text='质量kg/h').grid(row=excel_1_x+2, column=excel_1_y+0)
e21 = Entry(frame, width=Entry_width, textvariable=v0A, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+2, column=excel_1_y+1)
e22 = Entry(frame, width=Entry_width, textvariable=v0B, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+2, column=excel_1_y+2)
e23 = Entry(frame, width=Entry_width, textvariable=v0C, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 2, column=excel_1_y+3)
e24 = Entry(frame, width=Entry_width, textvariable=v0D, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+2, column=excel_1_y+4)
e25 = Entry(frame, width=Entry_width, textvariable=v0E, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+2, column=excel_1_y+5)

Label(frame, text='TDS  mg/L').grid(row=excel_1_x+3, column=excel_1_y+0)
e31 = Entry(frame, width=Entry_width, textvariable=v0F, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+3, column=excel_1_y+1)
e32 = Entry(frame, width=Entry_width, textvariable=v10, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+3, column=excel_1_y+2)
e33 = Entry(frame, width=Entry_width, textvariable=v11, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 3, column=excel_1_y+3)
e34 = Entry(frame, width=Entry_width, textvariable=v12, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+3, column=excel_1_y+4)
e35 = Entry(frame, width=Entry_width, textvariable=v13, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+3, column=excel_1_y+5)

Label(frame, text='SO42- mg/L').grid(row=excel_1_x+4, column=excel_1_y+0)
e41 = Entry(frame, width=Entry_width, textvariable=v14, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+4, column=excel_1_y+1)
e42 = Entry(frame, width=Entry_width, textvariable=v15, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+4, column=excel_1_y+2)
e43 = Entry(frame, width=Entry_width, textvariable=v16, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 4, column=excel_1_y+3)
e44 = Entry(frame, width=Entry_width, textvariable=v17, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+4, column=excel_1_y+4)
e45 = Entry(frame, width=Entry_width, textvariable=v18, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+4, column=excel_1_y+5)

Label(frame, text='Cl- mg/L').grid(row=excel_1_x+5, column=excel_1_y+0)
e51 = Entry(frame, width=Entry_width, textvariable=v19, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+5, column=excel_1_y+1)
e52 = Entry(frame, width=Entry_width, textvariable=v1A, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+5, column=excel_1_y+2)
e53 = Entry(frame, width=Entry_width, textvariable=v1B, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 5, column=excel_1_y+3)
e54 = Entry(frame, width=Entry_width, textvariable=v1C, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+5, column=excel_1_y+4)
e55 = Entry(frame, width=Entry_width, textvariable=v1D, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+5, column=excel_1_y+5)

Label(frame, text='Ca2+ mg/L').grid(row=excel_1_x+6, column=excel_1_y+0)
e61 = Entry(frame, width=Entry_width, textvariable=v1E, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+6, column=excel_1_y+1)
e62 = Entry(frame, width=Entry_width, textvariable=v1F, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+6, column=excel_1_y+2)
e63 = Entry(frame, width=Entry_width, textvariable=v20, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 6, column=excel_1_y+3)
e64 = Entry(frame, width=Entry_width, textvariable=v21, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+6, column=excel_1_y+4)
e65 = Entry(frame, width=Entry_width, textvariable=v22, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+6, column=excel_1_y+5)

Label(frame, text='Mg2+ mg/L').grid(row=excel_1_x+7, column=excel_1_y+0)
e71 = Entry(frame, width=Entry_width, textvariable=v23, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+7, column=excel_1_y+1)
e72 = Entry(frame, width=Entry_width, textvariable=v24, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+7, column=excel_1_y+2)
e73 = Entry(frame, width=Entry_width, textvariable=v25, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 7, column=excel_1_y+3)
e74 = Entry(frame, width=Entry_width, textvariable=v26, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+7, column=excel_1_y+4)
e75 = Entry(frame, width=Entry_width, textvariable=v27, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+7, column=excel_1_y+5)

Label(frame, text='Na+ mg/L').grid(row=excel_1_x+8, column=excel_1_y+0)
e81 = Entry(frame, width=Entry_width, textvariable=v28, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+8, column=excel_1_y+1)
e82 = Entry(frame, width=Entry_width, textvariable=v29, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+8, column=excel_1_y+2)
e83 = Entry(frame, width=Entry_width, textvariable=v2A, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 8, column=excel_1_y+3)
e84 = Entry(frame, width=Entry_width, textvariable=v2B, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+8, column=excel_1_y+4)
e85 = Entry(frame, width=Entry_width, textvariable=v2C, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+8, column=excel_1_y+5)

Label(frame, text='HCO3- mg/L').grid(row=excel_1_x+9, column=excel_1_y+0)
e91 = Entry(frame, width=Entry_width, textvariable=v2D, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+9, column=excel_1_y+1)
e92 = Entry(frame, width=Entry_width, textvariable=v2E, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+9, column=excel_1_y+2)
e93 = Entry(frame, width=Entry_width, textvariable=v2F, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 9, column=excel_1_y+3)
e94 = Entry(frame, width=Entry_width, textvariable=v30, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+9, column=excel_1_y+4)
e95 = Entry(frame, width=Entry_width, textvariable=v31, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+9, column=excel_1_y+5)

Label(frame, text='NO3- mg/L').grid(row=excel_1_x+10, column=excel_1_y+0)
eA1 = Entry(frame, width=Entry_width, textvariable=v32, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+10, column=excel_1_y+1)
eA2 = Entry(frame, width=Entry_width, textvariable=v33, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+10, column=excel_1_y+2)
eA3 = Entry(frame, width=Entry_width, textvariable=v34, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x + 10, column=excel_1_y+3)
eA4 = Entry(frame, width=Entry_width, textvariable=v35, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+10, column=excel_1_y+4)
eA5 = Entry(frame, width=Entry_width, textvariable=v36, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+10, column=excel_1_y+5)

Label(frame, text='氨氮 mg/L').grid(row=excel_1_x+11, column=excel_1_y+0)
eB1 = Entry(frame, width=Entry_width, textvariable=v37, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+11, column=excel_1_y+1)
eB2 = Entry(frame, width=Entry_width, textvariable=v38, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+11, column=excel_1_y+2)

Label(frame, text='总硅 mg/L').grid(row=excel_1_x+12, column=excel_1_y+0)
eC1 = Entry(frame, width=Entry_width, textvariable=v39, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+12, column=excel_1_y+1)
eC2 = Entry(frame, width=Entry_width, textvariable=v3A, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+12, column=excel_1_y+2)

Label(frame, text='COD mg/L').grid(row=excel_1_x+13, column=excel_1_y+0)
eD1 = Entry(frame, width=Entry_width, textvariable=v3B, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_1_x+13, column=excel_1_y+1)
eD2 = Entry(frame, width=Entry_width, textvariable=v3C, validatecommand=(testCMD, '%P')).grid(row=excel_1_x+13, column=excel_1_y+2)


#paras of excel2
excel_2_x = 2
excel_2_y = 10


Label(frame, text='硫酸钠等级').grid(row=excel_2_x+0, column=excel_2_y+0)
e_01 = Entry(frame, width=Entry_width, textvariable=v3D, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x + 0, column=excel_2_y+1)

Label(frame, text='工业干盐—1/2').grid(row=excel_2_x+1, column=excel_2_y+0)
e_11 = Entry(frame, width=Entry_width, textvariable=v3E, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+1, column=excel_2_y+1)

Label(frame, text='品质').grid(row=excel_2_x+2, column=excel_2_y+0)
e_21 = Entry(frame, width=Entry_width, textvariable=v3F, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+2, column=excel_2_y+1)

Label(frame, text='蒸发结晶经验值').grid(row=excel_2_x+3, column=excel_2_y+0)
e_31 = Entry(frame, width=Entry_width, textvariable=v40, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+3, column=excel_2_y+1)

Label(frame, text='氯化钠纯度').grid(row=excel_2_x+4, column=excel_2_y+0)
e_41 = Entry(frame, width=Entry_width, textvariable=v41, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+4, column=excel_2_y+1)

Label(frame, text='硫酸盐含量').grid(row=excel_2_x+5, column=excel_2_y+0)
e_51 = Entry(frame, width=Entry_width, textvariable=v42, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+5, column=excel_2_y+1)

Label(frame, text='结晶盐经验值').grid(row=excel_2_x+6, column=excel_2_y+0)
e_61 = Entry(frame, width=Entry_width, textvariable=v43, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+6, column=excel_2_y+1)

Label(frame, text='钠离子经验值').grid(row=excel_2_x+7, column=excel_2_y+0)
e_71 = Entry(frame, width=Entry_width, textvariable=v44, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+7, column=excel_2_y+1)

Label(frame, text='冷凝水携带离子率').grid(row=excel_2_x+8, column=excel_2_y+0)
e_81 = Entry(frame, width=Entry_width, textvariable=v45, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+8, column=excel_2_y+1)

Label(frame, text='NaCl').grid(row=excel_2_x+9, column=excel_2_y+0)
e_91 = Entry(frame, width=Entry_width, textvariable=v46, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+9, column=excel_2_y+1)

Label(frame, text='Na2SO4').grid(row=excel_2_x+10, column=excel_2_y+0)
e_A1 = Entry(frame, width=Entry_width, textvariable=v47, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_2_x+10, column=excel_2_y+1)



excel_3_x = 2
excel_3_y = 13

Label(frame, text='COD设定值').grid(row=excel_3_x+0, column=excel_3_y+0)
e_02 = Entry(frame, width=Entry_width, textvariable=v48, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_3_x+0, column=excel_3_y+1)

Label(frame, text='COD最大值').grid(row=excel_3_x+1, column=excel_3_y+0)
e_12 = Entry(frame, width=Entry_width, textvariable=v49, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_3_x+1, column=excel_3_y+1)

Label(frame, text='COD').grid(row=excel_3_x+2, column=excel_3_y+0)
e_13 = Entry(frame, width=Entry_width, textvariable=v4A, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_3_x+2, column=excel_3_y+1)

Label(frame, text='结晶浓缩倍数').grid(row=excel_3_x+3, column=excel_3_y+0)
e_14 = Entry(frame, width=Entry_width, textvariable=v4B, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_3_x+3, column=excel_3_y+1)

excel_4_x = 2
excel_4_y = 15


Label(frame, text='总硅设定值').grid(row=excel_4_x+0, column=excel_4_y+0)
e_03 = Entry(frame, width=Entry_width, textvariable=v4C, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_4_x+0, column=excel_4_y+1)

Label(frame, text='总硅最大值').grid(row=excel_4_x+1, column=excel_4_y+0)
e_13 = Entry(frame, width=Entry_width, textvariable=v4D, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_4_x+1, column=excel_4_y+1)

Label(frame, text='总硅').grid(row=excel_4_x+2, column=excel_4_y+0)
e_23 = Entry(frame, width=Entry_width, textvariable=v4E, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_4_x+2, column=excel_4_y+1)

Label(frame, text='结晶浓缩倍数').grid(row=excel_4_x+3, column=excel_4_y+0)
e_33 = Entry(frame, width=Entry_width, textvariable=v4F, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_4_x+3, column=excel_4_y+1)

excel_5_x = 2
excel_5_y = 18


Label(frame, text='碳酸根设定值').grid(row=excel_5_x+0, column=excel_5_y+0)
e_04 = Entry(frame, width=Entry_width, textvariable=v50, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_5_x+0, column=excel_5_y+1)

Label(frame, text='碳酸根最大值').grid(row=excel_5_x+1, column=excel_5_y+0)
e_14 = Entry(frame, width=Entry_width, textvariable=v51, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_5_x+1, column=excel_5_y+1)

Label(frame, text='碳酸根').grid(row=excel_5_x+2, column=excel_5_y+0)
e_24 = Entry(frame, width=Entry_width, textvariable=v52, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_5_x+2, column=excel_5_y+1)

Label(frame, text='结晶浓缩倍数').grid(row=excel_5_x+3, column=excel_5_y+0)
e_34 = Entry(frame, width=Entry_width, textvariable=v53, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_5_x+3, column=excel_5_y+1)

excel_6_x = 2
excel_6_y = 21


Label(frame, text='Na2SO4设定值').grid(row=excel_6_x+0, column=excel_6_y+0)
e_05 = Entry(frame, width=Entry_width, textvariable=v54, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_6_x+0, column=excel_6_y+1)

Label(frame, text='Na2SO4最大值').grid(row=excel_6_x+1, column=excel_6_y+0)
e_05= Entry(frame, width=Entry_width, textvariable=v55, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_6_x+1, column=excel_6_y+1)

Label(frame, text='Na2SO4').grid(row=excel_6_x+2, column=excel_6_y+0)
e_25 = Entry(frame, width=Entry_width, textvariable=v56, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_6_x+2, column=excel_6_y+1)

Label(frame, text='结晶浓缩倍数').grid(row=excel_6_x+3, column=excel_6_y+0)
e_35 = Entry(frame, width=Entry_width, textvariable=v57, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_6_x+3, column=excel_6_y+1)

excel_7_x = 2
excel_7_y = 24


Label(frame, text='NaCl设定值').grid(row=excel_7_x+0, column=excel_7_y+0)
e_06 = Entry(frame, width=Entry_width, textvariable=v58, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_7_x+0, column=excel_7_y+1)

Label(frame, text='NaCl最小值').grid(row=excel_7_x+1, column=excel_7_y+0)
e_16 = Entry(frame, width=Entry_width, textvariable=v59, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_7_x+1, column=excel_7_y+1)

Label(frame, text='NaCl').grid(row=excel_7_x+2, column=excel_7_y+0)
e_26 = Entry(frame, width=Entry_width, textvariable=v5A, validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_7_x+2, column=excel_7_y+1)


excel_8_x = 10 + excel_2_x
excel_8_y = excel_7_y + 1

Label(frame, text='浓缩倍数最小值').grid(row=excel_8_x, column=excel_8_y-1)
e_66 = Entry(frame, width=Entry_width, textvariable=v5B, state='readonly',validate='key', validatecommand=(testCMD, '%P')).grid(row=excel_8_x, column=excel_8_y)

def ReadAndWrite():
	workbook = xlrd.open_workbook('NaCl.xlsx')
	sheet0 = workbook.sheet_by_index(0)
	sheet1 = workbook.sheet_by_index(1)
	i = 0
	for num in range(1,12):
		i = i + 1
		a[i].set(sheet0.cell_value(num,1))
		i = i + 1
		a[i].set(sheet0.cell_value(num,2))
		i = i + 1
		a[i].set(sheet0.cell_value(num,6))
		i = i + 1
		a[i].set(sheet0.cell_value(num,7))
		i = i + 1
		a[i].set(sheet0.cell_value(num,8))
	i = 55
	for num1 in range(12,15):
		i = i + 1
		a[i].set(sheet0.cell_value(num1,1))
		i = i + 1
		a[i].set(sheet0.cell_value(num1,2))
	#v3D 开始
	a[62].set(sheet1.cell_value(0,1))
	a[63].set(sheet1.cell_value(3,0))
	a[64].set(sheet1.cell_value(5,0))
	a[65].set(sheet1.cell_value(2,1))
	a[66].set(sheet1.cell_value(4,1))
	a[67].set(sheet1.cell_value(6,1))
	a[68].set(sheet1.cell_value(8,1))
	a[69].set(sheet1.cell_value(10,1))
	a[70].set(sheet1.cell_value(12,1))
	a[71].set(sheet1.cell_value(3,2))
	a[72].set(sheet1.cell_value(5,2))

	a[73].set(sheet1.cell_value(3,3))
	a[74].set(sheet1.cell_value(3,4))
	a[75].set(sheet1.cell_value(3,5))
	a[76].set(sheet1.cell_value(3,6))#COD 最小


	a[77].set(sheet1.cell_value(5,3))
	a[78].set(sheet1.cell_value(5,4))
	a[79].set(sheet1.cell_value(5,5))
	a[80].set(sheet1.cell_value(5,6))#总硅最小

	a[81].set(sheet1.cell_value(7,3))
	a[82].set(sheet1.cell_value(7,4))
	a[83].set(sheet1.cell_value(7,5))
	a[84].set(sheet1.cell_value(7,6))#碳酸根最小

	a[85].set(sheet1.cell_value(9,3))
	a[86].set(sheet1.cell_value(9,4))
	a[87].set(sheet1.cell_value(9,5))
	a[88].set(sheet1.cell_value(9,6))#硫酸钠最小

	a[89].set(sheet1.cell_value(11,3))
	a[90].set(sheet1.cell_value(11,4))
	a[91].set(sheet1.cell_value(11,5))



ReadAndWrite()

def calc():
	v5B.set(str(min(a[76].get(),a[80].get(),a[84].get(),a[88].get())))
	tree_result.insert('',0,text='流量m3/h',values=(float(v00.get()),float(v00.get())/float(v5B.get()),4.6,'','0.2','','氯化钠 w%',95))
	tree_result.insert('',1,text='密度kg/m3',values=(float(v05.get()),float(v06.get()),1000,'','','','水不溶物 w%',0.001))
	tree_result.insert('',2,text='质量kg/h',values=(float(v0A.get()),float(v06.get())*float(v00.get())/float(v5B.get()),'w1','3','33','','钙和镁 w%',0.05))
	tree_result.insert('',3,text='TDS  mg/L',values=(float(v0F.get()),21,'n1','4','44','','硫酸根离子 w%',0.007))
	tree_result.insert('',4,text='SO42- mg/L',values=(float(v14.get()),21,'n1','1','55','','水分 w%',0.035))
	tree_result.insert('',5,text='CL- mg/L',values=(float(v19.get()),21,'a1','2','66'))
	tree_result.insert('',6,text='Ca2+ mg/L',values=(float(v1E.get()),21,'w1','3','77'))
	tree_result.insert('',7,text='Mg2+ mg/L',values=(float(v23.get()),21,'n1','4','88'))
	tree_result.insert('',8,text='Na+ mg/L',values=(float(v28.get()),21,'n1','1','99'))
	tree_result.insert('',9,text='HCO3- mg/L',values=(float(v2C.get()),21,'a1','2','1010'))
	tree_result.insert('',10,text='NO3- mg/L',values=(float(v31.get()),21,'w1','3','1111'))
	tree_result.insert('',11,text='氨氮 mg/L',values=(float(v37.get()),21,'n1','4','1212'))
	tree_result.insert('',12,text='总硅 mg/L',values=(float(v39.get()),21,'w1','3','1313'))
	tree_result.insert('',13,text='COD mg/L',values=(float(v3B.get()),21,'n1','4','1414'))

Button(frame, text='result', command=calc, width=10,bg = 'green').grid(row=excel_8_x+5, column=excel_8_y, pady=5)

mainloop()
