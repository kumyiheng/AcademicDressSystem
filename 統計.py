import tkinter as tk
import shutil
import os
import numpy as np
import tkinter.font as tkFont
from tkcalendar import Calendar, DateEntry
from tkinter import ttk, filedialog, messagebox, simpledialog
from openpyxl import load_workbook, Workbook
from datetime import datetime, timedelta
from openpyxl.styles import Alignment, Border, Side

window = tk.Tk()
window.title('window')
window.geometry('1050x630')
window.configure(bg="#7AFEC6")
columns_1 = ("學號", "姓名", "系所", "歸還時間", "歸還狀況")
ID = tk.StringVar()
tree_1 = ttk.Treeview(window, height=1, show="headings", columns=columns_1)
columns_2 = ("日期", "學士", "碩士", "博士")
text = ["文學院", "理學院", "社科院", "工學院", "管理學院", "法學院", "教育學院"]
tree_2 = ttk.Treeview(window, height=8, padding = 3, show="headings", columns=columns_2)
ttk.Style().configure("Treeview.Heading", font=('Arial', 20), rowheight=45)
ttk.Style().configure("Treeview", font=('Arial', 20), rowheight=45)

total = [0, 0, 0, 0]
college = [[0 for _ in range(4)] for _ in range(7)]
date = tk.StringVar(window, Calendar.date.today().strftime("%m/%d/%y"))
combo = ttk.Combobox(window, values=["上午", "下午", "全天"], state="readonly", width=10)
msfont = tkFont.Font(family="Arial", size=50)

def reset_st():
    for i in range(3):
        total[i] = 0
        for j in range(7):
            college[j][i] = 0

def create_path(pathname):
    #建立資料夾
    if not os.path.isdir(pathname):
        os.mkdir(pathname)

def search():
    reset_st()
    wb = load_workbook("所有系所名單.xlsx")
    sheet = wb['Sheet']
    day = datetime.strptime(date.get()+" 00:00:00", '%m/%d/%y %H:%M:%S')
    time_1 = timedelta(0, 0, 0, 0, 0, 0, 0)
    time_2 = timedelta(0, 0, 0, 0, 0, 13, 0)
    time_3 = timedelta(1, 0, 0, 0, 0, 0, 0)
    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
        code = str(row[2]).lstrip('=RIGHT("')
        if row[6] != None:
            time = row[6] - day
            if (time > time_1 and time < time_2 and combo.current() == 0) or (time > time_2 and time < time_3 and combo.current() == 1) or (time > time_1 and time < time_3 and combo.current() == 2):
                for i in range (1,8):
                    if int(code[0]) == i:
                        if isinstance(row[0], int):
                            x = list(row)
                            x[0] = str(row[0])
                            row = tuple(x)
                        if str(row[0][0]) == '4':
                            college[i-1][0]+=1
                            total[0]+=1
                        elif str(row[0][0]) == '6' or str(row[0][0]) == '5':
                            college[i-1][1]+=1
                            total[1]+=1
                        elif str(row[0][0]) == '8':
                            college[i-1][2]+=1
                            total[2]+=1
    delButton(tree_2)
    order = [0, 2, 4, 3, 1, 5, 6]
    x = 0
    for i in order:
        tree_2.heading("日期", text=str(day.strftime('%m/%d')))
        tree_2.insert("",x ,text="" ,values=(text[i], college[i][0],college[i][1],college[i][2]))
        x += 1
    tree_2.insert("",7 ,text="" ,values=("總和", total[0], total[1], total[2]))
    st = 0

def outputfile(): 
    search()
    wb = load_workbook("所有系所名單.xlsx")
    sheet = wb['Sheet']
    wb2 = Workbook()
    sheet2 = wb2['Sheet']
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    day = datetime.strptime(date.get()+" 00:00:00", '%m/%d/%y %H:%M:%S')
    time_1 = timedelta(0, 0, 0, 0, 0, 0, 0)
    time_2 = timedelta(0, 0, 0, 0, 0, 13, 0)
    time_3 = timedelta(1, 0, 0, 0, 0, 0, 0)
    order = [0, 2, 4, 3, 1, 5, 6]
    for i in range(ord('A'),ord('D')+1):
        sheet2.column_dimensions[chr(i)].width = 23
    for i in range(1,10):
        sheet2.row_dimensions[i].height = 25
    sheet2.append([str(day.strftime('%m/%d')), '學士', '碩士', '博士'])
    for i in order:
        sheet2.append([text[i], college[i][0], college[i][1], college[i][2]])
    sheet2.append(["總和", total[0], total[1], total[2]])
    for row in sheet2.iter_rows(min_row=1, max_col=4):
        for cell in row:
            sheet2[cell.coordinate].alignment = Alignment(horizontal='center',vertical='center')
            sheet2[cell.coordinate].border = border
    for i in range(0,7):
        flat = True
        for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
            code = str(row[2]).lstrip('=RIGHT("')
            if row[6] != None:
                time = row[6] - day
                if (time > time_1 and time < time_2 and combo.current() == 0) or (time > time_2 and time < time_3 and combo.current() == 1) or (time > time_1 and time < time_3 and combo.current() == 2):
                    if int(code[0]) == order[i]+1:
                        if flat:
                            sheet2.append([""])
                            sheet2.append([text[order[i]]])
                            sheet2.append(["學號", "姓名", "系所", "電話"])
                            flat = False
                        sheet2.append([row[0], row[1], row[2], row[3]])

    for row in sheet2.iter_rows(min_row=1, max_col=4):
        for cell in row:
            sheet2[cell.coordinate].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            if sheet2[cell.coordinate].value in text and int(cell.coordinate[1:]) >= 10:
                sheet2.merge_cells('A'+cell.coordinate[1:]+':D'+cell.coordinate[1:])
                sheet2[cell.coordinate].border = border
            elif sheet2[cell.coordinate].value == '':
                break
            else:
                sheet2[cell.coordinate].border = border
    create_path("當日歸還統計資料")
    if combo.current() == 0:
        filename = "當日歸還統計資料/"+str(day.strftime('%m.%d'))+"上午歸還統計.xlsx"
    elif combo.current() == 1:
        filename = "當日歸還統計資料/"+str(day.strftime('%m.%d'))+"下午歸還統計.xlsx"
    elif combo.current() == 2:
        filename = "當日歸還統計資料/"+str(day.strftime('%m.%d'))+"全天歸還統計.xlsx"
    wb2.save(filename)
    tk.messagebox.showinfo(title = '成功',message='輸出成功')

def enter(self):
    search_s()

def search_s():
    delButton(tree_1)
    find, borrow, back = 0, 0, 0
    wb = load_workbook("所有系所名單.xlsx")
    sheet = wb['Sheet']
    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
        if str(row[0]) == ID.get():
            find = 1
            name = row[1]
            department = row[2]
            time = row[6]
            if row[4] == 1:
                borrow = 1
                if row[5] == 1:
                    back = 1
            break
    if find == 1:
        if borrow == 1:
            if back == 1:
                tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), time.strftime("%m/%d %H:%M"), '已歸還'))
            else:
                tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), "", '已借未還'))
        else:
            tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), "", '未借用'))
    else:
        tk.messagebox.showerror(title = 'error!',message='查無此人')

def borrow_s():
    delButton(tree_1)
    try:
        wb = load_workbook("所有系所名單.xlsx")
        sheet = wb['Sheet']
        find, borrow, back = 0, 0, 0
        coordi = 'A0'
        for row in sheet.iter_rows(min_row=2, max_col=7):
            if str(row[0].value) == ID.get():
                find = 1
                name = row[1].value
                if row[4].value == 1:
                    borrow = 1
                    if row[5].value == 1:
                        back = 1
                department = row[2].value
                time = row[6].value
                coordi = row[4].coordinate
                break
        if find == 1:
            if borrow == 1:
                if back == 1:
                    tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), time.strftime("%m/%d %H:%M"), '已歸還'))
                else:
                    tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), "", '已借未還'))
                tk.messagebox.showerror(title = 'error!',message='無法重複借用')
            else:
                sheet[coordi].value = 1
                wb.save('所有系所名單.xlsx')
                tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), "", '已借未還'))
                tk.messagebox.showinfo(title = 'confirm!',message='借用成功')
        else:
            tk.messagebox.showerror(title = 'error!',message='查無此人')
    except:
        tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), "", '借用失敗'))
        tk.messagebox.showerror(title = 'error!',message='請關閉excel檔案')

def revert_s():
    delButton(tree_1)
    try:
        wb = load_workbook("所有系所名單.xlsx")
        sheet = wb['Sheet']
        find, borrow, back = 0, 0, 0
        coordi = 'A0'
        for row in sheet.iter_rows(min_row=2, max_col=7):
            if str(row[0].value) == ID.get():
                find = 1
                name = row[1].value
                department = row[2].value
                time = row[6].value
                if row[4].value == 1:
                    borrow = 1
                    if row[5].value == 1:
                        back = 1
                    elif row[5].value == None:
                        coordi_1 = row[5].coordinate
                        coordi_2 = row[6].coordinate
                break
        if find == 1:
            if borrow == 1:
                if back == 1:
                    tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), time.strftime("%m/%d %H:%M"), '已歸還'))
                    tk.messagebox.showerror(title = 'error!',message='無法重複歸還')
                else:
                    sheet[coordi_1].value = 1
                    sheet[coordi_2].value = datetime.now()
                    wb.save('所有系所名單.xlsx')
                    tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), datetime.now().strftime("%m/%d %H:%M"), '已歸還'))
                    tk.messagebox.showinfo(title = 'confirm!',message='已歸還成功')
            else:
                tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), "", '未借用'))
                tk.messagebox.showerror(title = 'error!',message='此人未借用')
        else:
            tk.messagebox.showerror(title = 'error!',message='查無此人')
    except:
        tree_1.insert("",1 ,text="" ,values=(ID.get(), name, department[12:-4].rstrip('"'), "", '歸還失敗'))
        tk.messagebox.showerror(title = 'error!',message='請關閉excel檔案')

def delButton(tree):
    x = tree.get_children()
    for item in x:
        tree.delete(item)

def content(event):
    item = tree_2.selection()
    if item:
        txt = tree_2.item(item[0],'values')
        if txt[0] == '文學院':
            separate(1)
        elif txt[0] == '理學院':
            separate(2)
        elif txt[0] == '社科院':
            separate(3)
        elif txt[0] == '工學院':
            separate(4)
        elif txt[0] == '管理學院':
            separate(5)
        elif txt[0] == '法學院':
            separate(6)
        elif txt[0] == '教育學院':
            separate(7)

def separate(i):
    x = 0
    top = tk.Toplevel(window)
    columns = ("學號", "系所", "姓名", "電話", "備註")
    treeview = ttk.Treeview(top, height=10, show="headings", columns=columns)
    treeview.column("學號", width=150, anchor='center')
    treeview.column("姓名", width=120, anchor='center')
    treeview.column("系所", width=630, anchor='center')
    treeview.column("電話", width=160, anchor='center')
    treeview.column("備註", width=70, anchor='center')
    treeview.heading("學號", text="學號")
    treeview.heading("姓名", text="姓名")
    treeview.heading("系所", text="系所")
    treeview.heading("電話", text="電話")
    treeview.heading("備註", text="備註")
    treeview.place(relx=0.004, rely=0.028, relwidth=0.964, relheight=0.95)
    VScroll1 = tk.Scrollbar(top, orient='vertical', command=treeview.yview)
    VScroll1.place(relx=0.99, rely=0.028, relwidth=0.015, relheight=0.958)
    treeview.configure(yscrollcommand=VScroll1.set)
    treeview.pack()
    wb = load_workbook("所有系所名單.xlsx")
    sheet = wb['Sheet']
    day = datetime.strptime(date.get()+" 00:00:00", '%m/%d/%y %H:%M:%S')
    time_1 = timedelta(0, 0, 0, 0, 0, 0, 0)
    time_2 = timedelta(0, 0, 0, 0, 0, 13, 0)
    time_3 = timedelta(1, 0, 0, 0, 0, 0, 0)
    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
        code = str(row[2]).lstrip('=RIGHT("')
        if row[6] != None:
            time = row[6] - day
            if (time > time_1 and time < time_2 and combo.current() == 0) or (time > time_2 and time < time_3 and combo.current() == 1) or (time > time_1 and time < time_3 and combo.current() == 2):
                if int(code[0]) == i:
                    treeview.insert("",x ,text="" ,values=(row[0], row[2][12:-4].rstrip('"'), row[1], row[3],''))
                    x+=1

def output_no_return():
    reset_st()
    wb = load_workbook("所有系所名單.xlsx")
    sheet = wb['Sheet']
    wb2 = Workbook()
    wb3 = Workbook()
    wb4 = Workbook()
    wb5 = Workbook()
    sheet2 = wb2['Sheet']
    sheet3 = wb3['Sheet']
    sheet4 = wb4['Sheet']
    sheet5 = wb5['Sheet']
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    department = []
    number_people = []
    day = datetime.strptime(date.get()+" 00:00:00", '%m/%d/%y %H:%M:%S')
    order = [0, 2, 4, 3, 1, 5, 6]
    sheet2.column_dimensions['A'].width = 10.5
    sheet2.column_dimensions['B'].width = 10
    sheet2.column_dimensions['C'].width = 50
    sheet2.column_dimensions['D'].width = 11
    sheet3.column_dimensions['A'].width = 8
    sheet3.column_dimensions['B'].width = 12
    sheet3.column_dimensions['C'].width = 12
    sheet3.column_dimensions['D'].width = 8
    sheet2.append(['學號', '姓名', '系所名稱', '電話'])
    sheet3.append(['sno', 'std_no', 'rem', 'yn'])
    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
        code = row[2].split('"')
        if row[4] == 1:
            if row[5] == None:
                if code[1] not in department:
                    department.append(code[1])
                    number_people.append(0)
                number_people[department.index(code[1])]+=1
                sheet2.append([row[0], row[1], row[2], row[3]])
                sheet3.append(['10', row[0], '禮服未歸還', 'N'])
                for i in range (1,8):
                    if int(code[1][0]) == i:
                        if isinstance(row[0], int):
                            x = list(row)
                            x[0] = str(row[0])
                            row = tuple(x)
                        if str(row[0][0]) == '4':
                            college[i-1][0]+=1
                            total[0]+=1
                        elif str(row[0][0]) == '6' or str(row[0][0]) == '5':
                            college[i-1][1]+=1
                            total[1]+=1
                        elif str(row[0][0]) == '8':
                            college[i-1][2]+=1
                            total[2]+=1
    for i in range(ord('A'),ord('D')+1):
        sheet4.column_dimensions[chr(i)].width = 23
    for i in range(1,12):
        sheet4.row_dimensions[i].height = 25
    sheet4.append(["未歸還統計人數"])
    sheet4.append([str(day.strftime('%m/%d')), '學士', '碩士', '博士'])
    for i in order:
        sheet4.append([text[i], college[i][0], college[i][1], college[i][2]])
    sheet4.append(["總和", total[0], total[1], total[2]])
    sheet4.append(["未歸還總和", total[0]+total[1]+total[2]])
    sheet4.merge_cells('A1:D1')
    sheet4.merge_cells('B11:D11')
    for row in sheet2.iter_rows(min_row=1, max_col=4):
        for cell in row:
            sheet2[cell.coordinate].alignment = Alignment(horizontal='center',vertical='center')
            sheet3[cell.coordinate].alignment = Alignment(horizontal='center',vertical='center')  
    for row in sheet4.iter_rows(min_row=1, max_col=4):
        for cell in row:
            sheet4[cell.coordinate].alignment = Alignment(horizontal='center',vertical='center')
            sheet4[cell.coordinate].border = border

    sheet5.column_dimensions['A'].width = 55
    sheet5.row_dimensions[1].height = 25
    sheet5.row_dimensions[2].height = 25
    for i in range(3,len(department)+3):
        sheet5.row_dimensions[i].height = 18
    sheet5.append([str(datetime.today().strftime('%m/%d'))+'未歸還統計人數'])
    sheet5.append(["系所名稱","人數"])
    sheet5.merge_cells('A1:B1')
    for i in department:
        sheet5.append([i[4:],number_people[department.index(i)]])
    for row in sheet5.iter_rows(min_row=1, max_col=2):
        for cell in row:
            sheet5[cell.coordinate].alignment = Alignment(horizontal='center',vertical='center')
            sheet5[cell.coordinate].border = border
    create_path("未歸還名單")
    wb2.save('未歸還名單/'+str(datetime.today().strftime('%m.%d'))+'未歸還名單.xlsx')
    wb3.save('未歸還名單/'+str(datetime.today().strftime('%m.%d'))+'離校系統未歸還名冊.xlsx')
    wb4.save('未歸還名單/'+str(datetime.today().strftime('%m.%d'))+'未歸還統計人數(學院).xlsx')
    wb5.save('未歸還名單/'+str(datetime.today().strftime('%m.%d'))+'未歸還統計人數(系).xlsx')
    tk.messagebox.showinfo(title = '成功',message='輸出成功')

def output_borrow():
    reset_st()
    wb = load_workbook("所有系所名單.xlsx")
    sheet = wb['Sheet']
    wb4 = Workbook()
    wb5 = Workbook()
    sheet4 = wb4['Sheet']
    sheet5 = wb5['Sheet']
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    order = [0, 2, 4, 3, 1, 5, 6]
    department = []
    number_people = []
    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
        code = row[2].split('"')
        if row[4] == 1:
            if code[1] not in department:
                department.append(code[1])
                number_people.append(0)
            number_people[department.index(code[1])]+=1
            for i in range (1,8):
                if int(code[1][0]) == i:
                    if isinstance(row[0], int):
                        x = list(row)
                        x[0] = str(row[0])
                        row = tuple(x)
                    if str(row[0][0]) == '4':
                        college[i-1][0]+=1
                        total[0]+=1
                    elif str(row[0][0]) == '6':
                        college[i-1][1]+=1
                        total[1]+=1
                    elif str(row[0][0]) == '8':
                        college[i-1][2]+=1
                        total[2]+=1
                    elif str(row[0][0]) == '5':
                        college[i-1][3]+=1
                        total[3]+=1
    for i in range(ord('A'),ord('E')+1):
        sheet4.column_dimensions[chr(i)].width = 18.5
    for i in range(1,12):
        sheet4.row_dimensions[i].height = 25
    sheet4.append(["借用統計人數"])
    sheet4.append(['', '學士', '碩士', '專班', '博士'])
    for i in order:
        sheet4.append([text[i], college[i][0], college[i][1], college[i][3], college[i][2]])
    sheet4.append(["總和", total[0], total[1], total[3], total[2]])
    sheet4.append(["借用總和", total[0]+total[1]+total[2]+total[3]]) 
    sheet4.merge_cells('A1:E1')
    sheet4.merge_cells('B11:E11')
    for row in sheet4.iter_rows(min_row=1, max_col=5):
        for cell in row:
            sheet4[cell.coordinate].alignment = Alignment(horizontal='center',vertical='center')
            sheet4[cell.coordinate].border = border
    
    sheet5.column_dimensions['A'].width = 55
    sheet5.row_dimensions[1].height = 25
    sheet5.row_dimensions[2].height = 25
    for i in range(3,len(department)+3):
        sheet5.row_dimensions[i].height = 18
    sheet5.append(["借用統計人數"])
    sheet5.append(["系所名稱","人數"])
    sheet5.merge_cells('A1:B1')
    for i in department:
        sheet5.append([i[4:],number_people[department.index(i)]])
    for row in sheet5.iter_rows(min_row=1, max_col=2):
        for cell in row:
            sheet5[cell.coordinate].alignment = Alignment(horizontal='center',vertical='center')
            sheet5[cell.coordinate].border = border
    create_path("借用統計人數")
    wb4.save('借用統計人數/借用統計人數(學院).xlsx')
    wb5.save('借用統計人數/借用統計人數(系).xlsx')
    tk.messagebox.showinfo(title = '成功',message='輸出成功')

def main():
    bt_1 = tk.Button(window, text = "查詢", bg="#96FED1", command = search_s, font = ('Arial', 24), width = 7, height = 1)
    bt_2 = tk.Button(window, text = "借用", bg="#96FED1", command = borrow_s, font = ('Arial', 24), width = 7, height = 1)
    bt_3 = tk.Button(window, text = "歸還", bg="#96FED1", command = revert_s, font = ('Arial', 24), width = 7, height = 1)
    bt_4 = tk.Button(window, text = "歸還統計", bg="#96FED1", command = search, font = ('Arial', 18), width = 10, height = 1)
    bt_5 = tk.Button(window, text = "列印歸還名單", bg="#96FED1", command = outputfile, font = ('Arial', 18), width = 10, height = 1)
    bt_6 = tk.Button(window, text = "輸出未還統計", bg="#96FED1", command = output_no_return, font = ('Arial', 18), width = 10, height = 1)
    bt_7 = tk.Button(window, text = "輸出借用統計", bg="#96FED1", command = output_borrow, font = ('Arial', 18), width = 10, height = 1)
    lbl_1 = tk.Label(window, text = '學號：', bg="#7AFEC6", fg = 'black', font = ('Arial', 30), width = 5, height = 2)
    ent_1 = tk.Entry(window, width = 20 ,textvariable = ID ,bd = 3, font = ('Arial', 24))
    cal = Calendar(window, font="Arial 14", selectmode='day', cursor="hand1", textvariable = date)
    window.bind('<Return>',enter)

    tree_1.column("學號", width=150, anchor='center')
    tree_1.column("姓名", width=120, anchor='center')
    tree_1.column("系所", width=450, anchor='center')
    tree_1.column("歸還時間", width=180, anchor='center')
    tree_1.column("歸還狀況", width=125, anchor='center')
    tree_1.heading("學號", text="學號")
    tree_1.heading("姓名", text="姓名")
    tree_1.heading("系所", text="系所")
    tree_1.heading("歸還時間", text="歸還時間")
    tree_1.heading("歸還狀況", text="歸還狀況")

    tree_2.column("日期", width=120, anchor='center')
    tree_2.column("學士", width=80, anchor='center')
    tree_2.column("碩士", width=80, anchor='center')
    tree_2.column("博士", width=80, anchor='center')
    tree_2.heading("日期", text="")
    tree_2.heading("學士", text="學士")
    tree_2.heading("碩士", text="碩士")
    tree_2.heading("博士", text="博士")
    num = [0, 2, 4, 3, 1, 5, 6]
    x = 0
    for i in num:
        tree_2.insert("",x ,text="" ,values=(text[i]))
        x += 1
    tree_2.insert("",7 ,text="" ,values=("總和"))

    lbl_1.grid(column = 0, row = 0)
    bt_1.grid(column = 2, row = 0, sticky = 'w', padx=40)
    bt_2.grid(column = 2, row = 0, sticky = 'e')
    bt_3.grid(column = 3, row = 0, sticky = 'w', padx=8)
    bt_4.grid(column = 3, row = 3, sticky = 'w', pady=20)
    bt_5.grid(column = 3, row = 4, sticky = 'wn', pady=20)
    bt_6.grid(column = 3, row = 5, sticky = 'wn', pady=20)
    bt_7.grid(column = 3, row = 5, sticky = 'ws', pady=20)
    ent_1.grid(column = 1, row = 0, sticky = 'w', ipady=10)
    tree_1.grid(column = 0, row = 1, sticky = 'w', padx=10, columnspan = 4)
    tree_2.grid(column = 2, row = 3, sticky = 'w', padx=20, pady=20, columnspan = 1, rowspan = 3)
    cal.grid(column = 0, row = 3, pady=20, columnspan = 2, rowspan = 3)
    combo.grid(column = 3, row = 3, sticky = 'n', pady=20)
    # combo.current()
    tree_2.bind('<ButtonRelease-1>', content)
    combo.current(0)
    
    window.mainloop()

if __name__ == '__main__':
    main()