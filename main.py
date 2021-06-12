# coding=<utf-8>
import pandas as pd
import os
import tkinter.ttk
import tkinter.messagebox
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from win32com.client import Dispatch


days = 0
times = []
gclass=[]
max_class=[]
max_class_stack = [0]

def getsubject():
    global subject
    subject = []

    check = 0
    for i in classplacement.head(0):
        if i == "학번":
            check += 1
        if check == 1:
            if i[0:8] == "Unnamed:":
                break
            subject.append(i)
########################################################################################
def makedataFrame():
    global times
    global days
    global gclass
    global max_class
    global max_class_stack
    global classplacement

    days = 0
    times = []
    gclass = []
    max_class = []
    max_class_stack = [0]
    try:
        readfile = askopenfilename()
        classplacement = pd.read_excel(readfile,sheet_name="학생별 반배정", header=1, dtype=str)
        classplacement = classplacement.fillna("")

        time = pd.read_excel(readfile, sheet_name="교실별 과목배정", header=0)
        time = time.fillna("")

        list = []#임시
        before = classplacement["학번"][0]
        for i in classplacement["학번"]:  # 인덱스(학번)
            if i == "":
                gclass.append(list)
                break
            if i[0:3] != before[0:3]:
                gclass.append(list)
                list = []
                list.append(i)
            else:
                list.append(i)
            before = i

        for i in time["교시"]:
            if str(type(i)) == "<class 'int'>":
                times.append(i)
            elif str(type(i)) == "<class 'float'>":
                times.append(int(i))

        getsubject()
        days = times.count(1)

        for i in range(len(times)):
            if i != 0 and times[i] == 1:
                max_class.append(times[i - 1])
                max_class_stack.append(i)

        max_class.append(times[-1])
        max_class_stack.append(len(times))

        listboxin()
    except:
        return
########################################################################################
def makeban():
    global ban
    global class_seat

    ban = classplacement.copy()
    class_seat = classplacement.copy()
    classplace_by_gclass = classplacement.set_index('학번')
    Ban=[]

    for col in subject[2:]:    #교실 추출
        for i in gclass:
            for j in i:
                a = classplace_by_gclass.loc[j][col]
                if a not in Ban:
                    if a != "":
                        Ban.append(a)
        for B in Ban:  # 반
            seat = 1
            for i in ban[ban[col].isin([B])].index:
                ban.at[i,col] = seat
                class_seat.at[i, col] = str(classplacement.at[i, col]) + "/좌석:" + str(ban.at[i, col])
                seat += 1
########################################################################################
def listboxin():#동시에 치는 과목 선택용 리스트 박스에 과목 불러오기
    listbox.delete(0, END)
    for i in subject[2:]:
        listbox.insert(END,i)
########################################################################################
def overrap():  # btn2 command 동시에 치는 과목 묶기
    global subject
    try:
        over_sub = []

        for i in listbox.curselection():
            over_sub.append(subject[i+2])

        bool = 0
        new = []

        for i in range(len(classplacement[over_sub[0]].values.tolist())):
            combine = ""
            temp = ""

            for j in range(len(over_sub)):
                bool += (classplacement[over_sub[0]][i] == classplacement[over_sub[j]][i])
                combine += str(classplacement[over_sub[j]][i])
                if str(classplacement[over_sub[j]][i]) != "":
                    temp = classplacement[over_sub[j]][i]
            if combine == "":
                break
            elif bool == len(over_sub):
                new.append(classplacement[over_sub[0]][i])
            else:
                new.append(temp)

        combine = ""
        for i in over_sub:
            combine += i + "/"

        combine = combine[0:-1]

        df = pd.DataFrame(new, columns=[combine])
        classplacement[over_sub[0]] = df[combine]
        classplacement.rename(columns={over_sub[0]: combine}, inplace=True)

        for i in range(len(over_sub)-1):
            del classplacement[over_sub[i + 1]]

        getsubject()
        listboxin()
    except:
        tkinter.messagebox.showerror(title="ERROR!", message="합칠 과목이 없습니다!")
########################################################################################
def write(): #데이터프레임 으로 엑셀 쓰기
    makeban()
    writefile = askdirectory()
    daypd = pd.DataFrame(data=['1일차', ' ', '2일차', ' ', '3일차', ' ', '4일차', ' ', '5일차', ' ', '6일차', ' ', '7일차', ' ', '8일차', ' ', '9일차',' ', '10일차', ' ', '11일차', ' ', '12일차', ' '],index=[' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',' ', ' ', ' '], columns=[' '])
    day2 =days*2
    if(writefile != ''):
        Writefile = writefile
        try:
            for i in gclass:
                writefile = Writefile
                writefile = writefile + '/' + str(i[0])[0:3] + '.xlsx'
                with pd.ExcelWriter(writefile) as writer:
                    for j in i:
                        class_seat.loc[:, subject[0:2]][class_seat.학번 == f'{j}'].to_excel(writer, sheet_name=f'{j}',startcol=1,index=FALSE)
                        for t in range(days):
                            if(t==0):
                                daypd[:day2].to_excel(writer, sheet_name=f'{j}', startrow=2 * (t + 1),startcol=0,index=FALSE)
                            class_seat.loc[:, subject[max_class_stack[t] + 2: max_class_stack[t + 1] + 2]][class_seat.학번 == f'{j}'].to_excel(writer, sheet_name=f'{j}', startrow=2 * (t + 1),startcol=1,index=FALSE)

            tkinter.messagebox.showinfo(title="완료", message="파일 생성이 완료 되었습니다")
        except:
            return

########################################################################################
def student_check(): #조회부분
    try:
        makeban()
        getsubject()
        def table_set():
            try:
                table.delete(*table.get_children())
                stu_num = txtb1.get()
                classplace_by_gclass = classplacement.set_index('학번')
                banplace_by_gclass = ban.set_index('학번')

                lbl1.configure(text="학번 : "+stu_num)
                lbl2.configure(text="이름 : "+classplace_by_gclass.loc[stu_num,'이름'])

                count = 0
                for i in range(len(times)):
                    if times[i] == 1:
                        count += 1
                    treelist = (subject[i + 2],classplace_by_gclass.loc[stu_num, subject[i + 2]], banplace_by_gclass.loc[stu_num, subject[i + 2]])
                    table.insert('', 'end',text=str(count) + "일차 " + str(times[i]) + "교시",values=treelist)
            except:
                lbl1.configure(text="학번 :")
                lbl2.configure(text="이름 :")
                tkinter.messagebox.showerror(title="ERROR!", message="없는 번호 입니다.")
        def printone():
            try:
                table.delete(*table.get_children())
                writefile = 'C:/' + str(txtb1.get()) + '.xlsx'
                sheet = []
                daypd = pd.DataFrame(data=['1일차', ' ', '2일차', ' ', '3일차', ' ', '4일차', ' ', '5일차', ' ', '6일차', ' ', '7일차', ' ', '8일차',' ', '9일차', ' ', '10일차', ' ', '11일차', ' ', '12일차', ' '],index=[' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',' ', ' ', ' ', ' ', ' ', ' '], columns=[' '])
                day2 = days * 2
                with pd.ExcelWriter(writefile) as writer:
                    j=txtb1.get()
                    class_seat.loc[:, subject[0:2]][class_seat.학번 == f'{j}'].to_excel(writer, sheet_name=f'{j}',startcol=1,index=FALSE)
                    for t in range(days):
                        if (t == 0):
                            daypd[:day2].to_excel(writer, sheet_name=f'{j}', startrow=2 * (t + 1), index=FALSE)
                        class_seat.loc[:, subject[max_class_stack[t] + 2: max_class_stack[t + 1] + 2]][class_seat.학번 == f'{j}'].to_excel(writer, sheet_name=f'{j}', startrow=2 * (t + 1),startcol=1,index=FALSE)
                        sheet.append(j)
                excel = Dispatch('Excel.Application')
                excel.Visible = False
                wb = excel.Workbooks.Open(writefile)
                for sheetnum in sheet:
                    excel.Worksheets(str(sheetnum)).Activate()
                    excel.ActiveSheet.Columns.AutoFit()
                wb.Save()
                wb.Close()
                os.startfile(writefile,"print")
                tkinter.messagebox.showinfo(title="인쇄 알림", message="출력이 시작됩니다.")
                os.remove(writefile)
                table.delete(*table.get_children())
            except:
                lbl1.configure(text="학번 :")
                lbl2.configure(text="이름 :")
                tkinter.messagebox.showerror(title="ERROR!", message="없는 번호 입니다.")

        new_window = Toplevel(root)
        new_window.resizable(width=FALSE, height=FALSE)
        lbl1 = Label(new_window, text="학번 :")
        lbl2 = Label(new_window, text="이름 :")
        lbl3 = Label(new_window, text="")
        lbl1.grid(row=0,column=0)
        lbl2.grid(row=0,column=1)
        lbl3.grid(row=0, column=2)

        table = tkinter.ttk.Treeview(new_window, columns=["one", "two","three"], displaycolumns=["one","two","three"])

        table.column("#0", width=100, anchor="center")
        table.heading("#0", text="교시", anchor="center")

        table.column("#1", width=150, anchor="center")
        table.heading("one", text="과목", anchor="center")

        table.column("#2", width=100, anchor="center")
        table.heading("two", text="반", anchor="center")

        table.column("#3", width=70, anchor="center")
        table.heading("three", text="좌석번호", anchor="center")

        table.grid(row=1, column =0, columnspan=3)

        lbl4 = Label(new_window,text="학번 : ex)30101")
        lbl4.grid(row=2,column=0)

        txtb1 = Entry(new_window)
        txtb1.grid(row=2, column =1)

        btn1 = Button(new_window, text="확인", width=15, command=table_set)
        btn1.grid(row=2, column=2)

        btn2 = Button(new_window, text="인쇄", width=15, command=printone)
        btn2.grid(row=3, column=1)


        new_window.mainloop()
    except:
        tkinter.messagebox.showerror(title="ERROR!", message="ERROR!")
########################################################################################
root = Tk()
root.geometry("400x445")
frame = Frame(root)

root.resizable(width=FALSE, height=FALSE)
root.title("김해고등학교 지필평가 개인별 시간표 생성")

scrollbary = Scrollbar(frame)
scrollbary.pack(side="right", fill="y")

scrollbarx = Scrollbar(frame,orient=HORIZONTAL)
scrollbarx.pack(side="bottom", fill="x")

listbox =  Listbox(frame,height=25, width=30,selectmode="multiple", xscrollcommand = scrollbarx.set, yscrollcommand = scrollbary.set)
listbox.pack(side="left")

scrollbary.config(command=listbox.yview)
scrollbarx.config(command=listbox.xview)

frame.grid(row=0,column=0,rowspan=4,padx=5,pady=5)

btn1 = Button(root, text = "파일 가져오기",height=3,width=20,command=makedataFrame)
btn2 = Button(root, text = "동시에 치는 과목 묶기",height=3,width=20,command=overrap)
btn3 = Button(root, text="시간표 작성하기",height=3,width=20,command=write)
btn4 = Button(root, text="학생 개별 시간표 확인",height=3,width=20,command=student_check)

btn1.grid(row=0,column=2)
btn2.grid(row=1,column=2)
btn3.grid(row=2,column=2)
btn4.grid(row=3,column=2)
root.mainloop()