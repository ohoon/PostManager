from tkinter import messagebox
from urllib import request as url
from bs4 import BeautifulSoup
import tkinter as tk
import tkinter.ttk
import re
import webbrowser


def search():
    try:
        fst = int(firstE.get())
        lst = int(lastE.get())
        if fst > lst:
            raise ValueError
        if len(str(fst)) != 13 or len(str(lst)) != 13:
            raise ValueError
    except:
        messagebox.showinfo("Error!","올바른 값을 입력해주세요.")
    else:
        result(fst, lst)


def result(f, l):

    def pop_up_s(event):
        webbrowser.open("https://service.epost.go.kr/trace.RetrieveDomRigiTraceList.comm?sid1=" + str(succT.focus()) + "&displayHeader=N")
    def pop_up_r(event):
        webbrowser.open("https://service.epost.go.kr/trace.RetrieveDomRigiTraceList.comm?sid1=" + str(retrnT.focus()) + "&displayHeader=N")

    resultC = tk.Toplevel(mainC)
    resultC.title("조회 결과")
    resultC.geometry("736x450+495+300")
    resultC.resizable(0, 0)

    succF = tk.Frame(resultC, bd=5)
    succF.grid(column=0, columnspan=2, row=0, rowspan=4)
    succL = tk.Label(succF, text="배달완료", font="함초롬돋움 12")
    succL.grid(column=0, columnspan=2, row=0)

    succT = tkinter.ttk.Treeview(succF, columns=["", "", ""], height=18)
    succT.grid(column=0, columnspan=2, row=1)
    succT.yview()
    succT.column("#0", width=120, anchor="center")
    succT.heading("#0", text="등기번호", anchor="center")
    succT.column("#1", width=45, anchor="center")
    succT.heading("#1", text="성명", anchor="center")
    succT.column("#2", width=110, anchor="center")
    succT.heading("#2", text="수령인", anchor="center")
    succT.column("#3", width=80, anchor="center")
    succT.heading("#3", text="수령일자", anchor="center")
    succT.bind('<<TreeviewOpen>>', pop_up_s)

    retrnF = tk.Frame(resultC, bd=5)
    retrnF.grid(column=2, columnspan=2, row=0, rowspan=4)
    retrnL = tk.Label(retrnF, text="미배달", font="함초롬돋움 12")
    retrnL.grid(column=0, columnspan=2, row=0)

    retrnT = tkinter.ttk.Treeview(retrnF, columns=["", "", ""], height=18)
    retrnT.grid(column=0, columnspan=2, row=1)
    retrnT.yview()
    retrnT.column("#0", width=120, anchor="center")
    retrnT.heading("#0", text="등기번호", anchor="center")
    retrnT.column("#1", width=45, anchor="center")
    retrnT.heading("#1", text="성명", anchor="center")
    retrnT.column("#2", width=110, anchor="center")
    retrnT.heading("#2", text="처리현황", anchor="center")
    retrnT.column("#3", width=80, anchor="center")
    retrnT.heading("#3", text="처리일자", anchor="center")
    retrnT.bind('<<TreeviewOpen>>', pop_up_r)

    resultF = tk.Frame(resultC)
    resultF.grid(column=0, columnspan=4, row=4)
    resultL = tk.Label(resultF, text="", font="함초롬돋움 10", anchor="center")
    resultL.grid(column=0, row=0)

    count = 0
    rcount = 0
    for pnum in range(f, l+1):
        with url.urlopen("https://service.epost.go.kr/trace.RetrieveDomRigiTraceList.comm?sid1=" + str(
                pnum) + "&displayHeader=N") as doc:
            html = doc.read().decode()
            soup = BeautifulSoup(html, features="lxml")
            data = soup.find("table", "table_col").find("tbody").find("tr").find_all("td")
            name = str(data[1])[4:-5].split("<br/>")[0]
            date = str(data[1])[4:-5].split("<br/>")[1]
            status = re.sub("\(.*\)", '', data[3].get_text())

            if status == "배달완료" or status == "교부":
                s_data = re.findall("\(수령인:.*\)", str(soup))
                s_status = re.sub("\(수령인:|\)", '', s_data[0])
                succT.insert("", "end", text=pnum, values=[name, s_status, date], iid=pnum)
                count += 1
            elif status == "배달준비":
                retrnT.insert("", "end", text=pnum, values=[name, status, date], iid=pnum)
                rcount += 1
            elif status == "" or status == "도착":
                retrnT.insert("", "end", text=pnum, values=[name, status, date], iid=pnum)
                rcount += 1
            elif status == "미배달":
                r_data = re.findall("<!-- 미배달사유 -->.*\s*.*\s*.*<!-- //미배달사유 -->", str(soup))
                r_status = re.sub("<.+?>|\s", '', r_data[len(r_data)-1])
                retrnT.insert("", "end", text=pnum, values=[name, r_status, date], iid=pnum)
                rcount += 1
    resultL.configure(text="배달완료 " + str(count) + "건\t미배달 " + str(rcount) + "건\t총 " + str(count+rcount) + "건")

mainC = tk.Tk()
mainC.title("등기번호 조회기")
mainC.geometry("230x130+250+300")
mainC.resizable(0, 0)

mainL = tk.Label(mainC, text="등기번호(13자리)", font="함초롬돋움 15")
mainL.grid(column=0, columnspan=2, row=0)

firstL = tk.Label(mainC, text="시작 번호\t", font="함초롬돋움 12")
firstL.grid(column=0, row=1)
firstE = tk.Entry(mainC)
firstE.grid(column=1, row=1)

lastL = tk.Label(mainC, text=" 끝 번호\t", font="함초롬돋움 12")
lastL.grid(column=0, row=2)
lastE = tk.Entry(mainC)
lastE.grid(column=1, row=2)

mainB = tk.Button(mainC, text="조회하기", width=15, font="함초롬돋움 13", command=search)
mainB.grid(column=0, columnspan=2, row=3)
mainC.mainloop()
