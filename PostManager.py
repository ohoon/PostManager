from tkinter import messagebox
from urllib import request as url
from bs4 import BeautifulSoup
import tkinter as tk
import tkinter.ttk
import re
import webbrowser

def start():
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
    global count
    global rcount

    def add():
        global count
        global rcount

        try:
            s = int(r_firstE.get())
            e = int(r_lastE.get())
            if s > e:
                raise ValueError
            if len(str(s)) != 13 or len(str(e)) != 13:
                raise ValueError
        except:
            messagebox.showinfo("Error!", "올바른 값을 입력해주세요.")
        else:
            for pnum in range(s, e + 1):
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
                        succT.insert("", "end", text="", values=[pnum, name, s_status, date], iid=pnum)
                        count += 1
                    elif status == "배달준비":
                        retrnT.insert("", "end", text="", values=[pnum, name, s_status, date], iid=pnum)
                        rcount += 1
                    elif status == "" or status == "도착":
                        retrnT.insert("", "end", text="", values=[pnum, name, s_status, date], iid=pnum)
                        rcount += 1
                    elif status == "미배달":
                        r_data = re.findall("<!-- 미배달사유 -->.*\s*.*\s*.*<!-- //미배달사유 -->", str(soup))
                        r_status = re.sub("<.+?>|\s", '', r_data[len(r_data) - 1])
                        retrnT.insert("", "end", text="", values=[pnum, name, r_status, date], iid=pnum)
                        rcount += 1

            resultL.configure(text="배달완료 " + str(count) + "건\t미배달 " + str(rcount) + "건\t총 " + str(count + rcount) + "건")

    def pop_up_s(event):
        webbrowser.open("https://service.epost.go.kr/trace.RetrieveDomRigiTraceList.comm?sid1=" + str(succT.focus()) + "&displayHeader=N")
    def pop_up_r(event):
        webbrowser.open("https://service.epost.go.kr/trace.RetrieveDomRigiTraceList.comm?sid1=" + str(retrnT.focus()) + "&displayHeader=N")

    def sort_column(tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

            # reverse sort next time
        tv.heading(col, command=lambda: sort_column(tv, col, not reverse))

    resultC = tk.Toplevel(mainC)
    resultC.title("조회 결과")
    resultC.geometry("736x460+495+300")
    resultC.resizable(0, 0)

    succF = tk.Frame(resultC, bd=5)
    succF.grid(column=0, columnspan=2, row=0, rowspan=4)
    succL = tk.Label(succF, text="배달완료", font="함초롬돋움 12")
    succL.grid(column=0, columnspan=2, row=0)

    succT = tkinter.ttk.Treeview(succF, columns=["", "", "", ""], height=18)
    succT.grid(column=0, columnspan=2, row=1)
    succT.yview()
    succT["show"] = "headings"
    succT.column("#1", width=120, anchor="center")
    succT.heading("#1", text="등기번호", anchor="center", command=lambda: sort_column(succT, "#1", False))
    succT.column("#2", width=45, anchor="center")
    succT.heading("#2", text="성명", anchor="center", command=lambda: sort_column(succT, "#2", False))
    succT.column("#3", width=110, anchor="center")
    succT.heading("#3", text="수령인", anchor="center", command=lambda: sort_column(succT, "#3", False))
    succT.column("#4", width=80, anchor="center")
    succT.heading("#4", text="수령일자", anchor="center", command=lambda: sort_column(succT, "#4", False))
    succT.bind('<<TreeviewOpen>>', pop_up_s)

    retrnF = tk.Frame(resultC, bd=5)
    retrnF.grid(column=2, columnspan=2, row=0, rowspan=4)
    retrnL = tk.Label(retrnF, text="미배달", font="함초롬돋움 12")
    retrnL.grid(column=0, columnspan=2, row=0)

    retrnT = tkinter.ttk.Treeview(retrnF, columns=["", "", "", ""], height=18)
    retrnT.grid(column=0, columnspan=2, row=1)
    retrnT.yview()
    retrnT["show"] = "headings"
    retrnT.column("#1", width=120, anchor="center")
    retrnT.heading("#1", text="등기번호", anchor="center", command=lambda: sort_column(retrnT, "#1", False))
    retrnT.column("#2", width=45, anchor="center")
    retrnT.heading("#2", text="성명", anchor="center", command=lambda: sort_column(retrnT, "#2", False))
    retrnT.column("#3", width=110, anchor="center")
    retrnT.heading("#3", text="처리현황", anchor="center", command=lambda: sort_column(retrnT, "#3", False))
    retrnT.column("#4", width=80, anchor="center")
    retrnT.heading("#4", text="처리일자", anchor="center", command=lambda: sort_column(retrnT, "#4", False))
    retrnT.bind('<<TreeviewOpen>>', pop_up_r)

    count = 0
    rcount = 0
    for pnum in range(f, l + 1):
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
                succT.insert("", "end", text="", values=[pnum, name, s_status, date], iid=pnum)
                count += 1
            elif status == "배달준비":
                retrnT.insert("", "end", text="", values=[pnum, name, s_status, date], iid=pnum)
                rcount += 1
            elif status == "" or status == "도착":
                retrnT.insert("", "end", text="", values=[pnum, name, s_status, date], iid=pnum)
                rcount += 1
            elif status == "미배달":
                r_data = re.findall("<!-- 미배달사유 -->.*\s*.*\s*.*<!-- //미배달사유 -->", str(soup))
                r_status = re.sub("<.+?>|\s", '', r_data[len(r_data) - 1])
                retrnT.insert("", "end", text="", values=[pnum, name, r_status, date], iid=pnum)
                rcount += 1

    resultF = tk.Frame(resultC)
    resultF.grid(column=0, columnspan=4, row=4)
    resultL = tk.Label(resultF, text="배달완료 " + str(count) + "건\t미배달 " + str(rcount) + "건\t총 " + str(count + rcount) + "건", font="함초롬돋움 10", anchor="center", padx=25)
    resultL.grid(column=0, row=0)
    r_firstE = tk.Entry(resultF)
    r_firstE.grid(column=1, row=0, padx=5)
    r_Label = tk.Label(resultF, text="~")
    r_Label.grid(column=2, row=0)
    r_lastE = tk.Entry(resultF)
    r_lastE.grid(column=3, row=0, padx=5)
    resultB = tk.Button(resultF, text="이어서 조회", width=10, font="함초롬돋움 10", command=add)
    resultB.grid(column=4, row=0)

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

mainB = tk.Button(mainC, text="조회하기", width=15, font="함초롬돋움 13", command=start)
mainB.grid(column=0, columnspan=2, row=3)
mainC.mainloop()
