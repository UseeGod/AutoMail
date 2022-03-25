import os
import tkinter.messagebox as msgbox
from tkinter import *
import tkinter.ttk as ttk
from tkinter import filedialog
import smtplib
from email.encoders import encode_base64
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
import sys
# -*- coding:utf-8 -*-


#  GUI 생성
root = Tk()
root.title("Send Mail Program")  # 타이틀 설정
root.geometry("640x600")  # 창 크기 (가로*세로)

# 제목 설정
subject_frame = LabelFrame(root, text="제목 설정")
subject_frame.pack(fill="x")

subject_txt = Entry(subject_frame)
subject_txt.pack(side="left", fill="x", expand=True)

# 본문 내용
content_frame = LabelFrame(root, text="본문 내용 설정")
content_frame.pack(fill="both", padx=5, pady=5)

scrollbar2 = Scrollbar(content_frame)
scrollbar2.pack(side="right", fill="y")

txt = Text(content_frame, width=30, height=7, yscrollcommand=scrollbar2.set)
txt.pack(fill="both")
scrollbar2.config(command=txt.yview)


# 첨부파일 추가
file_frame = LabelFrame(root, text="첨부파일 추가")
file_frame.pack(fill="x", padx=5, pady=5, ipady=5)

scrollbar3 = Scrollbar(file_frame)
scrollbar3.pack(side="right", fill="y")

file_list = Listbox(file_frame, selectmode="extended",
                    height=2, yscrollcommand=scrollbar3.set)
file_list.pack(fill="both", expand=True)
scrollbar3.config(command=file_list.yview)


# 파일 추가 버튼 프레임
def add_file():
    files = filedialog.askopenfilenames(
        title="이미지 파일을 선택하세요", filetypes=(("PNG 파일", "*.png"), ("모든 파일", "*.*")), initialdir=r"C:/")  # initialdir에 r은 탈출문자 없이 그대로 쓰겠다는 의미

    for file in files:
        file_list.insert(END, file)


def del_file():
    for index in reversed(file_list.curselection()):
        file_list.delete(index)


addfile_frame = Frame(root)
addfile_frame.pack(fill="x", padx=5, pady=5)

btn_del_user = Button(addfile_frame, text="파일 삭제",
                      command=del_file, padx=3, pady=3, width=12)
btn_del_user.pack(side="right")

btn_add_user = Button(addfile_frame, text="파일 추가",
                      command=add_file, padx=3, pady=3, width=12)
btn_add_user.pack(side="right")


# 수신자 목록 프레임
list_frame = LabelFrame(root, text="수신자 이메일 목록")
list_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")

email_list = Listbox(list_frame, selectmode="extended",
                     height=5, yscrollcommand=scrollbar.set)
email_list.pack(side="left", fill="both", expand=True)
scrollbar.config(command=email_list.yview)


def add_user():
    files = filedialog.askopenfilename(
        title="텍스트 파일을 선택하세요", filetypes=(("TXT 파일", "*.txt"), ("모든 파일", "*.*")), initialdir=r"C:/")  # initialdir에 r은 탈출문자 없이 그대로 쓰겠다는 의미

    with open(files, 'r') as file:
        for line in file:
            i = 0
            i += 1
            url = line.strip('\n')
            email_list.insert(END, url)


def del_user():
    for index in reversed(email_list.curselection()):
        email_list.delete(index)


# 수신자 버튼 프레임
user_frame = Frame(root)
user_frame.pack(fill="x", padx=5, pady=5)

btn_del_user = Button(user_frame, text="수신자 삭제",
                      command=del_user, padx=3, pady=3, width=12)
btn_del_user.pack(side="right")

btn_add_user = Button(user_frame, text="수신자 추가",
                      command=add_user, padx=3, pady=3, width=12)
btn_add_user.pack(side="right")

# 진행 상황 Progress Bar
# frame_progress = LabelFrame(root, text="진행상황")
# frame_progress.pack(fill="x", padx=5, pady=5, ipady=5)

# p_var = DoubleVar()
# progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
# progress_bar.pack(fill="x", padx=5, pady=5)

# 이메일 입력 프레임
id_frame = LabelFrame(root, text="아이디,패스워드 입력")
id_frame.pack(fill="x", padx=5, pady=5)
user_id, password = StringVar(), StringVar()
ttk.Entry(id_frame, textvariable=user_id).pack(fill="x")
ttk.Entry(id_frame, textvariable=password, show="*").pack(fill="x")


# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=3, pady=3, text="닫기",
                   width=12, command=root.quit)
btn_close.pack(side="right")

# 메일 전송 함수(아아디)


def send_mail():
    msg = MIMEMultipart()

    # 수신자 가져오기
    email = email_list.get(0, END)  # 튜플형식으로 변환
    list(email)  # 리스트 형식으로 변환
    email_str = ",".join(email)  # 리스트 배열을 str 문자열로 변환

    # 이메일 가져오기
    email_id = user_id.get()
    email_password = password.get()

    msg['From'] = email_id
    msg['To'] = email_str
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = Header(s=subject_txt.get(), charset='utf-8')
    body = MIMEText(txt.get("1.0", END))
    msg.attach(body)

    # 첨부파일 추가
    files_list = file_list.get(0, END)  # 튜플형식으로 변환
    list(files_list)  # 리스트박스 형식의 파일을 리스트로 형변환

    for f in files_list:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(f, "rb").read())
        encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="%s"' % os.path.basename(f))
        msg.attach(part)

    mailServer = smtplib.SMTP_SSL('smtp.gmail.com')
    # 본인의 계정과 비밀번호를 사용
    # mailServer.login('아이디입력', '비밀번호입력')
    mailServer.login(email_id, email_password)
    mailServer.send_message(msg)
    mailServer.quit()
    msgbox.showinfo("전송 완료", "메일 전송이 완료되었습니다")


btn_start = Button(frame_run, padx=3, pady=3, text="전송",
                   width=12, command=send_mail)
btn_start.pack(side="right")


root.resizable(False, False)
root.mainloop()  # 창 실행
