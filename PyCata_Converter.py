
#import zone

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import xl_to_json as x2j
import os

#stupid constant. Well, it can be changed.
WIDTH:int = 340
HEIGHT:int = 480

# callback 함수 설정
def click1():
    label2.configure(text = "로드된 파일: "+ f"{tk_filename.get()} 파일 로드됨.")
    btn_JSONize.configure(state="normal")
    
def checker(*args):
    if len(tk_filename.get()) != 0 and tk_filename.get() != placeholder_text:
        btn.configure(state="normal")
    else:
        btn["state"]="disabled"
        
def clickJSON():
    x2j.Xl_to_Json(tk_filename.get(), xl_path0, jo_path0).conv() # 모듈에 저장된 함수 실행
    label4.configure(text=f"{tk_filename.get()} 파일이 성공적으로 JSON으로 변환되었습니다!\n 경로: {jo_path0_look} 이니, 찾아주시기 바랍니다.")

#Window 설정
win = tk.Tk()
win.title('XLon Py') # XL + JSON + Py
win.geometry( str(WIDTH) + 'x' + str(HEIGHT) )

# 기타 변수

tk_filename = tk.StringVar()
placeholder_text = "파일 선택..."

xl_path0 = "./Excel/"
jo_path0 = "./JSON/"
jo_path0_look = jo_path0

if jo_path0.startswith("./"):
    jo_path0_look = jo_path0[2:]
    jo_path0_look = "(현재 디렉토리)/"+jo_path0_look

Xl_list = os.listdir(xl_path0)


## tk 구현 영역
# 컴포넌트

frame0 = tk.LabelFrame(win)
frame0.grid(row=0,column=0,sticky="ew",padx=10,pady=10)

label1 = tk.Label(frame0, text = "파일 로드: ")
label1.grid(row=0, column = 0)

combox = ttk.Combobox(frame0, textvariable=tk_filename, state="readonly")
combox.grid(row=0, column=1)
combox["values"] = Xl_list
combox.bind("<<ComboboxSelected>>", checker)

btn = ttk.Button(frame0, text="가져오기", command=click1, state="disabled")
btn.grid(row=0, column=2)

label2 = tk.Label(win, text = "로드된 파일: ")
label2.grid(row=1, column = 0, sticky="w", padx=10, pady= 10)

# TODO JSON 미리보기 계획중. 나중에!
# frame1 = tk.LabelFrame(win, text="|파일 정보|")
# frame1.grid(row = 2, column= 0, sticky="ew", padx=10, pady= 10)

# label3 = tk.Label(frame1, text= "여기에 파일 정보가 로드됩니다.")
# label3.grid(row=0, column=0)

# json_databox = ScrolledText(frame1, width= 42, height = 7)
# json_databox.grid(row=1, column=0)
# json_databox.configure(state='disabled')

# JSON 만드는 버튼
btn_JSONize = ttk.Button(win, text="Extract to JSON!", state="disabled", command=clickJSON)
btn_JSONize.grid(row=3, column=0)


label4 = tk.Label(win, text ="xxx")
label4.grid(row=4,column=0, sticky="EW", padx=10, pady=10)



#실행!
win.mainloop()
