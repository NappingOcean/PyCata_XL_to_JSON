
#import zone

import tkinter as tk
from tkinter import ttk
import xl_to_json as x2j

#stupid constant. Well, it can be changed.
WIDTH:int = 240
HEIGHT:int = 360

#Window 설정
app = tk.Tk()
app.title('XLon Py')
app.geometry( str(WIDTH) + 'x'+ str(HEIGHT) )

#컴포넌트
label1 = tk.Label(app, text = "파일 로드: ")
label1.grid(row=0, column = 0, sticky="w", padx=10, pady= 10)

label2 = tk.Label(app, text = "로드된 파일: ")
label2.grid(row=1, column = 0, sticky="w", padx=10, pady= 10)

frame1 = tk.LabelFrame(app, text="|파일 정보|")
frame1.grid(row = 2, column= 0, sticky="w", padx=10, pady= 10)

label4 = tk.Label(frame1, text= "여기에 파일 정보가 로드됩니다.")
label4.grid(row=1, column=0)

#실행!
app.mainloop()
