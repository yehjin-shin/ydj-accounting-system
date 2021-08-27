from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import io
import googleapiclient.http

from PIL import Image,ImageDraw,ImageFont
import os

import numpy as np
import pandas as pd
import re

import img2pdf

import tkinter as tk
from tkinter import ttk
from datetime import datetime
from tkinter import filedialog
import time

from Receipt import Receipt

def get_folder():
    folder = filedialog.askdirectory(title="저장할 위치를 선택하세요.")
    return folder

def get_file():
    file = filedialog.askopenfilename(title="파일을 선택하세요.", filetypes=(("csv files", "*.csv"), ("xlsx files", "*.xlsx")))   
    btn2.config(command=swap)
    return file

def input_raw_df():
	filepath = get_file()
	original = Receipt(filepath)



win = tk.Tk() # 창 생성
win.geometry("700x300") # 창 크기
win.title("exercise") # 제목
win.option_add("*Font", "맑은고딕 15") # 전체 폰트

introduce = "1. 아래의 파일 업로드 버튼을 클릭하여 구글폼을 통해 다운 받은 파일 (csv, excel)을 업로드해주세요. \n2. 파일 업로드가 완료되었다면, 영수증 파일과 예산사용내역서를 다운 받을 폴더를 선택해주세요. \n   이때, 폴더 안에 이전에 프로그램을 다운 받았던 파일이 있으면 자동으로 덮어씌워지니 폴더를 잘 선택해주세요.\n4. quit 버튼을 눌러 종료해주세요.\n \n"

label_text = tk.StringVar()
label_text.set(introduce)

label = ttk.Label(win, textvariable=label_text)
label.pack()

progress = tk.DoubleVar()
progress_bar = ttk.Progressbar(win, maximum=100, length=200, variable=progress)
progress_bar.pack()


def start():
	for i in range(100):
		time.sleep(0.1)
		progress.set(i+1)
		progress_bar.update()

btn2 = tk.Button(win) # 버튼 생성
btn2.config(width=10, height=1) # 버튼 크기
btn2.config(text='upload file (csv, excel)')
filepath = btn2.config(command=get_file)
btn2.pack()


btn = tk.Button(win) # 버튼 생성
btn.config(width=10, height=1) # 버튼 크기
btn.config(text='quit')
btn.config(command=quit)
btn.pack()

win.mainloop() # 창 And 실행








