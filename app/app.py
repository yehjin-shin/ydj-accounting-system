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

import pandas as pd
import re

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class Receipt:
    def __init__(self, filepath):
        self.raw = pd.read_csv(filepath)
        self.main_path = 'app/'

        SCOPES = ['https://www.googleapis.com/auth/drive']

        creds = None
        if os.path.exists(f'{self.main_path}/token.pickle'):
            with open(f'{self.main_path}/token.pickle', 'rb') as token:
                creds = pickle.load(token)
                
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(f'{self.main_path}/credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.drive_service = build('drive', 'v3', credentials=creds)

    def make_use_list(self):
        df = self.raw.copy()

        df['영수증'] = self.get_id(df['영수증 업로드'])
        df['개수'] = [len(v) for v in df['영수증']]
        df = df.drop('영수증 업로드', axis=1)
        self.real_receipt = sum(df['개수'][df['실물영수증 여부']=='실물영수증 O'])

        df = df[['프로그램명', '인원', '계정사유', '적요', '결제일자', '금액', '가맹점명','개수']]
        self.use_list_df = pd.DataFrame(columns=['프로그램명', '인원', '계정사유', '적요', '결제일자', '금액', '가맹점명','개수'])

        index = 0
        for idx, val in df.iterrows():
            num = val['개수']
            for i in range(num):
                self.use_list_df.loc[index] = val
                index +=1
                
    def save_use_list(self, filepath):
        self.use_list_df.to_excel(f'{filepath}/예산사용내역서.xlsx')
        
    def make_receipt_df(self):
        df = self.raw.copy()

        # 필요한 column만 남기기
        df = df[['사용 RA','인원','결제일자','프로그램명','금액', '업체선정사유', '품명/수량/금액','영수증 업로드']]

        # id만 list로 뽑아내는 과정
        df['영수증'] = self.get_id(df['영수증 업로드'])
        df = df.drop('영수증 업로드', axis=1)

        # 원하는 꼴로 변형
        df['인원'] = [str(v)+" 명" for v in df['인원']]
        df['금액'] = [str(self.make_comma(v))+" 원" for v in df['금액']]
        
        self.receipt_df = df

    def save_receipt(self, row, num, filepath):
        loc_short = [(1640, 205), (2600, 205), (690,270), (1600, 270),(2555, 270), (550,525)]
        loc_long = [(1640, 205), (2600, 205), (690,270), (1550, 270),(2555, 270), (550,525)]
        loc_house = (690, 205)
        loc_image = (300,670)
        col_item = [2260, 2690, 2782]
        row_item = [685, 830, 964, 1107, 1265, 1402, 1540, 1677, 1830, 1962]

        # 폰트 경로와 사이즈 설정
        regularFont =ImageFont.truetype(f'{self.main_path}/fonts/HANBatang.ttf',36)
        smallFont =ImageFont.truetype(f'{self.main_path}/fonts/HANBatang.ttf',30)

        target_image = Image.open(f"{self.main_path}/form/form.jpeg")
        draw = ImageDraw.Draw(target_image)

        # 하우스명 기입
        draw.text(loc_house, "윤동주 하우스", fill="black", font=regularFont, align='center')

        # 사용RA, 인원, 결제일자, 프로그램명, 금액, 업체선정사유 기입
        if (len(row[3]) > 9):
            for i in range(6):
                draw.text(loc_long[i], str(row[i]), fill="black", font=regularFont, align='center')
        else:
            for i in range(6):
                draw.text(loc_short[i], str(row[i]), fill="black", font=regularFont, align='center')

        # 품명 기입
        if type(row[6]) == type('string'):
            item_list = self.divide_item(row[6])
            for i in range(len(item_list)):
                # 품명
                if len(item_list[i][0]) > 13: # 품명 긴 경우
                    front = item_list[i][0][:13]
                    back = item_list[i][0][13:]
                    draw.text((col_item[0], row_item[i]-20), front, fill="black", font=smallFont, align='center') 
                    draw.text((col_item[0], row_item[i]+20), back, fill="black", font=smallFont, align='center') 
                else: # 품명 짧은 경우
                    draw.text((col_item[0], row_item[i]), item_list[i][0], fill="black", font=smallFont, align='center') 
                draw.text((col_item[1], row_item[i]), item_list[i][1], fill="black", font=smallFont, align='center') # 수량
                draw.text((col_item[2], row_item[i]), item_list[i][2], fill="black", font=smallFont, align='center') # 금액

        # 영수증 사진
        check = 0
        for i in range(len(row[7])):
            check += 1
            fh = io.BytesIO()
            file_id = row[7][i]
            request = self.drive_service.files().get_media(fileId=file_id)
            downloader = googleapiclient.http.MediaIoBaseDownload(fh, request)

            done = False
            while done is False:
                status, done = downloader.next_chunk()

            add_image = Image.open(fh)
            if int(add_image.size[0]*(1300/add_image.size[1])) > 1300:
                target_image.paste(im = add_image.resize((1300, int(add_image.size[1]*(1300/add_image.size[0])))), box=loc_image)              
            else:
                target_image.paste(im = add_image.resize((int(add_image.size[0]*(1300/add_image.size[1])), 1300)), box=loc_image)
            
            target_image.save(f"{filepath}/{num}({check}) {row[3]}.png")
            
    # 품명/수량/금액 나누는 메소드
    def divide_item(self, full):
        item = full.split('\n')
        item_list = [i.split("/") for i in item]
        for i in range(len(item_list)):
            item_list[i][2] = str(self.make_comma(item_list[i][2]))+" 원"
        return item_list

    # 사진 링크에서 id 뽑아내는 메소드
    def get_id(self, url_list):
        url_list = [url+"," for url in url_list]
        pattern = re.compile('{}(.*?){}'.format(re.escape('id='), re.escape(',')))
        return [pattern.findall(url) for url in url_list] 

    # 숫자를 000,000,000 꼴로 바꾸는 정규식
    def make_comma(self, num):
        return re.sub('(?<=\d)(?=(\d{3})+(?!\d))',',',str(num))


class Program:
    def __init__(self):
        win = tk.Tk() # 창 생성
        win.geometry("600x500") # 창 크기
        win.title("Receipt Image Maker") # 제목
        win.option_add("*Font", "맑은고딕 12") # 전체 폰트

        btn_width = 40
        btn_height = 2

        self.input_btn = tk.Button(win)
        self.input_btn.config(width=btn_width, height=btn_height, text='파일(csv)을 업로드해주세요', command=self.input_df)
        self.input_btn.pack()

        self.path_btn = tk.Button(win)
        self.path_btn.config(width=btn_width, height=btn_height, text='다운 받을 폴더를 선택해주세요', command=self.input_path)
        self.path_btn.pack()

        self.start_btn = tk.Button(win)
        self.start_btn.config(width=btn_width, height=btn_height, text='시작', command=self.make_receipt)
        self.start_btn.pack()

        self.progress = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(win, maximum=1, length=200, variable=self.progress)
        self.progress_bar.pack()

        quit_btn = tk.Button(win)
        quit_btn.config(width=btn_width, height=btn_height, text='종료', command=quit)
        quit_btn.pack()

        self.info = tk.Tk() # 창 생성
        self.info.geometry("300x100") # 창 크기
        self.info.title("Information") # 제목
        self.info.option_add("*Font", "맑은고딕 12") # 전체 폰트

        win.mainloop()

    def input_df(self):
        filepath = filedialog.askopenfilename(title="csv 파일을 선택하세요.", filetypes=([("csv files", "*.csv")]))   
        self.receipt = Receipt(filepath)
        self.receipt.make_use_list()
        self.receipt.make_receipt_df()

        show_path = '/'.join(filepath.split('/')[-3:])
        self.input_btn.config(text=f'업로드 완료\n{show_path}')

        whole_num = len(self.receipt.use_list_df)
        real_num = self.receipt.real_receipt
        info_text = f"전체 영수증 개수: {whole_num} 개\n실물 영수증 개수: {real_num} 개"

        info_label = tk.Label(self.info, text=info_text, width=20, height=5)
        info_label.pack()

        self.info.mainloop()

    def input_path(self):
        self.folder = filedialog.askdirectory(title="저장할 위치를 선택하세요.")
        show_path = '/'.join(self.folder.split('/')[-3:])
        self.path_btn.config(text=f'폴더 선택 완료\n{self.folder}')

    def make_receipt(self):
        self.receipt.save_use_list(self.folder)
        num = len(self.receipt.receipt_df)
        for i in range(num):
            row = self.receipt.receipt_df.iloc[i]
            self.receipt.save_receipt(row, i, self.folder)
            self.progress.set((i+1)/num)
            self.progress_bar.update()
        self.start_btn.config(text='제작이 완료되었습니다')

    def num_receipt_info(self):
        whole_num = len(self.receipt.use_list_df)
        real_num = self.receipt.real_receipt
        info_text = f"전체 영수증 개수: {whole_num} 개\n실물 영수증 개수: {real_num} 개"

        info_label = tk.Label(self.info, text=info_text, width=25, height=5)
        info_label.pack()

        info.mainloop()

Program()
