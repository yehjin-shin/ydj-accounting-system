# https://hogni.tistory.com/102
# https://www.youtube.com/watch?v=Es1fQqqxIFQ&t=285s
# https://ponyozzang.tistory.com/437
# https://zzsza.github.io/development/2020/07/05/python-class/
# https://ponyozzang.tistory.com/574
# https://ucwoogong.com/189

class Program:
	def __init__():
		win = tk.Tk() # 창 생성
		win.geometry("700x300") # 창 크기
		win.title("exercise") # 제목
		win.option_add("*Font", "맑은고딕 15") # 전체 폰트

		introduce = "1. 아래의 파일 업로드 버튼을 클릭하여 구글폼을 통해 다운 받은 파일 (csv, excel)을 업로드해주세요. \n2. 파일 업로드가 완료되었다면, 영수증 파일과 예산사용내역서를 다운 받을 폴더를 선택해주세요. \n   이때, 폴더 안에 이전에 프로그램을 다운 받았던 파일이 있으면 자동으로 덮어씌워지니 폴더를 잘 선택해주세요.\n4. quit 버튼을 눌러 종료해주세요.\n \n"
		label_text = tk.StringVar()
		label_text.set(introduce)
		label = ttk.Label(win, textvariable=label_text)
		label.pack()


		input_btn = tk.Button(win)
		input_btn.config(width=10, height=1, text='upload file (csv, excel)', command=self.input_df)
		input_btn.pack()

		path_btn = tk.Button(win)
		path_btn.config(width=10, height=1, text='Select a folder', command=self.input_path)
		path_btn.pack()

		start_btn = tk.Button(win)
		start_btn.config(width=10, height=1, text='Start', command=self.make_receipt)
		start_btn.pack()

		progress = tk.DoubleVar()
		progress_bar = ttk.Progressbar(win, maximum=1, length=200, variable=progress)
		progress_bar.pack()

		quit_btn = tk.Button(win)
		quit_btn.config(width=10, height=1, text='Quit', command=quit)
		quit_btn.pack()

		win.mainloop() # 창 And 실행

	def input_df():
		filepath = filedialog.askopenfilename(title="파일을 선택하세요.", filetypes=(("csv files", "*.csv"), ("xlsx files", "*.xlsx")))   
		self.receipt = Receipt(filepath)
		self.receipt.make_use_list()
		self.receipt.make_receipt_df()
		input_btn.config(text='upload finished')

	def input_path():
		self.folder = filedialog.askdirectory(title="저장할 위치를 선택하세요.")
		path_btn.config(text='folder selected')

	def make_receipt():
		self.receipt.save_use_list(self.folder)
		num = len(self.receipt.raw)
		for i in range(num):
			row = self.receipt.raw.iloc[i]
			self.receipt.save_receipt(row, i+1, folder)
			progress.set((i+1)/num)



