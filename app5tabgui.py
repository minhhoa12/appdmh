from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, Listbox, Scrollbar
import os
import shutil
import xlwings as xw
import win32com.client
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
xl = win32com.client.Dispatch("Excel.Application")
xlPasteValues = -4163
import pandas as pd
import glob
import yagmail
import re
import win32api
import time
import subprocess
import win32print
from unidecode import unidecode

class MyApp:
	def __init__(self, root):
		self.root = root
		self.root.title("dmh APP") #Tiêu đề cửa sổ
		#self.root.resizable(False, False)
        #Khởi tạo biến
		self.copy_file = tk.StringVar()
		self.list_dir = tk.StringVar()
		self.path_bbg = tk.StringVar()
		self.path_po_tam = tk.StringVar()
		self.path_destination = tk.StringVar()
		self.new_po_name = tk.StringVar()
		self.folder_path = tk.StringVar()
		self.directory = tk.StringVar()
		#Biến tab5
		self.folder_path1 = tk.StringVar()
		self.folder_pathcpct = tk.StringVar()
		self.folder_pathcppo = tk.StringVar()
		
		self.status = tk.StringVar() 
  
		self.tab_control = ttk.Notebook(root)
#Tạo tab 1
		self.tab1 = ttk.Frame(self.tab_control)
		self.tab_control.add(self.tab1, text='MakePO')

		self.create_make_po()
        
#Tạo Tab2
        
		self.tab2 = ttk.Frame(self.tab_control)
		self.tab_control.add(self.tab2, text='Send Email')

		self.create_send_email_tab()
#Tạo Tab3  
		self.found_files = []  # Danh sách các tệp đã tìm thấy
		self.root_directory = tk.StringVar()
		self.destination_directory = tk.StringVar()
		
		self.tab3 = ttk.Frame(self.tab_control) #Tạo tab3
		self.tab_control.add(self.tab3, text='Find_Copy')
		self.setup_ui()

#Tạo tab 4       
#        self.tab4 = ttk.Frame(self.tab_control)
#        self.tab_control.add(self.tab4, text="Tab 4")
#        self.create_tab4_content()
		self.tab4 = ttk.Frame(self.tab_control)
		self.tab_control.add(self.tab4, text="Listdirfile")
		self.create_tab4_content()
# Tạo tab 5
		self.tab5 = ttk.Frame(self.tab_control)
		self.tab_control.add(self.tab5, text="Inpdf")
		self.create_tab5_content()

# Tạo tab 6
		self.tab6 = ttk.Frame(self.tab_control)
		self.tab_control.add(self.tab6, text="Rename")
		self.create_tab6_content()
		
		#self.tab_control.pack(expand=1, fill='both')
		self.tab_control.pack(expand=False, fill='both')
#Bắt đầu Tab6
	def create_tab6_content(self):
		self.label1 = tk.Label(self.tab6, text="Chọn thư mục để đổi tên:")
		self.label1.pack()

		self.folder_path_entry = tk.Entry(self.tab6, textvariable=self.directory, width=40)
		self.folder_path_entry.insert(0, r'E:\LUU TAM')
		self.folder_path_entry.pack()
		
		self.browse_path = tk.Button(self.tab6, text="Browse", command=self.browse_path)
		self.browse_path.pack()
		
		self.rename_files_and_directories = tk.Button(self.tab6, text="Đổi tên", command=self.rename_files_and_directories)
		self.rename_files_and_directories.pack()

		self.status_label = tk.Label(self.tab6, textvariable=self.status)
		self.status_label.pack()

	def browse_path(self):
		folder = filedialog.askdirectory()
		self.directory.set(folder)
		
	def remove_diacritics(self, text):
		return unidecode(text)

	def has_diacritics(self, text):
		return any(c != self.remove_diacritics(c) for c in text)

	def rename_files_and_directories(self):
		directory = self.folder_path_entry.get()
		self.rename_files_and_directories_recursive(directory)

	def rename_files_and_directories_recursive(self, directory):
		for item in os.listdir(directory):
			full_item_path = os.path.join(directory, item)
			if os.path.isdir(full_item_path):
				self.rename_files_and_directories_recursive(full_item_path)
				if self.has_diacritics(item):
					new_item_name = self.remove_diacritics(item)
					new_full_item_path = os.path.join(directory, new_item_name)
					os.rename(full_item_path, new_full_item_path)
					print(f'Renamed directory: {item} -> {new_item_name}')
			elif os.path.isfile(full_item_path):
				if self.has_diacritics(item):
					new_item_name = self.remove_diacritics(item)
					new_full_item_path = os.path.join(directory, new_item_name)
					os.rename(full_item_path, new_full_item_path)
					print(f'Renamed file: {item} -> {new_item_name}')

#Bắt đầu Tab5
	def create_tab5_content(self):
		self.label1 = tk.Label(self.tab5, text="Chọn thư mục để in:")
		self.label1.pack()

		self.folder_path_entry = tk.Entry(self.tab5, textvariable=self.folder_path1, width=40)
		self.folder_path_entry.insert(0, r'E:\LUU TAM')
		self.folder_path_entry.pack()

		self.po_tam_button = tk.Button(self.tab5, text="In PDF", command=self.check_and_print)
		self.po_tam_button.pack()

		self.labelcp = tk.Label(self.tab5, text="Cập nhật danh sách file:")
		self.labelcp.pack()
		self.labelcpct = tk.Label(self.tab5, text="Đường dẫn file contact")
		self.labelcpct.pack()
		self.folder_path_cpct = tk.Entry(self.tab5, textvariable=self.folder_pathcpct, width=40)
		self.folder_path_cpct.insert(0, r'Contact.xlsx')
		self.folder_path_cpct.pack()
		self.labelcppo = tk.Label(self.tab5, text="Đường dẫn file PO")
		self.labelcppo.pack()
		self.folder_path_cppo = tk.Entry(self.tab5, textvariable=self.folder_pathcppo, width=40)
		self.folder_path_cppo.insert(0, r'TheodoiData.xlsx')
		self.folder_path_cppo.pack()
		self.po_tam_buttoncp = tk.Button(self.tab5, text="Cập nhật", command=self.update_contact)
		self.po_tam_buttoncp.pack()

		self.status_label = tk.Label(self.tab5, textvariable=self.status)
		self.status_label.pack()

	def update_contact(self):
	    # Get the file paths from the Entry widgets
	    file_path_cpct = self.folder_pathcpct.get()
	    file_path_cppo = self.folder_pathcppo.get()

	    # Check if the file paths are valid (you may want to add more validation)
	    if not os.path.exists(file_path_cpct) or not os.path.exists(file_path_cppo):
	        self.status.set("File paths are invalid. Please check the paths.")
	        return

	    with xw.App(visible=False) as app:
	        # Update the contact file
	        with xw.Book(file_path_cpct) as book_cpct:
	            book_cpct.api.RefreshAll()
	            book_cpct.save()

	        # Update the PO file
	        with xw.Book(file_path_cppo) as book_cppo:
	            book_cppo.api.RefreshAll()
	            book_cppo.save()

	    self.status.set("Cập nhật hoàn thành")



	def normalize_path(self, path):
		#pass
		return path.replace('\\', '/')

	def print_pdfs(self, folder_path1):
		# Lay danh sach file pdf
		pdf_files = glob.glob(os.path.join(folder_path1, '*.pdf'))
		#pdf_files = [os.path.abspath(self.normalize_path(file)) for file in pdf_files]

		pdf_files = [self.normalize_path(file) for file in pdf_files]
		#print("PDF File Path:", pdf_files)  # Print the file path before using ShellExecute
		# Vong lap in tung file
		for pdf_file in pdf_files:
			try:
				print("Trying to print:", pdf_file)
				#subprocess.run(["rundll32.exe", "shimgvw.dll,ImageView_PrintTo", pdf_file])
				win32api.ShellExecute(0, "print", pdf_file, None, ".", 0)  # None /n:1 in trang 1
				#os.startfile(pdf_file, "print")
				self.status.set(f"Dang in file: {pdf_file}")
				time.sleep(7)
				# Xoa file da in
				os.remove(pdf_file)
				self.status.set(f"Da xoa file: {pdf_file}")
				#time.sleep(5)
			except Exception as e:
				self.status.set(f"Loi khi in {pdf_file}: {e}")

	def check_and_print(self):
		folder_path1 = self.folder_path1.get()  # Lấy đường dẫn từ biến folder_path
		# Kiem tra ton tai thu muc
		if not os.path.exists(folder_path1):
			self.status.set(f"Thư mục '{folder_path1}' không tồn tại.")
			return

		# Lay danh sach file pdf
		pdf_files = glob.glob(os.path.join(folder_path1, '*.pdf'))
		pdf_files = [self.normalize_path(file) for file in pdf_files]

		if not pdf_files:
			self.status.set("Không tìm thấy file pdf để in")
			return
		else:
			self.status.set("Đã tìm thấy, tiến hành in file pdf")
			self.print_pdfs(folder_path1)
#Kết thúc tab5			

#Bắt đầu các module của MakePO---------------
	def create_make_po(self):
		self.label1 = tk.Label(self.tab1, text="Chọn form BBGH File:")
		self.label1.pack()

		self.bbg_entry = tk.Entry(self.tab1, textvariable=self.path_bbg, width=40)
		self.bbg_entry.insert(0,r'BBGH1.xlsm')
		self.bbg_entry.pack()

		self.bbg_button = tk.Button(self.tab1, text="Browse", command=self.browse_bbg)
		self.bbg_button.pack()

		self.label2 = tk.Label(self.tab1, text="Select thư mục chứa PO_TAM:")
		self.label2.pack()

		self.po_tam_entry = tk.Entry(self.tab1, textvariable=self.path_po_tam, width=40)
		self.po_tam_entry.insert(0,r'E:\LUU TAM')
		self.po_tam_entry.pack()

		self.po_tam_button = tk.Button(self.tab1, text="Browse", command=self.browse_po_tam)
		self.po_tam_button.pack()
		
		self.po_tam_update_button = tk.Button(self.tab1, text="Cập nhật lại PO", command=self.update_po_tam_list)
		self.po_tam_update_button.pack()
		
		self.label3 = tk.Label(self.tab1, text="Chọn thư mục lưu PO:")
		self.label3.pack()

		self.destination_entry = tk.Entry(self.tab1, textvariable=self.path_destination, width=40)
		self.destination_entry.insert(0,r'Z:\13. DIEU PHOI HANG\5. BAO TIEU\1 Danh Son\1 DNP-DSG\1 Don hang\2023\Thang 10')
		self.destination_entry.pack()

		self.destination_button = tk.Button(self.tab1, text="Browse", command=self.browse_destination)
		self.destination_button.pack()
		
		self.label4 = tk.Label(self.tab1, text="Đặt tên PO:")
		self.label4.pack()

		self.new_po_entry = tk.Entry(self.tab1, textvariable=self.new_po_name, width=40)
		self.new_po_entry.pack()

		self.copy_button = tk.Button(self.tab1, text="Thực hiện", command=self.copy_files)
		self.copy_button.pack()

		self.listbox_label = tk.Label(self.tab1, text="Danh sách PO trong thư mục:")
		self.listbox_label.pack()
			 
		self.listbox = Listbox(self.tab1, selectmode=tk.SINGLE, width=50)
		self.listbox.pack()

		self.scrollbar = Scrollbar(self.tab1, orient="vertical")
		self.scrollbar.config(command=self.listbox.yview)
		self.scrollbar.pack(side="right", fill="y")

		self.listbox.config(yscrollcommand=self.scrollbar.set)

		self.status_label = tk.Label(self.tab1, textvariable=self.status)
		self.status_label.pack()

	def browse_bbg(self):
		self.path_bbg.set(filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm")]))
		self.status.set('Đã chọn xong form BBGH')
	def browse_po_tam(self):
		folder = filedialog.askdirectory()
		self.path_po_tam.set(folder)

		excel_files = [f for f in os.listdir(folder) if f.lower().endswith(".xlsx")]
		self.listbox.delete(0, tk.END)  # Clear the listbox
		for file in excel_files:
			self.listbox.insert(tk.END, file)
		self.status.set('Liệt kê các PO xong')    
	def update_po_tam_list(self):
		po_tam_folder = self.path_po_tam.get()
		if po_tam_folder:
			excel_files = [f for f in os.listdir(po_tam_folder) if f.lower().endswith(".xlsx")]
			self.listbox.delete(0, tk.END)  # Clear the listbox
			for file in excel_files:
				self.listbox.insert(tk.END, file)
				self.status.set('Đã cập nhật danh mục')
		else:
			#print("Please select PO_TAM folder first.")
			self.status.set('Chưa chọn thư mục chứa PO')

	def browse_destination(self):
		self.path_destination.set(filedialog.askdirectory())
		self.status.set('Đã chọn xong thư mục lưu PO')
	
	def copy_files(self):
		bbg_path = self.path_bbg.get()
		po_tam_folder = self.path_po_tam.get()
		destination_path = self.path_destination.get()
		new_po_name = self.new_po_name.get()

		selected_file_index = self.listbox.curselection()
		if selected_file_index:
			selected_file = self.listbox.get(selected_file_index)
			po_tam_path = os.path.join(po_tam_folder, selected_file)

			if bbg_path and po_tam_path and destination_path and new_po_name:
				app = xw.App(visible=False)
				wb_bbg = xw.Book(bbg_path)
				wb_po_tam = xw.Book(po_tam_path)

				try:
					sh_bbg = wb_bbg.sheets["BBGH"]
					sh_po_tam = wb_po_tam.sheets["Sheet1"]
					# sh_contact_df = pd.read_excel(wb_bbg.fullname, sheet_name="contact", engine="openpyxl")
					sh_contact_df = pd.read_excel('contact.xlsx', sheet_name="Contact", engine="openpyxl")
					Noinhan = self.new_po_name.get().split(" ")[-1]
					matching_row = sh_contact_df[sh_contact_df['CODE'] == Noinhan]

					if not matching_row.empty:
						B12_result = matching_row.iloc[0]['DON_VI_NHAN_HANG']
						B13_result = matching_row.iloc[0]['DIA_CHI_NHAN_HANG']
						B15_result = matching_row.iloc[0]['NGUOI_LIEN_HE']
					else:
						B12_result = 'LOI DU LIEU'
						B13_result = 'LOI DU LIEU'
						B15_result = 'LOI DU LIEU'

					last_row = sh_po_tam.used_range.last_cell.row
					list=[]
					for i in range(1, last_row+1):
						cell=sh_po_tam.range('I'+str(i)).value
						if cell is not None:
							list=list+[i]
					d1=min(list)
					d2=max(list)
					data_range=sh_po_tam.range('C'+str(d1)+':I'+str(d2)).api.Copy()
					# data_range=sh_po_tam.range(f'C{d1}:I{d2}')
					# df_vungchon = pd.DataFrame(data_range.value)

					sh_bbg.range('A6').value = new_po_name
					sh_bbg.range('B12').value = B12_result
					sh_bbg.range('B13').value = B13_result
					sh_bbg.range('B15').value = B15_result
					sh_bbg.range('B19').api.PasteSpecial(Paste=xlPasteValues)
					# sh_bbg.range('B19').value = df_vungchon.values

					new_po_path = os.path.join(os.path.dirname(bbg_path), f"{new_po_name}.xlsm")
					wb_bbg.save(new_po_path)
					wb_bbg.close()
					wb_po_tam.close()
					app.quit()
					wb_bbg=None
					wb_po_tam=None
					app=None
					try:
						wb = excel.Workbooks("Theo doi dieu phoi hang di 2023.xlsm")
						print("The file is currently open.")
						ws = wb.Worksheets("DATA")
						# last_row = sh_po_tam.used_range.last_cell.row
						last_row = ws.Range("C" + str(ws.Rows.Count)).End(-4162).Row
						ws.Range("C" + str(last_row + 1)).Value = new_po_name
						#ws.Save()
						print('DA THEM PO',new_po_name,'VAO FILE THEO DOI PO')

					except Exception as e:
						if "Cannot access" in str(e):
							print("The file is not open.")
						else:
							print("An unknown error occurred:", e)

					shutil.move(new_po_path, destination_path)
					self.status.set('Di chuyển PO thành công')
					os.remove(po_tam_path)
				except Exception as e:
					self.status.set(f'Có lỗi xảy ra: {str(e)}')
			else:
				self.status.set('Có trường chưa nhập thông tin')
		else:
			self.status.set('Chọn file PO để thực hiện')
#Kết thúc các module của MakePO---------------

#Bắt đầu module sendmail
	def create_send_email_tab(self):

		label_po = tk.Label(self.tab2, text="Enter PO Number:")
		label_po.pack()

		self.entry_po = tk.Entry(self.tab2)
		self.entry_po.pack()

		label_folder = tk.Label(self.tab2, text="Choose Folder Path:")
		label_folder.pack()

		self.entry_folder = tk.Entry(self.tab2, textvariable=self.folder_path)
		self.entry_folder.insert(0,r'Z:\13. DIEU PHOI HANG\5. BAO TIEU\1 Danh Son\1 DNP-DSG\1 Don hang\2023\Thang 10')
		self.entry_folder.pack()

		browse_button = tk.Button(self.tab2, text="Browse", command=self.browse_folder)
		browse_button.pack()

		label_recipient = tk.Label(self.tab2, text="Choose Recipient Group:")
		label_recipient.pack()

		self.recipient_var = tk.StringVar(value="DAP_MN")
		recipient_options = ["DAP_MN", "DAP_MB", "DAP_MT", "Custom"]
		recipient_menu = tk.OptionMenu(self.tab2, self.recipient_var, *recipient_options)
		recipient_menu.pack()

		label_custom_recipients = tk.Label(self.tab2, text="Custom Recipients (comma(,)-separated):")
		label_custom_recipients.pack()

		self.entry_custom_recipients = tk.Entry(self.tab2)
		self.entry_custom_recipients.pack()

		send_button = tk.Button(self.tab2, text="Send Email", command=self.send_email)
		send_button.pack()
		
		#self.status_label = tk.Label(self.tab2, text="",width=40)
		self.status_label = tk.Label(self.tab2, textvariable=self.status,width=40)
		self.status_label.pack()

	def browse_folder(self):
		folder_selected = filedialog.askdirectory()
		if folder_selected:
			self.folder_path.set(folder_selected)
		#self.status.set('Chọn thư mục chứa PO để gửi:')
	def send_email(self):
		ten_po = self.entry_po.get()
		folder_path = self.folder_path.get()

		pdf_files = glob.glob(folder_path + "\\" + ten_po + "\\*.pdf")
		jpg_files = glob.glob(folder_path + "\\" + ten_po + "\\*.jpg")
		danhsach_file = pdf_files + jpg_files

		nguoinhan = self.recipient_var.get()

		if nguoinhan == "Custom":
			custom_recipients = self.entry_custom_recipients.get()
			to = custom_recipients.split(',')
		else:
			recipient_dict = {
				"DAP_MN": ['dungln@daplogistics.vn','duyenntn@daplogistics.vn','sonnh@daplogistics.vn','nghiann@daplogistics.vn','thaclb@daplogistics.vn','nhottb@daplogistics.vn','bichthoa89@gmail.com','tra.vy.danapha@gmail.com','nam.vo@danapha.com'],
				"DAP_MB": ['nhungnt@daplogistics.vn','huyennt2@daplogistics.vn','tra.vy.danapha@gmail.com','nam.vo@danapha.com','thuylb@ds-group.vn'],
				"DAP_MT": ['minhhoavosadn@gmail.com']
			}
			to = recipient_dict.get(nguoinhan, [])

		email = 'hoa.danapha@gmail.com'
		password = 'kmtjoloqsxwznfnb'
		yag = yagmail.SMTP(email, password)

		subject = f'PKN theo PO so {ten_po}'
		body = '''Dear All,
		Hoa gui PKN theo file dinh kem
		
		Tran trong.
		Duong Minh Hoa
		Phone: 0905686872
		'''
		attachments = danhsach_file

		if not attachments:
			#print(f'No attachments found. Email not sent.')
			status_text = "No attachments found. Email not sent."
		else:
			yag.send(to=to, subject=subject, contents=body, attachments=attachments)
			yag.close()
			print(f'Email sent to: {to}\n{attachments}')
			status_text = f'Email sent to: {", ".join(to)}\n{attachments}'
			self.status.set("Đã gửi email")
			#self.status_label.config(text=status_text)
#Kết thúc module sendmail

#Bắt đầu find and copy
#browse tab 3
	def browse_root_directory(self):
		self.root_directory.set(filedialog.askdirectory())
		self.status.set('Đã chọn xong đường dẫn tìm kiếm')
	def browse_destination_directory(self):
		self.destination_directory.set(filedialog.askdirectory())
		self.status.set('Đã chọn xong đường dẫn lưu trữ')
	def setup_ui(self):

		self.root_directory_label = tk.Label(self.tab3, text="Chọn đường dẫn tìm kiếm:")
		self.root_directory_entry = tk.Entry(self.tab3, textvariable=self.root_directory)
		self.root_directory_entry.insert(0,r'E:\LUU TAM\PKN')
		self.root_directory_browse_button = tk.Button(self.tab3, text="Browser..", command=self.browse_root_directory)

		self.destination_directory_label = tk.Label(self.tab3, text="Chọn đường dẫn lưu")
		self.destination_directory_entry = tk.Entry(self.tab3, textvariable=self.destination_directory)
		self.destination_directory_entry.insert(0,r'E:\LUU TAM')
		self.destination_directory_browse_button = tk.Button(self.tab3, text="Browser..", command=self.browse_destination_directory)

		self.search_label = tk.Label(self.tab3, text="Nhập text cần tìm")
		self.search_entry = tk.Entry(self.tab3)

		self.search_button = tk.Button(self.tab3, text="Tìm. . .", command=self.search_files)

		self.additional_search_label = tk.Label(self.tab3, text="Additional Search Term")
		self.additional_search_entry = tk.Entry(self.tab3)
		self.additional_search_button = tk.Button(self.tab3, text="Find Additional", command=self.additional_search)

		self.results_listbox = tk.Listbox(self.tab3, selectmode=tk.SINGLE, width=40)

		self.copy_button = tk.Button(self.tab3, text="Copy. . .", command=self.copy_selected_file)
		self.copy_status_label = tk.Label(self.tab3, text="")
		
		self.root_directory_label.pack()
		self.root_directory_entry.pack()
		self.root_directory_browse_button.pack()

		self.destination_directory_label.pack()
		self.destination_directory_entry.pack()
		self.destination_directory_browse_button.pack()

		self.search_label.pack()
		self.search_entry.pack()
		self.search_button.pack()

		self.additional_search_label.pack()
		self.additional_search_entry.pack()
		self.additional_search_button.pack()

		self.results_listbox.pack()

		self.copy_button.pack()
		self.copy_status_label.pack()

	def browse_root_directory(self):
		self.root_directory.set(filedialog.askdirectory())

	def browse_destination_directory(self):
		self.destination_directory.set(filedialog.askdirectory())

	def copy_selected_file(self):
		selected_file = self.results_listbox.get(tk.ACTIVE)
		if selected_file and selected_file != "Khong tim thay ket qua":
			selected_file_index = self.results_listbox.curselection()[0]  # Lay chi so tep da chon
			
			#source_file_name = self.found_files[selected_file_index]
			source_file_path = os.path.abspath(os.path.join(self.root_directory.get(), self.found_files[selected_file_index]))  # Tạo đường dẫn hoàn chỉnh
			#source_file = os.path.join(new_root_directory, self.found_files[selected_file_index])  # Lay duong dan tu danh sach
			destination_file = os.path.join(self.destination_directory.get(), os.path.basename(source_file_path))
			
			#source_file_path = source_file_path.replace("\\", "/")
			#destination_file = destination_file.replace("\\", "/")
			#print("Source file:", source_file_path)
			#print("Destination file:", destination_file)
			
			if os.path.exists(source_file_path):
				try:
					shutil.copy(source_file_path, destination_file)
					self.copy_status_label.config(text=f"File copied: {os.path.basename(source_file_path)} DONE!!!")
				except Exception as e:
					self.copy_status_label.config(text=f"Error copying file: {e}")
			else:
				self.copy_status_label.config(text=f"Source file not found")
				
	def search_files(self):
		search_term = self.search_entry.get()
		
		self.found_files = []  # Đặt lại danh sách
		self.results_listbox.delete(0, tk.END)  # Xóa kết quả cũ

		if search_term:
			self.found_files = self.perform_search(search_term)
			self.display_results()

	def additional_search(self):
		additional_search_term = self.additional_search_entry.get()
	
		self.found_files = []  # Đặt lại danh sách
		self.results_listbox.delete(0, tk.END)  # Xóa kết quả cũ
			
		if additional_search_term:
			additional_found_files = self.perform_search(additional_search_term)
			self.found_files.extend(additional_found_files)
			self.display_results()

	def perform_search(self, search_term):
		result_files = []
		san_pham = search_term[:-6]
		so_lo = search_term[-6:]
		search_pattern = re.compile(fr"{san_pham}.*{so_lo}", re.IGNORECASE)
		for root, dirs, files in os.walk(self.root_directory.get()):
			for file in files:
				if search_pattern.search(file):
					file_path = os.path.abspath(os.path.join(root, file))
					result_files.append(file_path)
		return result_files

	def display_results(self):
		for file_path in self.found_files:
			self.results_listbox.insert(tk.END, file_path)

		if not self.found_files:
			self.results_listbox.insert(tk.END, "Khong tim thay ket qua")
#Kết thúc
#browse tab4
	def browse_list_dir(self):
		self.list_dir.set(filedialog.askdirectory())
		self.status.set('Đã chọn xong đường dẫn tìm kiếm')
	def browse_copy_file(self):
		self.copy_file.set(filedialog.askdirectory())
		self.status.set('Đã chọn xong đường dẫn lưu trữ')
		
	def create_tab4_content(self):
	
		self.list_dir_label = tk.Label(self.tab4, text="Chọn đường dẫn tìm kiếm:")
		self.list_dir_label.pack()
		self.list_dir_entry = tk.Entry(self.tab4, textvariable=self.list_dir)
		self.list_dir_entry.pack()
		self.list_dir_entry.insert(0,r'X:\6- PKN- THANH PHAM')
		#current_value_list_dir = self.list_dir.get()
		self.list_dir_browse_button = tk.Button(self.tab4, text="Browser..", command=self.browse_list_dir)
		self.list_dir_browse_button.pack()

		self.copy_file_label = tk.Label(self.tab4, text="Chọn đường dẫn lưu")
		self.copy_file_label.pack()
		self.copy_file_entry = tk.Entry(self.tab4, textvariable=self.copy_file)
		self.copy_file_entry.pack()
		self.copy_file_entry.insert(0,r'E:\LUU TAM\PKN')
		#current_value_copy_file = self.copy_file.get()
		self.copy_file_browse_button = tk.Button(self.tab4, text="Browser..", command=self.browse_copy_file)
		self.copy_file_browse_button.pack()
		
		self.copy_button = tk.Button(self.tab4, text="Thực hiện", command=self.listdirfile)
		self.copy_button.pack()
		
		self.status_label = tk.Label(self.tab4, textvariable=self.status)
		self.status_label.pack()
		
	
	def listdirfile(self):
		# Load configuration from config.txt
		#config_file = "config_lisdirfile.txt"
		#config = {}
		#with open(config_file, "r", encoding="utf-8") as cfg:
		#	for line in cfg:
		#		line = line.strip()
		#		if "=" in line:
		#			key, value = line.split("=", 1)
		#			config[key.strip()] = value.strip()

		#list_dir = config.get("list_dir", "")
		#copy_file = config.get("copy_file", "")
		
		list_dir = self.list_dir.get()
		copy_file = self.copy_file.get()
		# Tao file luu tru
		list_file = "list_file.txt"

		# Tao tap hop de luu tru cac ten file da tai xuong
		downloaded_files = set()

		# Nap tap hop tu file luu tru (neu co)
		if os.path.exists(list_file):
			with open(list_file, "r", encoding="utf-8") as file:
				downloaded_files = set(file.read().splitlines())

		# Liet ke cac file trong thu muc list_dir va cap nhat vao list_file.txt neu chua ton tai
		updated_files = set()  # Tap hop cac file da cap nhat
		with open(list_file, "a", encoding="utf-8") as f:
			for root, _, files in os.walk(list_dir):
				for file in files:
					file_path = os.path.join(root, file)
					if file_path not in downloaded_files:
						f.write(file_path + '\n')
						downloaded_files.add(file_path)
						updated_files.add(file_path)

		if updated_files:
			self.status.set(f"Da cap nhat {len(updated_files)} file")
			print(f"Da cap nhat {len(updated_files)} duong dan moi vao list_file.txt.")
		else:
			self.status.set("Khong co duong dan moi de cap nhat.")
			print("Khong co duong dan moi de cap nhat.")

		# Copy va doi ten cac file vao thu muc
		if not os.path.exists(copy_file):
			os.makedirs(copy_file)

		for file in updated_files:
			filename = os.path.basename(file)
			parent_folder_name = os.path.basename(os.path.dirname(file))  # Get the parent folder name
			new_filename = f"{parent_folder_name} {filename}"
			dest_file_path = os.path.join(copy_file, new_filename)
			shutil.copy(file, dest_file_path)
			#self.status.set(f"Da copy va doi ten file {file} thanh {dest_file_path}")
			print(f"Da copy va doi ten file {file} thanh {dest_file_path}")
#Kết thúc tab4
    
if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()
