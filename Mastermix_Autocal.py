import xlwings as xw 
import pandas as pd 
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter.ttk import *
import easygui
from easygui import *
import os

'''
Tạo cửa sổ tkinter, chú ý phải có main.loop nếu không cửa sổ sẽ không hiện.
Dùng lệnh zoom để mở fullscreen màn hình
Lệnh geometry để khai báo kích thước khi thu nhỏ màn hình
'''
window_main = Tk()
window_main.title('Mastermix Auto Cal_230222')    # thay đổi version để biết version đang sử dụng
window_main.state('zoomed')      # tạo màn hình full screen nhưng vẫn còn thanh task bar
window_main.geometry('800x600')      # kích thước màn hình khi minimize

'''
Tạo scrollbar và các frame để chứa dữ liệu. Tạo canvas để dùng cho scrollbar
'''
'Tạo main_frame'
frame_main = Frame(window_main)
frame_main.pack(fill=BOTH, expand=True)
'Tạo canvas nằm trong main frame'
canvas_main = Canvas(frame_main)
canvas_main.pack(side=LEFT, fill=BOTH,expand=True)
'Tạo scrollbar nằm trong canvas'
scrollbar_1 = Scrollbar(frame_main,orient=VERTICAL,command=canvas_main.yview)
scrollbar_1.pack(side=RIGHT,fill=Y)
'Điều chỉnh canvas'
canvas_main.configure(yscrollcommand=scrollbar_1.set)
canvas_main.bind('<Configure>',lambda e: canvas_main.configure(scrollregion=canvas_main.bbox('all')))
'Tạo sub frame nằm trong canvas để để đưa các frame phụ vào'
frame_sub = Frame(canvas_main)
frame_sub.pack()
'Tạo cửa sổ cho sub frame'
canvas_main.create_window((0,0),window=frame_sub,anchor='nw')


'''
'mở file excel và chọn vị trí các sheet tương ứng'
'''
app = xw.App(visible=False, add_book=False)                     # dùng lệnh visible =false để không nhìn thấy file excel được mở lên
wb1 = app.books.open('MasterMixAutoCal.xlsx')
sh_data = wb1.sheets('DATA')                                    # sheet DATA sẽ lưu trữ các thông tin về các thành phần phản ứng
sh_chi_tieu_1 = wb1.sheets('CHI_TIEU_1')                        # sheet chỉ tiêu 1, chỉ tiêu 2 sẽ hiện thông tin lên treeview.
sh_chi_tieu_2 = wb1.sheets('CHI_TIEU_2')
sh_chi_tieu_3 = wb1.sheets('CHI_TIEU_3')
sh_print = wb1.sheets('PRINT')                                  # sheet PRINT sẽ lưu thông tin để xuất ra file pdf
list_chi_tieu_total = sh_data.range('B7:ZZ7').value             # tạo list các chỉ tiêu trong phạm vi từ ô B7 đến ô ZZ7
list_chi_tieu_sort = [i for i in list_chi_tieu_total if i]      # loại bỏ các phần tử rỗng trong list chỉ tiêu
so_phan_ung = [i for i in range (0,101)]                        # tạo các list số lượng phản ứng đê nhập hoặc chọn số phản ứng

sh_chi_tieu_1.visible = True
sh_chi_tieu_2.visible = True
sh_chi_tieu_3.visible = True
sh_print.visible = True

'tạo sheet để copy data'
wb_data_copy = app.books.open('du_lieu_pha_mix.xlsx')
sh_data_copy = wb_data_copy.sheets('du_lieu_pha_mix')


'''
Tạo các frame để đưa dữ liệu vào.
- Tạo frame 1 cho chỉ tiêu 1. tương tự cho các chỉ tiêu 2 và chỉ tiêu 3
mỗi 1 frame là 1 hàng khác nhau.
frame 1 là 1 label
'''
frame_1=Label(frame_sub)
frame_1.pack(fill=X)

'tạo frame 1a nằm trong frame 1'
frame_1a = Frame(frame_1)
frame_1a.pack(fill=X,padx=5,pady=5)

'tạo frame dọc trong frame ngang, và add label chỉ tiêu, số lượng mẫu vào'
frame_1c = Frame(frame_1a)                      # frame 1c nằm trong frame 1a, fill theo chiều dọc
frame_1c.pack(side=LEFT,fill=Y,padx=5,pady=5)

'tạo frame 1d trong 1c, và fill ngang, chứa tên chỉ tiêu và combobox chọn chỉ tiêu'
frame_1d = Frame(frame_1c)
frame_1d.pack(side=TOP,fill=X,padx=5,pady=5)
label_chi_tieu_1 = Label(frame_1d, text='Chỉ tiêu 1: ',width=10)
label_chi_tieu_1.pack(side=LEFT,padx=5,pady=5)
combobox_chi_tieu_1 = Combobox(frame_1d,width=35)
combobox_chi_tieu_1.pack(side=LEFT,padx=5,pady=5)
combobox_chi_tieu_1['values']=list_chi_tieu_sort

'TẠO HÀM SEARCH ĐỂ TÌM NHANH TÊN CÁC CHỈ TIÊU'
def check_input_1 (event):
    value = event.widget.get()
    if value == '':
        combobox_chi_tieu_1['values']=list_chi_tieu_sort
    else:
        data =[]
        for chi_tieu in list_chi_tieu_sort:
            if value.lower() in chi_tieu.lower():
                data.append(chi_tieu)
        combobox_chi_tieu_1['values']=data
combobox_chi_tieu_1.bind('<KeyRelease>',check_input_1)

'tạo frame 1e trong frame 1c, để chứa số mẫu'
frame_1e = Frame(frame_1c)
frame_1e.pack(side=TOP,fill=X,padx=5,pady=5)
label_so_mau = Label(frame_1e,text='Số mẫu: ',width=10)
label_so_mau.pack(side=LEFT,padx=5,pady=5)
combobox_so_mau = Combobox(frame_1e,width=15)
combobox_so_mau.pack(side=LEFT,padx=5,pady=5)
combobox_so_mau['values']=so_phan_ung

'tạo label và entry box cho code mẫu'
label_code_mau = Label(frame_1a,text='Code mẫu phân tích: ',width=20)
label_code_mau.pack(side=LEFT,padx=5,pady=5)

'tạo text box để nhập code mẫu phân tích'
text_code_mau = Text(frame_1a,height=2,width=40,wrap=WORD)
text_code_mau.pack(side=LEFT,padx=5,pady=5)

'tạo label và combobox cho số phản ứng'
label_so_phan_ung_1 = Label(frame_1a,text='Số phản ứng:',width=12)
label_so_phan_ung_1.pack(side=LEFT,padx=5,pady=5)
combobox_so_phan_ung_1 = Combobox(frame_1a,width=12)
combobox_so_phan_ung_1.pack(side=LEFT,padx=5,pady=5)
combobox_so_phan_ung_1['values']=so_phan_ung
Btn_OK_1 = Button(frame_1a, text='OK',width=25)
Btn_OK_1.pack(side=LEFT,padx=5,pady=5)
Btn_clear_1 = Button(frame_1a,text='Xóa chỉ tiêu 1', width=25)
Btn_clear_1.pack(side=LEFT,padx=5,pady=5)

'tạo frame 1b chứa ghi chú và thể tich 1 phản ứng'
frame_1b = Frame(frame_1)
label_ghi_chu = Label(frame_1b,text='',wraplength=500)
label_ghi_chu.pack(side=LEFT,padx=5,pady=1)

'Tạo label nằm trong frame 1 để đưa treeview vào'
Label_1 = Label(frame_1b, text='')
Label_1.pack(side=LEFT,padx=5,pady=5)


'''
TẠO TREEVIEW VÀ ĐƯA TREEVIEW VÀO GUI tkinter
'''
tree_chi_tieu_1 = ttk.Treeview(Label_1)

'tạo style và đánh dấu tag cho treeview, phục vụ cho việc tô màu từng dòng'
style = ttk.Style()

'change selected color'
style.configure('Treeview',background='#D3D3D3',
                foreground='black',
                fieldbackground='#D3D3D3')
style.map('Treeview',background=[('selected','blue')])

'tạo tag để đánh dấu dòng chẵn dòng lẻ'
tree_chi_tieu_1.tag_configure('oddrow',background='white')
tree_chi_tieu_1.tag_configure('evenrow',background='lightblue')

'tạo hàmg click_1 để chạy lệnh khi nhấn nút button 1'
def click_1():
    frame_1b.pack(fill=X,padx=5,pady=5)         # add frame 1 vào bảng
    sh_chi_tieu_1.clear_contents()
    sh_data.range('D1').value= combobox_so_phan_ung_1.get()                # Lấy số lượng phản ứng
    sh_data.range('E1').value = int(combobox_so_mau.get()) + 1
    sh_data.range('B1').value = combobox_chi_tieu_1.get()
    cell_begin = sh_data.range('B2').value
    cell_end = sh_data.range('B3').value
    cell_volume = sh_data.range('B4').value
    cell_note = sh_data.range('B5').value
    range_pu_1 = sh_data.range(f'{cell_begin}:{cell_end}').value
    sh_chi_tieu_1.range('A1').value = range_pu_1
    wb1.save()
    
    'Đưa nội dung thể tích phản ứng và ghi chú vào label ghi chú'
    str_volume_1 = sh_data.range(f'{cell_volume}').value
    str_note_1 = sh_data.range(f'{cell_note}').value
    if str_note_1 != None:
        label_ghi_chu.configure(text=str_volume_1+'\n'+str_note_1,borderwidth=2,relief='solid') # các đặc tính của boder (relief) bao gồm: "flat", "raised", "sunken", "ridge", "solid", and "groove". groove và ridge đòi boder with tối thiểu =2
    else:
        label_ghi_chu.configure(text=str_volume_1,borderwidth=2,relief='solid')
    
    'dùng pandas để đọc dữ liệu'
    df_1 = pd.read_excel('MasterMixAutoCal.xlsx','CHI_TIEU_1')
    
    'clear tree'
    clear_tree()
    
    'set up new treeview'
    tree_chi_tieu_1['column']=list(df_1.columns)
    tree_chi_tieu_1.column(0,width=35)          # điều chỉnh chiều rộng cho cột 0, chính là cột đầu tiên, cột Số thứ tự
    tree_chi_tieu_1['show']='headings'
    
    'tạo vòng lặp, loop through column list for headers. để tạo tiêu đều cho các cột'
    for column in tree_chi_tieu_1['column']:
        tree_chi_tieu_1.heading(column,text=column)
    
    'put data in treeview'
    df_1_row = df_1.to_numpy().tolist()

    'tạo dòng max và cột max'
    row_max = len(df_1_row)
    column_max = len(tree_chi_tieu_1['column'])
    tree_chi_tieu_1.configure(height=row_max)

    'tạo các dòng chẵn dòng lẻ dựa trên cột Số thứ tự để tô màu'
    for row in df_1_row:
        if row[0] %2 ==0:
            tree_chi_tieu_1.insert('','end',values=row,tags='evenrow')
        else:
            tree_chi_tieu_1.insert('','end',values=row,tags='oddrow')
    
    'pack the treeview finally'
    tree_chi_tieu_1.pack()
    
    'pack frame_2, để xuất hiện frame 2'
    frame_2a.pack(fill=X,padx=5,pady=5)
    frame_2b.pack(fill=X,padx=5,pady=5)

'tạo hàm để clear tree khi nhấn nút lại'
def clear_tree():
    tree_chi_tieu_1.delete(*tree_chi_tieu_1.get_children())

'configure nút bấm button để chạy lệnh'
Btn_OK_1.configure(command=click_1)

'''
TẠO FRAME 2 CHO CHỈ TIÊU 2
'''
frame_2=Label(frame_sub) 
frame_2.pack(fill=X,padx=5,pady=5)         

'tạo frame 2a nằm trong frame 2'
frame_2a = Frame(frame_2)

'tạo label và combobox cho chỉ tiêu phân tích'
frame_2c = Frame(frame_2a)
frame_2c.pack(side=LEFT,fill=Y,padx=5,pady=5)

frame_2d = Frame(frame_2c)
frame_2d.pack(side=TOP,fill=X,padx=5,pady=5)
label_chi_tieu_2 = Label(frame_2d, text='Chỉ tiêu 2: ',width=10)
label_chi_tieu_2.pack(side=LEFT,padx=5,pady=5)
combobox_chi_tieu_2 = Combobox(frame_2d,width=35)
combobox_chi_tieu_2.pack(side=LEFT,padx=5,pady=5)
combobox_chi_tieu_2['values']=list_chi_tieu_sort

'Tạo hàm search chỉ tiêu 2'
def check_input_2(event):
    value = event.widget.get()
    if value =='':
        combobox_chi_tieu_2['values']=list_chi_tieu_sort
    else:
        data = []
        for chi_tieu in list_chi_tieu_sort:
            if value.lower() in chi_tieu.lower():
                data.append(chi_tieu)
        combobox_chi_tieu_2['values']=data 
combobox_chi_tieu_2.bind('<KeyRelease>',check_input_2)

frame_2e = Frame(frame_2c)
frame_2e.pack(side=LEFT,padx=5,pady=5)
label_so_mau_2 = Label(frame_2e,text='Số mẫu: ',width=10)
label_so_mau_2.pack(side=LEFT,padx=5,pady=5)
combobox_so_mau_2 = Combobox(frame_2e,width=15)
combobox_so_mau_2.pack(side=LEFT,padx=5,pady=5)
combobox_so_mau_2['values']=so_phan_ung

'tạo label và entry box cho code mẫu'
label_code_mau_2 = Label(frame_2a,text='Code mẫu phân tích: ',width=20)
label_code_mau_2.pack(side=LEFT,padx=5,pady=5)

'tạo text box để nhập code mẫu phân tích'
text_code_mau_2 = Text(frame_2a,height=2,width=40,wrap=WORD)
text_code_mau_2.pack(side=LEFT,padx=5,pady=5)

'tạo label và combobox cho số phản ứng'
label_so_phan_ung_2 = Label(frame_2a,text='Số phản ứng:',width=12)
label_so_phan_ung_2.pack(side=LEFT,padx=5,pady=5)
combobox_so_phan_ung_2 = Combobox(frame_2a,width=10)
combobox_so_phan_ung_2.pack(side=LEFT,padx=5,pady=5)
combobox_so_phan_ung_2['values']=so_phan_ung
Btn_OK_2 = Button(frame_2a, text='OK',width=25)
Btn_OK_2.pack(side=LEFT,padx=5,pady=5)
Btn_clear_2 = Button(frame_2a,text='Xóa chỉ tiêu 2',width=25)
Btn_clear_2.pack(side=LEFT,padx=5,pady=5)

'tạo frame 2b chứa ghi chú và thể tich 1 phản ứng'
frame_2b = Frame(frame_2)
label_ghi_chu_2 = Label(frame_2b,text='',wraplength=500)
label_ghi_chu_2.pack(side=LEFT,padx=5,pady=1)

'Tạo label nằm trong frame 1 để đưa treeview vào'
Label_2 = Label(frame_2b, text='')
Label_2.pack(side=LEFT,padx=5,pady=5)

'ẩn 2 frame 2a và 2b'
frame_2a.pack_forget()
frame_2b.pack_forget()

'tạo treeview 2'
tree_chi_tieu_2 = ttk.Treeview(Label_2)

'tạo tag để đánh dấu dòng chẵn dòng lẻ'
tree_chi_tieu_2.tag_configure('oddrow',background='white')
tree_chi_tieu_2.tag_configure('evenrow',background='lightblue')

'tạo hàmg click_1 để chạy lệnh khi nhấn nút button 1'
def click_2():
    frame_2b.pack(fill=X,padx=5,pady=5)
    sh_chi_tieu_2.clear_contents()
    sh_data.range('D1').value= combobox_so_phan_ung_2.get()                # Lấy số lượng phản ứng
    sh_data.range('E1').value = int(combobox_so_mau_2.get())+1
    sh_data.range('B1').value = combobox_chi_tieu_2.get()
    cell_begin_2 = sh_data.range('B2').value
    cell_end_2 = sh_data.range('B3').value
    cell_volume_2 = sh_data.range('B4').value
    cell_note_2 = sh_data.range('B5').value
    range_pu_2 = sh_data.range(f'{cell_begin_2}:{cell_end_2}').value
    sh_chi_tieu_2.range('A1').value = range_pu_2
    wb1.save()
    
    'Đưa nội dung thể tích phản ứng và ghi chú vào label ghi chú'
    str_volume_2 = sh_data.range(f'{cell_volume_2}').value
    str_note_2 = sh_data.range(f'{cell_note_2}').value
    if str_note_2 != None:
        label_ghi_chu_2.configure(text=str_volume_2+'\n'+str_note_2,borderwidth=2,relief='solid')
    else:
        label_ghi_chu_2.configure(text=str_volume_2,borderwidth=2,relief='solid')

    'dùng pandas để đọc dữ liệu'
    df_2 = pd.read_excel('MasterMixAutoCal.xlsx','CHI_TIEU_2')
    
    'clear tree'
    clear_tree_2()
    
    'set up new treeview'
    tree_chi_tieu_2['column']=list(df_2.columns)
    tree_chi_tieu_2.column(0,width=35)          
    tree_chi_tieu_2['show']='headings'
    
    'tạo vòng lặp, loop through column list for headers. để tạo tiêu đều cho các cột'
    for column in tree_chi_tieu_2['column']:
        tree_chi_tieu_2.heading(column,text=column)
    
    'put data in treeview'
    df_2_row = df_2.to_numpy().tolist()

    'tạo dòng max và cột max'
    row_max_2 = len(df_2_row)
    tree_chi_tieu_2.configure(height=row_max_2)

    'tạo các dòng chẵn dòng lẻ dựa trên cột Số thứ tự để tô màu'
    for row in df_2_row:
        if row[0] %2 ==0:
            tree_chi_tieu_2.insert('','end',values=row,tags='evenrow')
        else:
            tree_chi_tieu_2.insert('','end',values=row,tags='oddrow')
    
    'pack the treeview finally'
    tree_chi_tieu_2.pack()
    
    'pack frame_4, để xuất hiện frame 4'
    frame_4a.pack(fill=X,padx=5,pady=5)
    frame_4b.pack(fill=X,padx=5,pady=5)

'tạo hàm để clear tree khi nhấn nút lại'
def clear_tree_2():
    tree_chi_tieu_2.delete(*tree_chi_tieu_2.get_children())

'configure nút bấm button để chạy lệnh'
Btn_OK_2.configure(command=click_2)


'''
TẠO CHỈ TIÊU 3_ĐẶT TÊN FRAME LÀ FRAME 4, VÌ FRAME 3 LÀ FRAME CUỐI CÙNG ĐỂ ĐẶT CÁC NÚT CLEAR
'''
frame_4=Label(frame_sub) 
frame_4.pack(fill=X,padx=5,pady=5)         

'tạo frame 4a nằm trong frame 4'
frame_4a = Frame(frame_4)

'tạo label và combobox cho chỉ tiêu phân tích'
frame_4c = Frame(frame_4a)
frame_4c.pack(side=LEFT,fill=Y,padx=5,pady=5)

frame_4d = Frame(frame_4c)
frame_4d.pack(side=TOP,fill=X,padx=5,pady=5)
label_chi_tieu_3 = Label(frame_4d, text='Chỉ tiêu 3: ',width=10)
label_chi_tieu_3.pack(side=LEFT,padx=5,pady=5)
combobox_chi_tieu_3 = Combobox(frame_4d,width=35)
combobox_chi_tieu_3.pack(side=LEFT,padx=5,pady=5)
combobox_chi_tieu_3['values']=list_chi_tieu_sort

'Tạo hàm search cho chỉ tiêu 3'
def check_input_3(event):
    value = event.widget.get()
    if value =='':
        combobox_chi_tieu_3['values']=list_chi_tieu_sort
    else:
        data = []
        for chi_tieu in list_chi_tieu_sort:
            if value.lower() in chi_tieu.lower():
                data.append(chi_tieu)
        combobox_chi_tieu_3['values']=data
combobox_chi_tieu_3.bind('<KeyRelease>',check_input_3)

frame_4e = Frame(frame_4c)
frame_4e.pack(side=LEFT,fill=X,padx=5,pady=5)
label_so_mau_3 = Label(frame_4e,text='Số mẫu: ',width=10)
label_so_mau_3.pack(side=LEFT,padx=5,pady=5)
combobox_so_mau_3 = Combobox(frame_4e,width=15)
combobox_so_mau_3.pack(side=LEFT,padx=5,pady=5)
combobox_so_mau_3['values']=so_phan_ung

'tạo label và entry box cho code mẫu'
label_code_mau_3 = Label(frame_4a,text='Code mẫu phân tích: ',width=20)
label_code_mau_3.pack(side=LEFT,padx=5,pady=5)

'tạo text box để nhập code mẫu phân tích'
text_code_mau_3 = Text(frame_4a,height=2,width=40,wrap=WORD)
text_code_mau_3.pack(side=LEFT,padx=5,pady=5)

'tạo label và combobox cho số phản ứng'
label_so_phan_ung_3 = Label(frame_4a,text='Số phản ứng:',width=12)
label_so_phan_ung_3.pack(side=LEFT,padx=5,pady=5)
combobox_so_phan_ung_3 = Combobox(frame_4a,width=10)
combobox_so_phan_ung_3.pack(side=LEFT,padx=5,pady=5)
combobox_so_phan_ung_3['values']=so_phan_ung
Btn_OK_3 = Button(frame_4a, text='OK',width=25)
Btn_OK_3.pack(side=LEFT,padx=5,pady=5)
Btn_clear_3 = Button(frame_4a,text='Xóa chỉ tiêu 3',width=25)
Btn_clear_3.pack(side=LEFT,padx=5,pady=5)

'tạo frame 4b chứa ghi chú và thể tich 1 phản ứng'
frame_4b = Frame(frame_4)
label_ghi_chu_3 = Label(frame_4b,text='',wraplength=500)
label_ghi_chu_3.pack(side=LEFT,padx=5,pady=1)

'Tạo label nằm trong frame 4 để đưa treeview vào'
Label_3 = Label(frame_4b, text='')
Label_3.pack(side=LEFT,padx=5,pady=5)

'ẩn 2 frame 4a và 4b'
frame_4a.pack_forget()
frame_4b.pack_forget()

'tạo treeview 3'
tree_chi_tieu_3 = ttk.Treeview(Label_3)

'tạo tag để đánh dấu dòng chẵn dòng lẻ'
tree_chi_tieu_3.tag_configure('oddrow',background='white')
tree_chi_tieu_3.tag_configure('evenrow',background='lightblue')

'tạo hàmg click_3 để chạy lệnh khi nhấn nút button 1'
def click_3():
    frame_4b.pack(fill=X,padx=5,pady=5)
    sh_chi_tieu_3.clear_contents()
    sh_data.range('D1').value= combobox_so_phan_ung_3.get()                # Lấy số lượng phản ứng
    sh_data.range('E1').value = int(combobox_so_mau_3.get())+1
    sh_data.range('B1').value = combobox_chi_tieu_3.get()
    cell_begin_3 = sh_data.range('B2').value
    cell_end_3 = sh_data.range('B3').value
    cell_volume_3 = sh_data.range('B4').value
    cell_note_3 = sh_data.range('B5').value
    range_pu_3 = sh_data.range(f'{cell_begin_3}:{cell_end_3}').value
    sh_chi_tieu_3.range('A1').value = range_pu_3
    wb1.save()
    
    'Đưa nội dung thể tích phản ứng và ghi chú vào label ghi chú'
    str_volume_3 = sh_data.range(f'{cell_volume_3}').value
    str_note_3 = sh_data.range(f'{cell_note_3}').value
    if str_note_3 != None:
        label_ghi_chu_3.configure(text=str_volume_3+'\n'+str_note_3,borderwidth=2,relief='solid')
    else:
        label_ghi_chu_3.configure(text=str_volume_3,borderwidth=2,relief='solid')

    'dùng pandas để đọc dữ liệu'
    df_3 = pd.read_excel('MasterMixAutoCal.xlsx','CHI_TIEU_3')
    
    'clear tree'
    clear_tree_3()
    
    'set up new treeview'
    tree_chi_tieu_3['column']=list(df_3.columns)
    tree_chi_tieu_3.column(0,width=35)          
    tree_chi_tieu_3['show']='headings'
    
    'tạo vòng lặp, loop through column list for headers. để tạo tiêu đều cho các cột'
    for column in tree_chi_tieu_3['column']:
        tree_chi_tieu_3.heading(column,text=column)
    
    'put data in treeview'
    df_3_row = df_3.to_numpy().tolist()

    'tạo dòng max và cột max'
    row_max_3 = len(df_3_row)
    tree_chi_tieu_3.configure(height=row_max_3)

    'tạo các dòng chẵn dòng lẻ dựa trên cột Số thứ tự để tô màu'
    for row in df_3_row:
        if row[0] %2 ==0:
            tree_chi_tieu_3.insert('','end',values=row,tags='evenrow')
        else:
            tree_chi_tieu_3.insert('','end',values=row,tags='oddrow')
    
    'pack the treeview finally'
    tree_chi_tieu_3.pack()

'tạo hàm để clear tree khi nhấn nút lại'
def clear_tree_3():
    tree_chi_tieu_3.delete(*tree_chi_tieu_3.get_children())

'configure nút bấm button để chạy lệnh'
Btn_OK_3.configure(command=click_3)


'''
TẠO FRAME 3 ĐỂ ĐẶT CÁC NÚT CLEAR VÀ XUẤT PDF
TẠO HÀM CHO CÁC NÚT CLEAR VÀ DELETE CHỈ TIÊU.
'''
frame_3=Label(frame_sub)
frame_3.pack(fill=X,padx=5,pady=5)

label_ngay_nguoi_tinh_mix = Label(frame_3,text='Ngày / người tính mix: ')
label_ngay_nguoi_tinh_mix.pack(side=LEFT,padx=5,pady=5)

text_ngay_nguoi_tinh_mix = Entry(frame_3,width=40)      # Dùng entry box để lấy dữ liệu (get) không bị xuống dòng. nếu là text box sẽ bị xuống dòng
text_ngay_nguoi_tinh_mix.pack(side=LEFT,padx=5,pady=5)

'tạp các nút, open PCR (mở quy trình rút gọn), save (dùng để xuất pdf) và clear all (xóa hết chỉ tiêu)'
Btn_open_PCR = Button(frame_3, text='MỞ QUY TRÌNH RÚT GỌN',width=30)
Btn_open_PCR.pack(side=LEFT,padx=5,pady=5,expand=True)
Btn_save = Button(frame_3,text='XUẤT FILE PDF',width=30)
Btn_save.pack(side=LEFT,padx=5,pady=5,expand=True)
Btn_clear_all = Button(frame_3,text='Clear All',width=30)
Btn_clear_all.pack(side=LEFT,padx=5,pady=5,expand=True)

'''
Tạo các hàm cho nút clear.
'''
def delete_chi_tieu_1():
    combobox_chi_tieu_1.delete(0,'end')
    label_ghi_chu.configure(text='',borderwidth=2,relief='flat')        # boder dạng flat tức là sẽ không có boder xuất hiện
    combobox_so_phan_ung_1.delete(0,'end')
    combobox_so_mau.delete(0,'end')
    text_code_mau.delete('1.0','end')
    clear_tree()
    tree_chi_tieu_1.pack_forget()
    frame_1b.pack_forget()
    sh_chi_tieu_1.clear_contents()
    
def delete_chi_tieu_2():
    combobox_chi_tieu_2.delete(0,'end')
    label_ghi_chu_2.configure(text='',borderwidth=2,relief='flat')
    combobox_so_phan_ung_2.delete(0,'end')
    text_code_mau_2.delete('1.0','end')
    combobox_so_mau_2.delete(0,'end')
    clear_tree_2()
    tree_chi_tieu_2.pack_forget()
    frame_2b.pack_forget()
    sh_chi_tieu_2.clear_contents()
    # frame_2a.pack_forget()           # xóa toàn bộ frame 2 chứa chỉ tiêu 2 đang vá lỗi pack lại frame 2 sau khi xóa
    # frame_2.pack_forget()

def delete_chi_tieu_3():
    combobox_chi_tieu_3.delete(0,'end')
    label_ghi_chu_3.configure(text='',borderwidth=2,relief='flat')
    combobox_so_phan_ung_3.delete(0,'end')
    text_code_mau_3.delete('1.0','end')
    combobox_so_mau_3.delete(0,'end')
    clear_tree_3()
    tree_chi_tieu_3.pack_forget()
    frame_4b.pack_forget()          # chỉ tiêu 3 nhưng frame 4. Do add vào sau khi đã có frame 3 (để chứa nút clear, xuất pdf)
    sh_chi_tieu_3.clear_contents()
    # frame_4a.pack_forget()           # xóa frame 4 chứa chỉ tiêu 3, đang vá lỗi pack lại frame 4 sau khi xóa

Btn_clear_1.configure(command=delete_chi_tieu_1)
Btn_clear_2.configure(command=delete_chi_tieu_2)
Btn_clear_3.configure(command=delete_chi_tieu_3)
Btn_clear_all.configure(command=lambda: [delete_chi_tieu_1(),delete_chi_tieu_2(),\
                        delete_chi_tieu_3(),sh_print.clear_contents()])                 # dùng hàm lambda để gọi nhiều def khi click nút

'Tạo hàm cho nút open pcr'
def open_PCR():
    import subprocess                           # dùng subprocess để mở file thì có thể vừa mở file vừa thao tác trên giao diện GUI, có thể mở được nhiều file cùng 1 lúc
    subprocess.Popen(['Chu_trinh_PCR_rut_gon.docx'],shell=True)

Btn_open_PCR.configure(command=open_PCR)

'''
TẠO NÚT XUẤT PDF
'''
'tạo hàm để đưa dữ liệu sang sheet PRINT và xuất file pdf'

def xuat_pdf():
    wb1.save()
    sh_print.clear_contents()
    sh_print.range('1:100').unmerge()
    
    'tìm dòng cuối và cột cuối trong sheet chỉ tiêu 1 và paste sang sheet print'
    last_row_1 = sh_chi_tieu_1.range(f'A{sh_chi_tieu_1.cells.last_cell.row}').end('up').row   # giá trị trả ra là số thứ tự dòng cuối có giá trị
    last_column_1 = sh_chi_tieu_1.range('A1').end('right').last_cell.column                     # trả vê giá trị là số cột cuối cùng. ví dụ cột A là 1, cột D là 4
    'Dùng hàm chr() để chuyển số thành chữ theo code ASCII. ví dụ số 65 sẽ được chuyển thành chữ A'
    final_letter_1 = chr(65+last_column_1-1)        # trừ 1 để trả về đúng giá trị, ví dụ A là 65, D sẽ là 68, nhưng cột D là cột thứ 4. 
    
    if sh_chi_tieu_1.range('A1').value != None:
        'thêm chỉ tiêu phân tích, code mẫu phân tích và số phản ứng vào sheet print'
        sh_print.range('A1').value = 'Chỉ tiêu phân tích: '+ combobox_chi_tieu_1.get()+ '          ' + 'Số mẫu: ' + combobox_so_mau.get()
        sh_print.range('A2').value = 'Số phản ứng: '+ combobox_so_phan_ung_1.get() 
        sh_print.range('A3').value = 'Code mẫu phân tích: '+ str(text_code_mau.get('1.0',END))       # lấy giá trị ở dòng 1, ký tự 0, cho đến hết
        range_copy_1 = sh_chi_tieu_1.range(f'A1:{final_letter_1+str(last_row_1)}').value
        sh_print.range('A4').value = range_copy_1
        last_row_copy_1 = sh_print.range(f'A{sh_print.cells.last_cell.row}').end('up').row      # chọn ô cuôi cùng tại sheet print sau khi copy từ sheet chỉ tiêu 1  
        last_column_copy_1 = sh_print.range('A4').end('right').last_cell.column
        final_letter_copy_1 = chr(65+last_column_copy_1-1)      # ký tự cột cuối cùng trong bảng
        'format cell các ô trong vùng copy '
        sh_print.range('A1:D1').merge()     # merge các ô chỉ tiêu phân tích
        sh_print.range('A1').font.bold = True
        sh_print.range('A2:D2').merge()     # merge ô số phản ứng
        sh_print.range('A3:D3').merge()     # merge ô code mẫu phân tích
        range_code_mau_1 = sh_print.range('A3').value
        sh_print.range('A3').value = range_code_mau_1.strip()  # xóa khoảng trắng
        range_table_1 = sh_print.range(f'A4:{final_letter_copy_1+str(last_row_copy_1)}')        # chọn vùng để vẽ bảng
        table_1 = sh_print.tables.add(range_table_1,table_style_name='TableStyleLight18')       # vẽ bảng 1
        
    else:
        last_column_copy_1 = 1     # tạo điều kiện trong trường hợp ô rỗng thì cột nhỏ nhất là 1
        pass

    'tìm dòng cuối và cột cuối trong sheet chỉ tiêu 2 và paste sang dòng cuối trong sheet print +2 dòng'
    last_row_2 = sh_chi_tieu_2.range(f'A{sh_chi_tieu_2.cells.last_cell.row}').end('up').row  
    last_column_2 = sh_chi_tieu_2.range('A1').end('right').last_cell.column                  
    final_letter_2 = chr(65+last_column_2 - 1)
    
    if sh_print.range('A1').value != None:
        last_row_print = sh_print.range(f'A{sh_print.cells.last_cell.row}').end('up').row
    else:
        last_row_print = int(-1)
    
    if sh_chi_tieu_2.range('A1').value != None:
        sh_print.range(f'A{last_row_print+2}').value = 'Chỉ tiêu phân tích 2: ' + combobox_chi_tieu_2.get()+ '          ' + 'Số mẫu: ' + combobox_so_mau_2.get()     # last row +2 tương đương ô A1 với chỉ tiêu 1
        sh_print.range(f'A{last_row_print+3}').value = 'Số phản ứng: ' + combobox_so_phan_ung_2.get()
        sh_print.range(f'A{last_row_print+4}').value = 'Code mẫu phân tích: '+text_code_mau_2.get('1.0',END)    # last row +3 tương đương ô A2 với chỉ tiêu 1
        range_copy_2 = sh_chi_tieu_2.range(f'A1:{final_letter_2+str(last_row_2)}').value
        sh_print.range(f'A{last_row_print+5}').value = range_copy_2                                             # last row +4 tương đương ô A3 với chỉ tiêu 1
        last_row_copy_2 = sh_print.range(f'A{sh_print.cells.last_cell.row}').end('up').row      # chọn ô cuôi cùng sau khi copy chỉ tiêu 2 
        last_column_copy_2 = sh_print.range(f'A{last_row_print+5}').end('right').last_cell.column
        final_letter_copy_2 = chr(65+last_column_copy_2-1)      # ký tự cột cuối cùng trong bảng
        
        'format cell các ô trong vùng copy '
        sh_print.range(f'A{last_row_print+2}:D{last_row_print+2}').merge()     # merge các ô chỉ tiêu phân tích
        sh_print.range(f'A{last_row_print+2}').font.bold = True
        sh_print.range(f'A{last_row_print+3}:D{last_row_print+3}').merge()     # merge ô số phản ứng
        sh_print.range(f'A{last_row_print+4}:D{last_row_print+4}').merge()     # merge ô code mẫu phân tích
        range_code_mau_2 = sh_print.range(f'A{last_row_print+4}').value
        sh_print.range(f'A{last_row_print+4}').value = range_code_mau_2.strip()  # xóa khoảng trắng
        range_table_2 = sh_print.range(f'A{last_row_print+5}:{final_letter_copy_2+str(last_row_copy_2)}')        # chọn vùng để vẽ bảng
        table_2 = sh_print.tables.add(range_table_2,table_style_name='TableStyleLight18')       # vẽ bảng 2
        
    else:
        last_column_copy_2 = 1
        pass    
    
    'tìm dòng cuối và cột cuối trong sheet chỉ tiêu 3 và paste sang dòng cuối trong sheet print +2 dòng'
    last_row_3 = sh_chi_tieu_3.range(f'A{sh_chi_tieu_3.cells.last_cell.row}').end('up').row  
    last_column_3 = sh_chi_tieu_3.range('A1').end('right').last_cell.column                  
    final_letter_3 = chr(65+last_column_3 - 1)
    
    if sh_print.range('A1').value != None:
        last_row_print_3 = sh_print.range(f'A{sh_print.cells.last_cell.row}').end('up').row
    else:
        last_row_print_3 = int(-1)
    
    if sh_chi_tieu_3.range('A1').value != None:
        sh_print.range(f'A{last_row_print_3+2}').value = 'Chỉ tiêu phân tích 3: ' + combobox_chi_tieu_3.get()+ '          ' + 'Số mẫu: ' + combobox_so_mau_3.get()     # last row +2 tương đương ô A1 với chỉ tiêu 1
        sh_print.range(f'A{last_row_print_3+3}').value = 'Số phản ứng: ' + combobox_so_phan_ung_3.get()
        sh_print.range(f'A{last_row_print_3+4}').value = 'Code mẫu phân tích: '+text_code_mau_3.get('1.0',END)     # last row +3 tương đương ô A2 với chỉ tiêu 1
        range_copy_3 = sh_chi_tieu_3.range(f'A1:{final_letter_3+str(last_row_3)}').value
        sh_print.range(f'A{last_row_print_3+5}').value = range_copy_3                                             # last row +4 tương đương ô A3 với chỉ tiêu 1
        last_row_copy_3 = sh_print.range(f'A{sh_print.cells.last_cell.row}').end('up').row      # chọn ô cuôi cùng sau khi copy chỉ tiêu 2 
        last_column_copy_3 = sh_print.range(f'A{last_row_print_3+5}').end('right').last_cell.column
        final_letter_copy_3 = chr(65+last_column_copy_3-1)      # ký tự cột cuối cùng trong bảng
        
        'format cell các ô trong vùng copy '
        sh_print.range(f'A{last_row_print_3+2}:D{last_row_print_3+2}').merge()     # merge các ô chỉ tiêu phân tích
        sh_print.range(f'A{last_row_print_3+2}').font.bold = True
        sh_print.range(f'A{last_row_print_3+3}:D{last_row_print_3+3}').merge()     # merge ô số phản ứng
        sh_print.range(f'A{last_row_print_3+4}:D{last_row_print_3+4}').merge()     # merge ô code mẫu phân tích
        range_code_mau_3 = sh_print.range(f'A{last_row_print_3+4}').value
        sh_print.range(f'A{last_row_print_3+4}').value = range_code_mau_3.strip()  # xóa khoảng trắng
        range_table_3 = sh_print.range(f'A{last_row_print_3+5}:{final_letter_copy_3+str(last_row_copy_3)}')        # chọn vùng để vẽ bảng
        table_3 = sh_print.tables.add(range_table_3,table_style_name='TableStyleLight18')       # vẽ bảng 2
        sh_print.range('A:A').column_width = 6      # chỉnh chiều rộng cột A(là cột số thứ tự)
    else:
        last_row_copy_3 = sh_print.range(f'A{sh_print.cells.last_cell.row}').end('up').row   # để trong trường hợp chỉ tiêu 3 rỗng.
        last_column_copy_3 = 1
        pass
        
    'tìm số cột lớn nhất để tính ra chiều rộng của cột'
    column_max = max(last_column_copy_1,last_column_copy_2,last_column_copy_3)
    final_letter_copy_max = chr(65+column_max-1)     # tìm ký tự lớn nhất
    if column_max > 1:
        sh_print.range(f'B{last_row_print+5}:{final_letter_copy_max}4').column_width = 70/(column_max -1)   # chỉnh chiều rộng cột
    else:
        sh_print.range(f'B{last_row_print+5}:{final_letter_copy_max}4').column_width = 70/(4)
    
    sh_print.range('A:A').column_width = 6
    sh_print.range(f'A{last_row_print+2}:{final_letter_copy_max+str(last_row_copy_3)}').wrap_text = True   # chỉnh wrap text cho các cột
    sh_print.range(f'1:{last_row_copy_3}').rows.autofit()
    
    'thêm dòng ngày /  người tính master mix'
    last_row_final = sh_print.range(f'A{sh_print.cells.last_cell.row}').end('up').row  # tìm dòng cuối cùng
    sh_print.range(f'A{last_row_final+2}').value = "Ngày / Người tính master mix: " + str(text_ngay_nguoi_tinh_mix.get())
    sh_print.range(f'A{last_row_final+2}').wrap_text=False
    
    'CHUYỂN FILE THÀNH PDF'
    wb1.save()
    
    'copy sheet print sang file excel khac'
    range_copy = sh_print.range(f'1:{last_row_final+2}').value
    last_row_data_copy = sh_data_copy.range(f'A{sh_data_copy.cells.last_cell.row}').end('up').row
    sh_data_copy.range(f'A{last_row_data_copy+2}').value = range_copy    
    wb_data_copy.save()
    
    'tạo đường dãn lưu file pdf'    
    save_pdf = filesavebox('Chọn đường dẫn lưu file PDF','Mastermix Autocal',default="*.pdf")
    
    if save_pdf != None:
        sh_print.api.ExportAsFixedFormat(0,save_pdf)
        import subprocess                           # dùng subprocess để mở file thì có thể vừa mở file vừa thao tác trên giao diện GUI, có thể mở được nhiều file cùng 1 lúc

        'Tạo điều kiện nếu khi lưu file bỏ mất chữ pdf'
        if '.pdf' not in save_pdf:
            subprocess.Popen([save_pdf+'.pdf'],shell=True)
        else:
            subprocess.Popen([save_pdf],shell=True)
    else:
        pass

Btn_save.configure(command=xuat_pdf)

window_main.mainloop()

'thoát chương trình'
sh_chi_tieu_1.clear_contents()
sh_chi_tieu_2.clear_contents()
sh_chi_tieu_3.clear_contents()
sh_print.clear_contents()
sh_print.range('1:500').unmerge()
sh_data.range('B1').clear_contents()
sh_data.range('D1:E1').clear_contents()

sh_chi_tieu_1.visible = False
sh_chi_tieu_2.visible = False
sh_chi_tieu_3.visible = False
sh_print.visible = False

wb1.save()
for app in xw.apps:
    app.quit()