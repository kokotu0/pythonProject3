import pathlib
import tkinter as tk
import tkinter.font
import tkinter.filedialog
import tkinter.messagebox
import pickle
from merge_and_release import *
import os


try:
    with open(file=f'{os.getenv("LOCALAPPDATA")}\\setup_file.pickle', mode='rb') as f:
        setup=pickle.load(f)
except:
    default_setup = {'coupang_supply_id': 'pm2100', 'coupang_supply_password': 'cleanfeel6!', 'open_dir': './',
                     'save_dir': './','geometry':'752x860+78+78'}
    setup=default_setup


root=tk.Tk()
try:
    root.geometry(setup['geometry'])
except:pass
root.title("Simple application")

file_frame_lbl=tk.Label(root,text="파일명")

file_frame_lbl.grid(row=0,column=0)
file_frame=tk.Frame(root,relief='groove',width=350,height=600)
file_frame.grid(row=1,column=0)
file_list=tk.Listbox(file_frame,width=80,height=50)
file_list.grid()
yscrollbar=tk.Scrollbar(file_frame,orient='v',command=file_list.yview)
yscrollbar.grid(row=0,column=1,sticky='ns')
file_list.configure(yscrollcommand=yscrollbar.set)
button_font=tkinter.font.Font(size=15)
button_frame=tk.Frame(root,relief='ridge',width=100,height=600)
button_frame.grid(row=1,column=1)

def open_files():
    global setup
    overlap=0
    result=tkinter.filedialog.askopenfilenames(initialdir=setup['open_dir'],title="파일 선택",filetypes=(("xslx files","*.xlsx"),("all files","*.*")))
    #중복 자동제거
    open_dir=str(pathlib.Path(result[0]).parent)
    result=list(set(result))
    for file_name in result:
        if file_name in file_list.get(0,tk.END):
            overlap+=1
            continue
        else:
            file_list.insert(tk.END,file_name)
    if overlap!=0 : tkinter.messagebox.showinfo(title="중복 확인",message="중복되는 항목 {}개를 제외하고 추가하였습니다.".format(overlap))
    setup['open_dir']=open_dir
def delete_list():
    file_list.delete(file_list.curselection())
def delete_all():
    count=file_list.size()
    for i in range(count):
        file_list.delete(tk.END)
from dewbel_coupang_order_automation_selenium import selenium_order_list_save
def save_file():
    global setup
    global directory_set
    global save_file_name
    global file_names
    save_file_name=tkinter.filedialog.asksaveasfilename(initialdir=setup['save_dir'],defaultextension=".xlsx",filetypes=(("xslx files","*.xlsx"),("all files","*.*")))
    # print(save_file_name)
    save_dir=str(pathlib.Path(save_file_name).parent)

    # table
    directory_set=file_list.get(0,tk.END)

    print(directory_set)
    file_names=[pathlib.Path(file_path).name for file_path in directory_set]
    print(directory_set)
    coupang_supply_id=id_entry.get()
    coupang_supply_password=password_entry.get()
    try:
        PO_table=selenium_order_list_save(coupang_supply_id=coupang_supply_id,coupang_supply_password=coupang_supply_password,path=save_dir)
    except Exception as e:
        print(e)
        tk.messagebox.showerror('에러발생','에러가 발생했습니다. 재실행하거나 id 및 비밀번호를 확인해주세요!')
        return
    merge_and_release(directory_set,save_file_name,PO_table=PO_table)

    tkinter.messagebox.showinfo(message="작업 완료")


    setup['save_dir']=save_dir
    setup['coupang_supply_id']=coupang_supply_id
    setup['coupang_supply_password']=coupang_supply_password
    with open(file=f'{os.getenv("LOCALAPPDATA")}\\setup_file.pickle', mode='wb') as f:
        pickle.dump(setup, f)
def destory_(root=root):
    coupang_supply_id=id_entry.get()
    coupang_supply_password=password_entry.get()

    setup['coupang_supply_id']=coupang_supply_id
    setup['coupang_supply_password']=coupang_supply_password
    setup['geometry']=(root.winfo_geometry())

    print(setup)
    with open(file=f'{os.getenv("LOCALAPPDATA")}\\setup_file.pickle', mode='wb') as f:
        pickle.dump(setup, f)
    root.destroy()
    pass



button_load=tk.Button(button_frame,text="불러오기",command=open_files)
button_load.grid(row=0,column=0,sticky='we')
button_delete=tk.Button(button_frame,text="선택 제거",command=delete_list)
button_delete.grid(row=1,column=0,sticky='we')
button_delete_all=tk.Button(button_frame,text="전체 제거",command=delete_all)
button_delete_all.grid(row=3,column=0,sticky='we')

button_save=tk.Button(button_frame,text="병합 후 내보내기",command=save_file)
button_save.grid(row=4,column=0,sticky='we')

button_exit=tk.Button(button_frame,text="종료하기",command=destory_)
button_exit.grid(row=5,column=0,sticky='we')

id_label=tk.Label(button_frame,text="id",)
id_label.grid(row=6,column=0,sticky='we')

id_entry=tk.Entry(button_frame,)
id_entry.insert(0,setup['coupang_supply_id'])
id_entry.grid(row=7,column=0,sticky='we')

password_label=tk.Label(button_frame,text="쿠팡서플라이어허브_패스워드",)
password_label.grid(row=8,column=0,sticky='we')

password_entry=tk.Entry(button_frame,)
password_entry.insert(0,setup['coupang_supply_password'])
password_entry.grid(row=9,column=0,sticky='we')

print(password_entry.get())
file_frame_lbl['font']=button_font
button_load['font']=button_font
button_delete['font']=button_font
button_delete_all['font']=button_font
button_save['font']=button_font
button_exit['font']=button_font


root.mainloop()

