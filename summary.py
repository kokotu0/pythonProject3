import tkinter as tk
import tkinter.font
import tkinter.filedialog
import tkinter.messagebox

from merge_and_release import *

root=tk.Tk()
root.title("Simple application")

file_frame_lbl=tk.Label(root,text="파일명")

file_frame_lbl.grid(row=0,column=0)
file_frame=tk.Frame(root,relief='groove',width=350,height=600)
file_frame.grid(row=1,column=0)
file_list=tk.Listbox(file_frame,width=80,height=50)
file_list.grid()
button_font=tkinter.font.Font(size=15)
button_frame=tk.Frame(root,relief='ridge',width=100,height=600)
button_frame.grid(row=1,column=1)

def open_files():
    overlap=0
    result=tkinter.filedialog.askopenfilenames(initialdir='/',title="파일 선택",filetypes=(("xslx files","*.xlsx"),("all files","*.*")))
    #중복 자동제거
    result=list(set(result))
    for file_name in result:
        if file_name in file_list.get(0,tk.END):
            overlap+=1
            continue
        else:
            file_list.insert(tk.END,file_name)
    if overlap!=0 : tkinter.messagebox.showinfo(title="중복 확인",message="중복되는 항목 {}개를 제외하고 추가하였습니다.".format(overlap))
def delete_list():
    file_list.delete(file_list.curselection())
def delete_all():
    count=file_list.size()
    for i in range(count):
        file_list.delete(tk.END)

def save_file():

    save_file_name=tkinter.filedialog.asksaveasfilename(initialdir='/',defaultextension=".xlsx",filetypes=(("xslx files","*.xlsx"),("all files","*.*")))
    print(save_file_name)
    directory_set=file_list.get(0,tk.END)
    print(directory_set)
    merge_and_release(directory_set,save_file_name)
    tkinter.messagebox.showinfo(message="작업 완료")


button_load=tk.Button(button_frame,text="불러오기",command=open_files)
button_load.grid(row=0,column=0,sticky='we')
button_delete=tk.Button(button_frame,text="선택 제거",command=delete_list)
button_delete.grid(row=1,column=0,sticky='we')
button_delete_all=tk.Button(button_frame,text="전체 제거",command=delete_all)
button_delete_all.grid(row=3,column=0,sticky='we')

button_save=tk.Button(button_frame,text="병합 후 내보내기",command=save_file)
button_save.grid(row=4,column=0,sticky='we')

button_exit=tk.Button(button_frame,text="종료하기",command=root.destroy)
button_exit.grid(row=5,column=0,sticky='we')

file_frame_lbl['font']=button_font
button_load['font']=button_font
button_delete['font']=button_font
button_delete_all['font']=button_font
button_save['font']=button_font
button_exit['font']=button_font


root.mainloop()