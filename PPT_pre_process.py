import tkinter as tk
import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
'''GUI는 tkinter의 응용프로그램. 자동화 구상도 참고. onedrive(uos)/포트폴리오/자동화/자동화 구상도.jpg'''
GUI=tk.Tk()
GUI.geometry('1200x800')
'''기본적인 프레임 구성 및 버튼  등등 구성==>'''
Filter_Frame=tk.Frame(GUI)
Result_Frame=tk.Frame(GUI)
Bottom_Frame=tk.Frame(GUI)

Filter_Listbox_1=tk.Listbox(Filter_Frame,width=30,height=20)
Filter_Listbox_2=tk.Listbox(Filter_Frame,width=30,height=20)
Filter_Listbox_3=tk.Listbox(Filter_Frame,width=30,height=20)
Filter_Listbox_4=tk.Listbox(Filter_Frame,width=30,height=20)


'''그리드 상태'''
Filter_Listbox_1.grid(row=0,column=0)
Filter_Listbox_2.grid(row=0,column=1)
Filter_Listbox_3.grid(row=0,column=2)
Filter_Listbox_4.grid(row=0,column=3)
Filter_Frame.grid(row=0,column=0)
Result_Frame.grid(row=0,column=1)
Bottom_Frame.grid(row=1,column=0,columnspan=2)

GUI.mainloop()
'''위치조정 : 바텀프레임은 columnspan을 통해 결합'''
