# -*- coding: utf-8 -*-
"""
Created on Fri Oct  9 05:28:24 2020

@author: Sreekanth
"""

import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import xlwings as xw
from xlwings import constants
import shutil
import os
from datetime import datetime

class BudgetPlaner():
    
    def __init__(self):
        self.PATH = r'D:\Sreekanth\DataScience\Python\Projects\03_Budget_Planner'
        #self.bg = '#E64D2C'
        self.bg = '#0294A5'
        self.fg = 'white'
        #self.fg = '#A79C93'

    def create_main_window(self):

        window = tk.Tk()
        window.title('Budget Planner')
        window.geometry('530x680')
        window.resizable(width=False, height=False)
        window.config(bg=self.bg)
        
        fontStyle = tk.font.Font(size=20)
        lb_main = tk.Label(window, text='BUDGET PLANNER', fg=self.fg, bg=self.bg, font=fontStyle)
        lb_main.pack()
        
        frame1 = tk.Frame(window,bg=self.bg)
        self.frame2 = tk.Frame(window, bg=self.bg)
        frame3 = tk.Frame(window, bg=self.bg)
        frame4 = tk.Frame(window, bg=self.bg)
        frame5 = tk.Frame(window, bg=self.bg)
        frame6 = tk.Frame(window, bg=self.bg)
        
        frame1.pack(fill='both', padx=10, pady=5)
        self.frame2.pack()
        frame3.pack(fill='both', padx=10, pady=5)
        frame4.pack(fill='both', padx=10, pady=5)
        frame5.pack(fill='both', expand='yes', padx=10, pady=5)
        frame6.pack(fill='both', padx=10, pady=5)
                
        lf_create = tk.LabelFrame(frame1, text='CREATE NEW FILE', fg=self.fg, bg=self.bg)
        lf_load = tk.LabelFrame(frame1, text='LOAD FILE', fg=self.fg, bg=self.bg)
        lf_view = tk.LabelFrame(frame3, text='PROJECT DATA', fg=self.fg, bg=self.bg)
        lf_add = tk.LabelFrame(frame4, text='ADD DATA', fg=self.fg, bg=self.bg)
        self.lf_info = tk.LabelFrame(frame4, text='BUDGET INFORMATION', fg=self.fg, bg=self.bg)
        lf_budget = tk.LabelFrame(frame6, text='MODIFY BUDGET', fg=self.fg, bg=self.bg)
         
        lf_create.pack(side=tk.LEFT, padx=10)
        lf_load.pack(side=tk.RIGHT, padx=10)
        lf_view.pack(fill='both', padx=10, pady=5)
        lf_add.pack(side=tk.LEFT, padx=10)
        self.lf_info.pack(side=tk.RIGHT, padx=10)
        lf_budget.pack(pady=5)
                
        # Create file area
        lb_create = tk.Label(lf_create, text='ENTER FILE NAME', width=15)
        lb_create.grid(row=0, column=0, padx=5, pady=5)
        
        ent_create = tk.Entry(lf_create, width=15)
        ent_create.grid(row=0, column=1, padx=5)
        
        # Create new file 
        def create_file(filename):
            
            # Check if file with same name exists
            if os.path.exists(self.PATH + '\\' + '02_Input' + '\\' + filename + '.xlsx'):
                messagebox.showwarning('Warning', 'A file with same name exist.')
            elif len(filename) == 0 :
                messagebox.showerror('Error', 'Provide a file name.')
            else:            
                infile = self.PATH + '\\' + '02_Input' + '\\' + 'template.xlsx'
                outfile = self.PATH + '\\' + '02_Input' + '\\' + filename + '.xlsx'
                shutil.copy(infile, outfile)
                ent_create.delete(0, tk.END) 
                messagebox.showinfo('Info', 'New file created.')
                  
        but_create = tk.Button(lf_create, text='CREATE', width=10,
                               command=lambda : create_file(ent_create.get()))
        but_create.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Load file area
        lb_load = tk.Label(lf_load, text='ENTER FILE NAME', width=15)
        lb_load.grid(row=0, column=0, padx=5, pady=5)
        
        ent_load = tk.Entry(lf_load, width=15)
        ent_load.grid(row=0, column=1, padx=5)
        
        # Load the file and display data
        def load_file(filename):
            
            self.fn = filename
            if os.path.exists(self.PATH + '\\' + '02_Input' + '\\' + filename + '.xlsx'):
                self.read_file(filename)
                self.close_file()
                ent_load.delete(0, tk.END)
                messagebox.showinfo('Info', 'File loaded.')
            elif len(filename) == 0 :
                messagebox.showerror('Error', 'Provide a file name.')
            else:
                messagebox.showerror('Error', 'File does not exist.')
        
        but_load = tk.Button(lf_load, text='LOAD', width=10, 
                             command=lambda:load_file(ent_load.get()))
        but_load.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Project data area
        lb_fileinfo = tk.Label(self.frame2, text='Current File Name', width=30)
        lb_fileinfo.grid(row=0, column=0, padx=10, pady=5)
        #lb_fileinfo.place(x=10)

        self.lb_filename = tk.Label(self.frame2, text='NO FILES LOADED', width=30)
        self.lb_filename.grid(row=0, column=1, padx=10, pady=5)
        #lb_filename.place(x=20, y=25)
        
        
        col_dict = {0:['Sl No', 50], 1:['Date', 95], 2:['Activity', 240], 3:['Cost', 80]}
        self.trv = tk.ttk.Treeview(lf_view, columns=col_dict.keys(), show='headings', height=8)
        #self.trv.pack(fill='both', padx=10, pady=10)
        self.trv.grid(row=1, columnspan=5, column=0, padx=10, pady=10)
        
        for c in range(4):
            self.trv.heading(c, text= col_dict[c][0])
            self.trv.column(c, width=col_dict[c][1]) 

        # Bind for single click, action is defined by single_click() function.
        self.trv.bind('<ButtonRelease-1>', self.single_click)
        
        # Bind for double click, action is defined by self.select_item function.
        self.trv.bind('<Double-1>', self.double_click)            
        
        # New data area
        lb1 = tk.Label(lf_add, text='Sl No', width=10)
        lb1.grid(row=0, column=0, padx=5, pady=5)
        self.entar1 = tk.Entry(lf_add, width = 20)
        self.entar1.grid(row=0, column=1, padx=5, pady=5)
        
        lb2 = tk.Label(lf_add, text='Date', width=10)
        lb2.grid(row=1, column=0, padx=5, pady=5)
        self.entar2 = DateEntry(lf_add, width=17, date_pattern='dd.mm.yyyy')
        self.entar2.grid(row=1, column=1)
        
        lb3 = tk.Label(lf_add, text='Activity', width=10)
        lb3.grid(row=2, column=0, padx=5, pady=5)
        self.entar3 = tk.Entry(lf_add, width = 20)
        self.entar3.grid(row=2, column=1, padx=5, pady=5)
        
        lb4 = tk.Label(lf_add, text='Cost', width=10)
        lb4.grid(row=4, column=0, padx=5, pady=5)
        self.entar4 = tk.Entry(lf_add, width = 20)
        self.entar4.grid(row=4, column=1, padx=5, pady=5)
        self.entar4.insert(0, '0.0')
        
        self.info_area()
        
        # Create various buttons
        def add_data():
           
            self.read_file(self.fn)
            # get all the sl_no from the file
            slno_range = 'A5:' + 'A' + str(self.nrows)
            slno_list = self.sheet.range(slno_range).value
            # Convert 'slno_list' to datatype list if there is only one row in the file
            if type(slno_list) == float:
                slno_list = [slno_list]
                
            new_row = self.nrows + 1
            self.close_file()
            
            try:
                int(self.entar1.get())
                float(self.entar4.get())
            except ValueError:
                messagebox.showerror('Error', 'Sl No or Cost not a valid value.')
                return
            
            if self.entar1.get() == '' or self.entar2.get() == '' or self.entar3.get() == '' or self.entar4.get() == '':
                messagebox.showerror('Error', 'Missing mandatory values.')
            elif float(self.entar1.get()) in slno_list:
                messagebox.showerror('Error', 'Sl No should be unique.')
            else:
                self.read_file(self.fn)
                self.sheet.range(new_row, 1).value = self.entar1.get()
                self.sheet.range(new_row, 2).value = self.entar2.get()
                self.sheet.range(new_row, 3).value = self.entar3.get()
                self.sheet.range(new_row, 4).value = self.entar4.get()
                
                self.close_file(param='write')
                clear_entry()
                messagebox.showinfo('Info', 'Data added.')       
            
        but_add = tk.Button(frame5, text='ADD', width=12, command=add_data)
        but_add.place(x=10, y=8)
        
        def update_data():
                    
            self.read_file(self.fn)
            # get all the sl_no from the file
            slno_range = 'A5:' + 'A' + str(self.nrows)
            slno_list = self.sheet.range(slno_range).value
            # Convert 'slno_list' to datatype list if there is only one row in the file
            if type(slno_list) == float:
                slno_list = [slno_list]
                
            sl_idx = 0
            self.close_file()
            
            try:
                int(self.entar1.get())
                float(self.entar4.get())
            except ValueError:
                messagebox.showerror('Error', 'Sl No or Cost not a valid value.')
                return
            
            if self.entar1.get() == '' or self.entar2.get() == '' or self.entar3.get() == '' or self.entar4.get() == '':
                messagebox.showerror('Error', 'Missing mandatory values')
            elif float(self.entar1.get()) not in slno_list:
                messagebox.showerror('Error', 'Sl No does not exist.')
            else:
                for r in range(5, self.nrows+1):    
                    if float(self.entar1.get()) == slno_list[sl_idx]:
                        self.read_file(self.fn)
                        self.sheet.range(r, 2).value = self.entar2.get()
                        self.sheet.range(r, 3).value = self.entar3.get()
                        self.sheet.range(r, 4).value = self.entar4.get()                       
                        self.close_file(param='write')
                        clear_entry()
                        messagebox.showinfo('Info', 'Data updated.')  
                        break
                    else:
                        sl_idx += 1

        but_update = tk.Button(frame5, text='UPDATE', width=12, command=update_data)
        but_update.place(x=142, y=8)
        
        def delete_data():
            self.read_file(self.fn)
            sl_idx = 0
            slno_range = 'A5:' + 'A' + str(self.nrows)
            slno_list = self.sheet.range(slno_range).value
            slno_to_delete = self.item['values'][0]
            if float(slno_to_delete) in slno_list:
                for r in range(5, self.nrows+1):
                    if float(slno_to_delete) == slno_list[sl_idx]:
                        del_range = 'A' + str(r) + ':' + 'D' + str(r)
                        self.sheet.range(del_range).api.Delete(constants.DeleteShiftDirection.xlShiftUp)
                        self.close_file(param='write')
                        clear_entry()
                        messagebox.showinfo('Info', 'Selected row deleted.')
                        break                      
                    else:
                        sl_idx += 1
            else:        
                self.close_file()
                messagebox.showerror('Error', 'Invalid Sl No.')
        
        but_delete = tk.Button(frame5, text='DELETE', width=12, command=delete_data)
        but_delete.place(x=274, y=8)       
        
        # Clear entry field / set it to defaul value
        def clear_entry():
            self.entar1.delete(0, tk.END)
            self.entar2.set_date(datetime.today())
            self.entar3.delete(0, tk.END)
            self.entar4.delete(0, tk.END)
            self.entar4.insert(0, '0.0')
        
        but_clear = tk.Button(frame5, text='CLEAR', width=12, command=clear_entry)
        but_clear.place(x=406, y=8)     
        
        # Create budget update area
        def modify_budget():
            
            self.read_file(self.fn)
            self.sheet.range(1, 2).value = ent_budget.get()
            self.close_file('write')
            ent_budget.delete(0, tk.END)
            ent_budget.insert(0, '0.0')
            messagebox.showinfo('Info', 'Budget amount updated.')          
              
        lb_budget = tk.Label(lf_budget, text='Budget Amount', width=20)
        lb_budget.grid(row=0, column=0, padx=10, pady=5)
        
        ent_budget = tk.Entry(lf_budget, width=10)
        ent_budget.grid(row=0, column=1, padx=10)
        ent_budget.insert(0, '0.0')
        
        bt_budget = tk.Button(lf_budget, text='ADD / UPDATE', width=12,
                              command=modify_budget)
        bt_budget.grid(row=0, column=2, padx=10)
        
        window.mainloop()
        
    # Information area
    def info_area(self, budget_amt=0, total_spent=0, balance=0, status='Good'): 
                
        lb_info1 = tk.Label(self.lf_info, text='Budget allocated', width=20)
        lb_info1.grid(row=0, column=0, padx=5, pady=5)
        lb_amt = tk.Label(self.lf_info, text= budget_amt, width = 10)
        lb_amt.grid(row=0, column=1, padx=5, pady=5)
        
        lb_info2 = tk.Label(self.lf_info, text='Total Spent', width=20)
        lb_info2.grid(row=1, column=0, padx=5, pady=5)
        lb_tot = tk.Label(self.lf_info, text= total_spent, width = 10)
        lb_tot.grid(row=1, column=1, padx=5, pady=5)
        
        lb_info3 = tk.Label(self.lf_info, text='Balance Available', width=20)
        lb_info3.grid(row=2, column=0, padx=5, pady=5)
        lb_bal = tk.Label(self.lf_info, text= balance, width = 10)
        lb_bal.grid(row=2, column=1, padx=5, pady=5)
        
        lb_info4 = tk.Label(self.lf_info, text='Budget Status', width=20)
        lb_info4.grid(row=3, column=0, padx=5, pady=5)
        lb_stat = tk.Label(self.lf_info, text= status, width = 10)
        lb_stat.grid(row=3, column=1, padx=5, pady=5)
    
    # Read file
    def read_file(self, filename):
        
        self.app = xw.App(visible=False, add_book=False)
        self.file = self.app.books.open(self.PATH + '\\' + '02_Input' + '\\' + self.fn + '.xlsx')
        self.sheet = self.file.sheets['data']
        self.nrows = self.sheet.range(1,1).end('down').row
    
    # Save and close file
    def close_file(self, param='read'):
        
        if param == 'write':
            self.file.save()
            self.nrows = self.sheet.range(1,1).end('down').row
        
        self.get_amt_data()
        self.display_data()
        self.file.close()
        self.app.quit()
    
    # Get data to populate information area    
    def get_amt_data(self):
        
        budget_amt = self.sheet.range('B1').value
        total_spent = self.sheet.range('B2').value
        balance = self.sheet.range('B3').value
        status = self.sheet.range('E1').value
        
        self.info_area(budget_amt, total_spent, balance, status)

    # Display data in treeview area
    def display_data(self):
           
        self.lb_filename = tk.Label(self.frame2, text=self.fn, width=30)
        self.lb_filename.grid(row=0, column=1, padx=10, pady=5)
        
        self.trv.delete(*self.trv.get_children())
        data = []
        for r in range(5, self.nrows+1):   
            sl_no = int(self.sheet.range(r, 1).value)
            date = self.sheet.range(r, 2).value
            activity = self.sheet.range(r, 3).value
            cost = self.sheet.range(r, 4).value
            data = [sl_no, date, activity, cost]
            self.trv.insert('', tk.END, values=data)

    # Function for tree view bind single click.
    def single_click(self, event):
        self.item = self.trv.item(self.trv.focus())

    # Function for tree view bind double click. 
    def double_click(self, event):
        try:
            self.item = self.trv.item(self.trv.focus())
            self.entar1.delete(0, tk.END)
            self.entar1.insert(0, self.item['values'][0])
            self.entar2.set_date(self.item['values'][1])
            self.entar3.delete(0, tk.END)
            self.entar3.insert(0, self.item['values'][2])
            self.entar4.delete(0, tk.END)   
            self.entar4.insert(0, self.item['values'][3])
        except IndexError:
            pass
        
        

app = BudgetPlaner()
app.create_main_window()