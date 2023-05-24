from tkinter import *
from tkinter import ttk, filedialog
from tkinter.messagebox import showerror
import pandas as pd
import thefuzz
from thefuzz import process
import re

property_key = ['4010 SOUTHERN AVE', '114 DARRINGTON SW', '1338 NORTH CAROLINA', 
        '1634 NICHOLSON ST NW', '1657 FOREST PARK DR', '4335 VAN NESS ST', 
        '4617 BROOKFIELD DR','4702 TEAK COURT', '5813 DIX ST', 
        '600 MARKHAM RD', '912 45TH PL', '9270 CHERRY LANE', 
        'UNK']

instructions = '''
1. This program assumes CSV columns and header match the home 
    depot raw file.
2. Open desired file by clicking Select File Button
3. Information in file will transformed in to target output.
4. To get results, click Save As Button, pick destination folder 
   enter file name.
Note: You can create a new file or overwrite and existing file
Note: you need to select a file before trying to save
Note: the original CSV is not altered in any way
Note: the output excel file will have all data on first sheet 
      and then a sheet for each group of data
Note: Job names are cleaned: 0, na, N/A are changed to UNK, Job names
      are matched to a pre-existing list.  If name doesn't match 
      with a certain rank, it will remain the same.
Note: Matched names are in new column can be compared to Original
'''


def runGUI():

    root = Tk()
  
    text_frame = Frame(root, width=200, height=400)
    text_frame.grid(row=0, column=0, padx=10, pady=5)

    file_frame = Frame(root, width=200, height=100)
    file_frame.grid(row=1, column=0, padx=10, pady=5)

    quit_frame = Frame(root, width=200, height=100)
    quit_frame.grid(row=2, column=0, padx=10, pady=5)

    instruction_text = Text(text_frame)
    instruction_text.pack(expand=True)
    instruction_text.insert('end', instructions)
    instruction_text.config(state='disabled')

    #frm = ttk.Frame(root, padding=10)
    #frm.grid()

    #ttk.Button(frm, text='Select File', command=getfilename).grid(column=0, row=0)
    #ttk.Button(frm, text='Save As', command=setnewfilename).grid(column=1, row=0)
    #ttk.Button(frm, text='Quit', command=root.destroy).grid(row=1)
    
    openfilebutton = Button(file_frame, text='Select File', command=getfilename).pack(side='left', padx=10)
    saveasbutton = Button(file_frame, text='Save As', command=setnewfilename).pack(side='right', padx=10)
    Button(quit_frame, text='Quit', command=root.destroy).pack(side='bottom', pady=10)
    
    #openfilebutton.pack()
    #saveasbutton.pack()
    
    root.mainloop()


def updatecsv(filename):
    #filename='Purchase_History_May-11-2023_2꞉31-PM.csv'

    global my_df 
    
    my_df = pd.read_csv(filename, header=5)

    #remove unwanted columns
    my_df=my_df.drop(['Register Number', 'Receipt Added Date', 'Purchaser', 
                    'Program Disc Amt', 'Other Disc', 'Text2Confirm', 'Order Origin', 
                    'Card/Account Nickname', 'Buyer Name-ID', 'Pre-tax Amount'], axis=1)
    
    #add 'UNK' for empty job names
    my_df['Job Name'].fillna(value='UNK', inplace=True)

    #change na, n/a to UNK
    par = re.compile('^n/?a$|^0$', flags=re.I)
    my_df['Job Name']= my_df['Job Name'].str.replace(par, 'UNK', regex=True)

    #fuzzy match job name
    my_df['Clean Job Name']=my_df['Job Name'].apply(lambda x: fuzz_m(property_key, x))
     
    print(my_df.head())

def getfilename():
    filetypes = (
        ('CSV', '*.csv'),
        ('All files', '*.*')
    )

    filename = filedialog.askopenfilename(
        title='Open File',
        filetypes=filetypes)
    
   
    updatecsv(filename)
    
#write dataframe to excel file
def setnewfilename():
    try:
        my_df
    except NameError:
        showerror(title='Error', message='You must first select a CSV file')

    else:
        filetypes = (
            ('.xlsx', '*.xlsx'),
            ('All files', '*.*')
        )

        filename = filedialog.asksaveasfilename(
            title='Open File',
            filetypes=filetypes,
            defaultextension='.xlsx')

        with pd.ExcelWriter(filename) as writer:
            #write all data to one sheet
            my_df.to_excel(writer, sheet_name='All Data', index=False)

            #write group of data to a new sheet named for that group
            df_groups = my_df.groupby(['Clean Job Name'])
            for name in df_groups.groups:
                df_groups.get_group(name).to_excel(writer, sheet_name=name)



#method for fuzzymatch job name to list
def fuzz_m(key, target):
    match, score = process.extractOne(target, key)
    if score < 80:
        return target
    else:
        return match        
  

if __name__=="__main__":
    runGUI()