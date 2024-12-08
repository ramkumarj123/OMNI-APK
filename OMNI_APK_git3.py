import pandas as pd
from tkinter import *
import tkinter.messagebox as tkmessagebox
from tkinter import filedialog
from customtkinter import*
from PIL import Image
from datetime import datetime
import os
import zipfile
import warnings
from tkinter import messagebox
warnings.filterwarnings("ignore")


imgae_loaction = r"C:\Users\kura2020\OneDrive - NIQ\Automation\New ARS\image/"

img = Image.open( imgae_loaction+  "report-icon-8.jpg")
img2 = Image.open(imgae_loaction+  "upload-file-icon-5.jpg")
app = CTk()
app.title("OMNI - 3.0")
set_appearance_mode("dark")
app.geometry("800x500")



app.iconbitmap(imgae_loaction+  "nielseniq-svg (1).ico")

#ADD Tab
tabview = CTkTabview(app,width=820,height=600)
tabview.pack(padx=20,pady=20)

tabview.add(" ARS FORMATING ")
tabview.add(" ARS ")
tabview.add(" COUNT ")
tabview.add(" CONSISTENCY ")
tabview.add(" FILE SPLIT ")
tabview.add(" IVD ")
tabview.add(" MAP VALUE ")


tab4 = CTkLabel(tabview.tab(" COUNT "),text="",width=1400,height=1000)
tab4.pack(padx=20,pady=20)

tab1 = CTkLabel(tabview.tab(" ARS FORMATING "),text="",width=1400,height=1000)
tab1.pack(padx=20,pady=20)

footer_label = CTkLabel(app, text="""   This application was designed and developed by ram.kumar@nielseniq.com @ 2024                                                                                                                                                                        """, text_color="#B2B2B2", font=("Arial", 10))
footer_label.place(relx=0.0, rely=1.0, anchor="sw", x=0.1, y=-0.1)  # Position at bottom-right corner

frame2 = CTkFrame(master=tab1,width=450,height=45,border_width=.6)
frame2.place(relx=0.18,rely=.75)

frame4 = CTkFrame(master=tab4,width=450,height=40,border_width=.6)
frame4.place(relx=0.10,rely=.75)


file_name_to_save = StringVar()
FR_FILE_PATH =  StringVar()
FR_File_NAME = StringVar()
FR_FILE =  StringVar()

def exit():
    Exit = tkmessagebox.askyesno("ARS Report",'Confim Do You Want To Exit',)
    if Exit>0:
        app.destroy()
        return
    

    
def Changes_file():
    F_filepath = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")]) 
    F_filename = os.path.split(F_filepath)
    FR_FILE_PATH.set(F_filename[0]) #--- File path only
    FR_File_NAME.set(F_filename[1]) #--- File name only
    FR_FILE.set(F_filepath) #--- Full file path and file name
    label.configure(text = "")

def FR_clear_all():
     if C_report_Entry.get() != "" :
          FR_File_NAME.set("")
          label.configure(text="""
                              Cleared!
""",anchor="center")
def arss():
     
     if C_report_Entry.get() == "":
          label.configure(text="""
                    Insert the files, please..
""",anchor="center")

     else:
          #file_name_to_reads = 'Extraction of mapped chars.CSV'
          with zipfile.ZipFile(FR_FILE.get(), 'r') as zip_ref:
                    with zip_ref.open(zip_ref.namelist()[0]) as FILE:
                         try:
                              data1 = pd.read_csv(FILE)

                         except:
                              data1 = pd.read_excel(FILE)
          for i in data1.columns:
               if i in ["PREV_IMDB" , "CUR_DATE", "PREV_DATE", "ITM_CREATION_DATE"]:
                    data1[i]= " "

          for dates in data1["REFERENCE_DATE"]:
               try:
                    orginal_date = "'"+dates.split(" ")[0] + " 12:00:00 AM"
# Parse the original string to a datetime object
                    datetime_obj = datetime.strptime(orginal_date, "%m/%d/%Y %I:%M:%S %p")
                    
# Format the datetime object to the desired string format
                    new_date_str = datetime_obj.strftime("%Y-%m-%d %H:%M:%S")
                    data1["REFERENCE_DATE"] = new_date_str
               except:
                    orginal_date = "'"+dates.split(" ")[0] + " 00:00:00 AM"
          
               data1["REFERENCE_DATE"] = orginal_date
               break


# Comments Report for POS_MAC Column 

          data2 = pd.DataFrame()
          data2.at[0,"CRC_ID"] = " "
          data2.at[0,"CRC_COMMENT"] = " "
          data2 = data2.drop(index=0).reset_index(drop=True)


#parameters

          parameter = {
          "Token": [
               "SYS_COU_CODE", "USR_COU_CODE", "SYS_COU_DESC", "USR_COU_DESC",
               "USR_NAME", "USR_ID", "SYS_LOG_ID", "posRefDate", "posCode",
               "currentWorkTeamCountryId"
          ],
          "Token Description": [
               "User Country Code (system)", "User Country Code", 
               "User Country Description (system)", "User Country Description",
               "User Name", "User ID", "eForte Log ID", "posrefdate", "poscode",
               "currentworkteamcountryid"
          ],
          "Value": [
               "US", "US", "United States", "United States", "", 
               "", "2192029", str(data1["REFERENCE_DATE"][0]), data1["RUN_ID"][0].split(":")[1][1:], "99"
          ],
          "Value Label": [""] * 10,
          "Operator": [""] * 10,
          "CSV Delimiter": [""] * 10
          }

          parameter = pd.DataFrame(parameter)
          output_file_path = os.path.join(FR_FILE_PATH.get(), "Changes Report for POS_MAC.xlsx")
          with pd.ExcelWriter(os.path.join(output_file_path)) as writer:
               data1.to_excel(writer,sheet_name="Changes Report for POS_MAC",index=False)

               data2.to_excel(writer, sheet_name="Comments Report for POS_MAC",index=False)
               
               parameter.to_excel(writer, sheet_name="parameters",index=False)
               label.configure(text="""
The report has been generated....,
The procedure is finished 100%
                            
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
               messagebox.showinfo('Success',"The reports have been successfully generated.")

          
# Label


F_location = CTkLabel(master=tab1,text = "ARS Formating",font=('Consolas',13.5),text_color="#DEDEDE")
F_location.place(relx=0.03,rely=0.1) 



#Entry


C_report_Entry = CTkEntry(master=tab1,placeholder_text="Changes Report...",textvariable=FR_File_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
C_report_Entry.place(relx=0.247,rely=0.1)



#Buttons


C_report_button = CTkButton(master=tab1,text=" Browse  ",width=10,height=10,command=Changes_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
C_report_button.place(relx=0.81,rely=0.1)

'''def history():
     if history_button.get()>0:
          E_report_button.configure(state=tk.DISABLED)
          E_report_Entry.configure(state=tk.DISABLED)
     else:
          E_report_button.configure(state=tk.NORMAL)
          E_report_Entry.configure(state=tk.DISABLED)'''

"""history_button = CTkCheckBox(tab1,text=" ARS Formating ",checkbox_height=20,checkbox_width=20,command=history,corner_radius=0)
history_button.place(relx=0.81,rely=0.4)"""

generate = CTkButton(master=tab1,text="Format ",width=20,height=10,command=arss,corner_radius=0,image=CTkImage(light_image=img,dark_image=img))
generate.place(relx=0.6,rely=0.39)

clear = CTkButton(tab1,text=" Clear All ",width=10,height=20,command=FR_clear_all,corner_radius=0)
clear.place(relx=0.762,rely=0.4)

exits = CTkButton(tab1,text="  Exit  ",width=10,height=20,command=exit,corner_radius=0,fg_color="#F95A45")
exits.place(relx=0.88,rely=0.4)

#Frame_2
#ComboBox
#option_button = CTkComboBox(master=tab1,values=["Blank","P&G","OTHERS"],font=("Calibri",13),width=110,height=25,text_color="White")
#option_button.place(relx=0.283,rely=0.5)



label = CTkLabel(master=frame2,text = "",width=0,height=0,text_color="#F5AF61",font=('Consolas',13))
label.pack(anchor="center",expand=True)

#Consistency

tab2 = CTkLabel(tabview.tab(" CONSISTENCY "),text="",width=1500,height=1000)
tab2.pack(padx=20,pady=20)

consistence_frame = CTkFrame(master=tab2,width=450,height=45,border_width=.6)
consistence_frame.place(relx=0.18,rely=.75)

CC_FILE_PATH = StringVar()
CC_File_NAME = StringVar()
CC_FILE = StringVar()
Consis_file_name_to_save = StringVar()

# RMS IV:
RMS_IV_PATH = StringVar()
RMS_IV_NAME = StringVar()
RMS_IV_FULL = StringVar()

def consis_exit():
    Exit = tkmessagebox.askyesno("ARS consistency",'Confim Do You Want To Exit',)
    if Exit>0:
        app.destroy()
        return
    


def Iteam_and_value_file():
    Consis_F_filepath = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")]) 
    CCfilename = os.path.split(Consis_F_filepath)
    CC_FILE_PATH.set(CCfilename[0]) #--- File path only
    CC_File_NAME.set(CCfilename[1]) #--- File name only
    CC_FILE.set(Consis_F_filepath) #--- Full file path and file name
    label2.configure(text = "")

def RMS_file_iv():
     RMS_IVPATH = filedialog.askopenfilename(filetypes=[("ZIP files","*.zip")]) 
     RMSFILENAME = os.path.split(RMS_IVPATH)
     RMS_IV_PATH.set(RMSFILENAME[0])
     RMS_IV_NAME.set(RMSFILENAME[1])
     RMS_IV_FULL.set(RMS_IVPATH)


def consis_clear_all():
     if CC_File_NAME.get() != "" or Consis_file_name_to_save.get() != "":
          Consis_file_name_to_save.set("")
          CC_File_NAME.set("")
          RMS_IV_NAME.set("")
          label2.configure(text="""
Cleared!"""
,anchor="center")
          
def consistence_auto():
     if CC_File_NAME.get() == "" or Consis_file_name_to_save.get() == "":
          label2.configure(text="""Fill out columns, please...""")
     else:
          
          with zipfile.ZipFile(CC_FILE.get(), 'r') as zip_ref:
                    with zip_ref.open(zip_ref.namelist()[0]) as FILE_CC:
                         try:
                              consistency = pd.read_csv(FILE_CC,keep_default_na=False)
                         except:
                              consistency = pd.read_excel(FILE_CC,keep_default_na=False)

                    """try:
                         column_name_changes = {'0': 'NAN_KEY',
                                        'Unnamed: 1': 'ITEM_DESCRIPTION',
                                        'Unnamed: 2': 'BARCODE',
                                        'Unnamed: 3': 'WEEK_FROM',
                                        'Unnamed: 4': 'WEEK_TO',
                                        'Unnamed: 5': 'NULL_SALES',
                                        'Unnamed: 6': 'RMS_SALES_52W',
                                        'Unnamed: 7': 'OMNI_VALUE_52W',
                                        'Unnamed: 8': 'RAW_OCCASIONS',
                                        'Unnamed: 9': 'MARKETPLACES_S&S_SALES_53W',
                                        'Unnamed: 10': 'OMNI_SALES_52W',
                                        'Unnamed: 11': 'NANKEY_CREATION_DATE'}
                    except:
                         column_name_changes = {'0': 'NAN_KEY',
                                        'Unnamed: 1': 'ITEM_DESCRIPTION',
                                        'Unnamed: 2': 'BARCODE',
                                        'Unnamed: 3': 'WEEK_FROM',
                                        'Unnamed: 4': 'WEEK_TO',
                                        'Unnamed: 5': 'NULL_SALES',
                                        'Unnamed: 6': 'SALES_52W',
                                        'Unnamed: 7': 'NANKEY_CREATION_DATE'}
                         
                    consistency = consistence_file.rename(columns=column_name_changes)"""
                    

                    # Save column headers
                    all_headers = consistency.iloc[0].tolist()
                    headers = [h for h in all_headers if  str(h).endswith(']') and h != 0]
                    consistency = consistency.iloc[1:, :]
                    columns = consistency.columns.tolist()

                    # Replace columns starting with "char" with new columns from the list
                    new_df_columns = []


                    # Replacing the "char" columns with the new column names
                    for col in columns:
                         if col.startswith("char"):
                              try:
                                   new_df_columns.append(headers.pop(0))  # Use new columns one by one
                              except:
                                   new_df_columns.append("")
                         else:
                              if col != '':
                                   new_df_columns.append(col)

                    # Assign the updated column names to the DataFrame
                  
                    consistency.columns = new_df_columns
                    consistency = consistency.iloc[1:].reset_index(drop=True)

                    #consistency.rename(columns={'nan_key': 'NAN_KEY'}, inplace=True)
                    print(f"{CC_File_NAME.get()} Validation Start..")
                    for delete_column in consistency.columns:
                         if delete_column[:7] == "Unnamed":
                              consistency.drop(columns=delete_column, inplace=True)



                    #Inserting New Columns & Comparision
                    Char_count = 0
                    for find in consistency.columns:
                              Char_count += 1
                              if find.split("[")[0] == "CHARACTERISTIC DETAIL ":
                                   break
                    
                    
                    AI_CHARA = []
                    count = 0

                    for AI_columns in consistency.columns:
                        count += 1
                        if AI_columns.split("[")[0] == "CHARACTERISTIC DETAIL ":
                            AI_CHARA.append(AI_columns)
                            break
                    
                    if DB_TYPE_button.get() == "REBUILD DB":
                        ITEM_SCOPE = []
                        with zipfile.ZipFile(RMS_IV_FULL.get(), "r") as zip_ref:
                            with zip_ref.open(zip_ref.namelist()[0]) as file:
                                RMS_IV_DATA = pd.read_csv(file, skiprows=1)
                            for RMS_IV_NANKEY in consistency["nan_key"]:
                                if RMS_IV_NANKEY in list(RMS_IV_DATA["0"]):
                                        ITEM_SCOPE.append("SAME IN BOTH")
                                else:
                                        ITEM_SCOPE.append("NEWLY ADDED")
                                        
                        consistency.insert(1,"ITEM SCOPE",ITEM_SCOPE)
                    else:
                        pass
                    New_data  = pd.DataFrame()
                    for i in consistency.columns:
                        if i != AI_CHARA[0]:
                            New_data [i] = consistency[i]
                        else:
                            break
                    New_data[AI_CHARA[0]] = consistency[AI_CHARA[0]]

                    remanining_data = pd.DataFrame()
                    for remain_columns in range(count,consistency.shape[1]):
                         try:
                              remanining_data[consistency.columns[remain_columns]] = consistency[consistency.columns[remain_columns]]
                         except:
                             pass
                    remanining_data = remanining_data.sort_index(axis=1)   
                    for add in remanining_data.columns:
                        New_data[add] = remanining_data[add]
                    for ii in range(count,200):
                        try:
                            for ii in range(len(New_data.columns) - 1):
                              # Check if the base column names match
                              if New_data.columns[ii].split("[")[0] == New_data.columns[ii + 1].split("[")[0]:
                              # Create the new column name
                                   new_col_name = f"STATEMENT FOR {New_data.columns[ii + 1].split('[')[0]}"
                                   # Insert the new column with TRUE/FALSE values
                                   New_data[new_col_name] = (New_data[New_data.columns[ii]] == New_data[New_data.columns[ii + 1]]).map({True: "TRUE", False: "FALSE"})
                                   insert_position = New_data.columns.get_loc(New_data.columns[ii + 1]) + 1
                                   New_data.insert(insert_position, new_col_name, New_data.pop(new_col_name))
                                                            
                        except:
                            pass
                    label2.configure(text="""
The report has been generated....,
The procedure is finished 100%
                                   
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
                    
                    cc_out = CC_FILE_PATH.get() + "/" + Consis_file_name_to_save.get().upper()+".xlsx"
                    New_data.to_excel(cc_out, index=False) 
                    messagebox.showinfo('Success',"The reports have been successfully generated.")

# Label
iteam_file = CTkLabel(master=tab2,text = "Iteam and value Report",font=('Consolas',13.5),text_color="#DEDEDE")
iteam_file.place(relx=0.03,rely=0.06)

Db_type = CTkLabel(master=tab2,text = "DB Type",font=('Consolas',13.5),text_color="#DEDEDE")
Db_type.place(relx=0.03,rely=0.18)

RMS_IV = CTkLabel(master=tab2,text = "RMS IV",font=('Consolas',13.5),text_color="#DEDEDE")
RMS_IV.place(relx=0.03,rely=0.31)

iteam_file_save = CTkLabel(master=tab2,text = "Filename to Save",font=('Consolas',13.5),text_color="#DEDEDE")
iteam_file_save.place(relx=0.03,rely=0.44)

"""client_name = CTkLabel(master=tab2,text = "Client Name",font=('Consolas',13.5),text_color="#DEDEDE")
client_name.place(relx=0.03,rely=0.6)"""


#Entry
iv_report_Entry = CTkEntry(master=tab2,textvariable=CC_File_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
iv_report_Entry.place(relx=0.283,rely=0.06)

RMS_report_Entry = CTkEntry(master=tab2,textvariable=RMS_IV_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
RMS_report_Entry.place(relx=0.283,rely=0.31)

save_report_Entry = CTkEntry(master=tab2,textvariable=Consis_file_name_to_save,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
save_report_Entry.place(relx=0.283,rely=0.44)


#Buttons
consis_iv_report_button = CTkButton(master=tab2,text=" Browse  ",width=10,height=10,command=Iteam_and_value_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
consis_iv_report_button.place(relx=0.81,rely=0.06)

RSM_iv_report_button = CTkButton(master=tab2,text=" Browse  ",width=10,height=10,command=RMS_file_iv,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
RSM_iv_report_button.place(relx=0.81,rely=0.31)

#combo box
DB_TYPE_button = CTkComboBox(master=tab2,values=["REBUILD DB","HISTORIC DB"],font=("Calibri",13),width=140,height=25,text_color="White")
DB_TYPE_button.place(relx=0.283,rely=0.18)

"""client_button = CTkComboBox(master=tab2,values=["AI","OTHERS"],font=("Calibri",13),width=140,height=25,text_color="White")
client_button.place(relx=0.283,rely=0.6)"""

#Main Buttons
Consis_generate = CTkButton(master=tab2,text="Run consistency",width=20,height=10,command=consistence_auto,corner_radius=0,image=CTkImage(light_image=img,dark_image=img))
Consis_generate.place(relx=0.78,rely=0.59)

consis_clear = CTkButton(tab2,text=" Clear All",width=10,height=20,command=consis_clear_all,corner_radius=0)
consis_clear.place(relx=0.665,rely=0.6)

consis_exits = CTkButton(tab2,text="  Exit  ",width=10,height=20,command=consis_exit,corner_radius=0,fg_color="#F95A45")
consis_exits.place(relx=0.58,rely=0.6)

#Message
label2 = CTkLabel(master=consistence_frame,text = "",width=500,height=50,text_color="#F5AF61",font=('Consolas',13))
label2.pack(anchor="center",expand=True)


#ARS
tab3 = CTkLabel(tabview.tab(" ARS "),text="",width=1500,height=1000)
tab3.pack(padx=20,pady=20)

ARS_frame = CTkFrame(master=tab3,width=450,height=45,border_width=.6)
ARS_frame.place(relx=0.18,rely=.75)



Extraction_File_LOACTION = StringVar()
EX_FILE_PATH = StringVar()
EX_File_NAME = StringVar()
EX_FILE = StringVar()
#RMS
RMS_File_NAME = StringVar()
RMS_FILE_PATH = StringVar()
RMS_File_NAME = StringVar()
RMS_FILE = StringVar()



def exit():
    Exit = tkmessagebox.askyesno("ARS Report",'Confim Do You Want To Exit',)
    if Exit>0:
        app.destroy()
        return
    
def Extraction_file():
    F_filepath = filedialog.askopenfilename() 
    EXfilename = os.path.split(F_filepath)
    EX_FILE_PATH.set(EXfilename[0]) #--- File path only
    EX_File_NAME.set(EXfilename[1]) #--- File name only
    EX_FILE.set(F_filepath) #--- Full file path and file name
    label4.configure(text = "")
    
def RMS_file():
    RMS_filepath = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")]) 
    RMSfilename = os.path.split(RMS_filepath)
    RMS_FILE_PATH.set(RMSfilename[0])  # File path only
    RMS_File_NAME.set(RMSfilename[1])  # File name only
    RMS_FILE.set(RMS_filepath)

def clear_all():
     if EX_File_NAME.get() != "" or file_name_to_save.get() != "":
          EX_File_NAME.set("")
          file_name_to_save.set("")
          label4.configure(text="""
                              Cleared!
""")


def ars():
     
     if location_entryss.get() == "OLD METHOD":
          if EX_File_NAME.get() == "" or file_name_to_save.get() == "":
               label4.configure(text="""
                    Insert the files, please...
""",anchor="center")

          else:
               label4.configure(text=f"""
Something Went Wrong.""")
               try:
                    data = pd.read_csv(EX_FILE.get())
               except:
                    data = pd.read_excel(EX_FILE.get())
                
          
               data.insert(loc=9, column="BAU_COMMENTS", value="") 

               for lenth in range(len(data["STATUS"])):
                    #RMS
                    if  str(data["EAN_CODE"][lenth])  != "nan":
                         data.at[lenth, "BAU_COMMENTS"] = "RMS ITEM : REVIEWED IN RMS ARS."

               for lenth in range(len(data["STATUS"])):   
                    #NEW      
                    if str(data["EAN_CODE"][lenth]) == "nan" and data["STATUS"][lenth] == "NEW":
                         if str(data["CURRENT_VALUE"][lenth]) == "nan":
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Need to be check."

                         if str(data["CURRENT_VALUE"][lenth]) == "NA" and str(data["MAC_DESCRIPTION"][lenth]) == "GMI_SEGMENT":
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Falling according to the NA rule, Valid."

                         if str(data["CURRENT_VALUE"][lenth]) !="nan" or str(data["CURRENT_VALUE"][lenth]) != "DETAIL UNKNOWN":
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Falling according to the rule, Valid."

                         if str(data["CURRENT_VALUE"][lenth]) == "DETAIL UNKNOWN":
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Falling in DETAIL UNKNOWN since there is no rule, Valid."

                         if str(data["CURRENT_VALUE"][lenth]) == "AO BRAND" or str(data["CURRENT_VALUE"][lenth]) == "AO MANUFACTURER" :
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Falling in AO BRAND/AO MANUFACTURER since there is no rule, Valid"
                              
               for lenth in range(len(data["STATUS"])): 
                    #DELETED
                    if data["STATUS"][lenth] == "DELETED":
                         if str(data["IAS_ITEM"][lenth]) == "nan":
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM : Need to be Check."
                         elif str(data["IAS_ITEM"][lenth]) == "NULL SALES":
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Dropped due to Null sales, Valid." 
                         elif str(data["IAS_ITEM"][lenth]) ==  "NOT IN IAS" :
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Dropped due to NOT IN IAS, Valid."
                    
               for lenth in range(len(data["STATUS"])): 
                    # CHANGED
                    if str(data["EAN_CODE"][lenth]) == "nan" and data["STATUS"][lenth] == "CHANGED" :
                         if str(data["CURRENT_VALUE"][lenth]) == "nan":
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Need to check." 
                         elif  str(data["CURRENT_VALUE"][lenth]) == "DETAIL UNKNOWN":
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM : Need to be Check.Falling into DETAIL UNKNOWN"
                         elif str(data["CURRENT_VALUE"][lenth]) == "PENDING PLACEMENT":
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM : Need to be Check.Falling into PENDING PLACEMENT" 
                         elif str(data["CURRENT_VALUE"][lenth]) == "DETAIL UNKNOWN" and str(data["CUR_RULE"][lenth]) == "DEFAULT":
                              data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM : Need to be Check.Falling into DETAIL UNKNOWN and DEFAULT"
                         elif str(data["CURRENT_VALUE"][lenth]) != "nan":
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Change due to MAC rule change/Value change,Valid."
                                   
               out = EX_FILE_PATH.get() +"/" +   file_name_to_save.get().upper()+".xlsx"
               data.to_excel(out, index=False) 
               label4.configure(text="""
The report has been generated....,
The procedure is finished 100%
                    
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
               messagebox.showinfo('Success',"The reports have been successfully generated.")
     else:
          if EX_File_NAME.get() == "" or file_name_to_save.get() == "":
               label4.configure(text="""
                    Insert the files, please...
""",anchor="center")

          else:
               label4.configure(text=f"""
                         Something Went Wrong.""")
               try:
                    data = pd.read_csv(EX_FILE.get())
               except:
                    data = pd.read_excel(EX_FILE.get())
               item_scope = []
               #file_name_to_read = 'ItemandCharsOmniSales.CSV'
               
               with zipfile.ZipFile(RMS_FILE.get(), 'r') as zip_ref:
                         with zip_ref.open(zip_ref.namelist()[0]) as file:
                              RMS_DATA = pd.read_csv(file, skiprows=1)

                         for RMS_NANKEY in data["ITM_ID"]:
                              if RMS_NANKEY in list(RMS_DATA["0"]):
                                   item_scope.append("SAME IN BOTH")
                              else:
                                   item_scope.append("NEWLY ADDED")
                         data.insert(7, 'ITEM SCOPE', item_scope) 

                              

               # Get the absolute path of the current script/module
                         data.insert(loc=9, column="BAU_COMMENTS", value="") 
                         for lenth in range(len(data["STATUS"])):
                              #RMS
                              if  str(data["ITEM SCOPE"][lenth])  == "SAME IN BOTH":
                                   data.at[lenth, "BAU_COMMENTS"] = "RMS ITEM : REVIEWED IN RMS ARS."


                         for lenth in range(len(data["STATUS"])):   
                                   
                              #NEW      
                              if (str(data["ITEM SCOPE"][lenth]) == "NEWLY ADDED" or "BRAND" in str(data["MAC_DESCRIPTION"][lenth]) or "MANUFACTURER" in str(data["MAC_DESCRIPTION"][lenth])) and data["STATUS"][lenth] == "NEW" and str(data["CURRENT_VALUE"][lenth]) == "nan":
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Need to be check."
                              if (str(data["ITEM SCOPE"][lenth]) == "NEWLY ADDED" or "BRAND" in str(data["MAC_DESCRIPTION"][lenth]) or "MANUFACTURER" in str(data["MAC_DESCRIPTION"][lenth])) and data["STATUS"][lenth] == "NEW" and str(data["CURRENT_VALUE"][lenth]) == "NA" and str(data["MAC_DESCRIPTION"][lenth]) == "GMI_SEGMENT":
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Falling according to the NA rule, Valid."
                              if (str(data["ITEM SCOPE"][lenth]) == "NEWLY ADDED" or "BRAND" in str(data["MAC_DESCRIPTION"][lenth]) or "MANUFACTURER" in str(data["MAC_DESCRIPTION"][lenth])) and data["STATUS"][lenth] == "NEW" and str(data["CURRENT_VALUE"][lenth]) != "DETAIL UNKNOWN":
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Falling according to the rule, Valid."
                              if  (str(data["ITEM SCOPE"][lenth]) == "NEWLY ADDED" or "BRAND" in str(data["MAC_DESCRIPTION"][lenth]) or "MANUFACTURER" in str(data["MAC_DESCRIPTION"][lenth])) and data["STATUS"][lenth] == "NEW" and  str(data["CURRENT_VALUE"][lenth]) !="nan":
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Falling according to the rule, Valid."
                              if (str(data["ITEM SCOPE"][lenth]) == "NEWLY ADDED" or "BRAND" in str(data["MAC_DESCRIPTION"][lenth]) or "MANUFACTURER" in str(data["MAC_DESCRIPTION"][lenth])) and data["STATUS"][lenth] == "NEW" and str(data["CURRENT_VALUE"][lenth]) == "DETAIL UNKNOWN":
                                   data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Falling in DETAIL UNKNOWN since there is no rule, Valid."

                              if "BRAND" in str(data["MAC_DESCRIPTION"][lenth]) and data["STATUS"][lenth] == "NEW" and  "BRAND" in str(data["CURRENT_VALUE"][lenth])  :
                                   data.at[lenth, "BAU_COMMENTS"] = f"OMNI ITEM: Falling in {str(data["CURRENT_VALUE"][lenth])}, Valid"
                              if "MANUFACTURER" in str(data["MAC_DESCRIPTION"][lenth]) and data["STATUS"][lenth] == "NEW" and "MANUFACTURER" in str(data["CURRENT_VALUE"][lenth]) :
                                   data.at[lenth, "BAU_COMMENTS"] = f"OMNI ITEM: Falling in {str(data["CURRENT_VALUE"][lenth])}, Valid" 

                         for lenth in range(len(data["STATUS"])): 
                              if  data["STATUS"][lenth] == "DELETED":
                                   if str(data["IAS_ITEM"][lenth]) == "nan":
                                             data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM : Need to be Check."
                                   elif str(data["IAS_ITEM"][lenth]) == "NULL SALES":
                                             data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Dropped due to Null sales, Valid." 
                                   elif str(data["IAS_ITEM"][lenth]) ==  "NOT IN IAS" :
                                             data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Dropped due to NOT IN IAS, Valid."

                         for lenth in range(len(data["STATUS"])): 

                              if ( str(data["ITEM SCOPE"][lenth]) == "NEWLY ADDED" or "BRAND" in str(data["MAC_DESCRIPTION"][lenth]) or "MANUFACTURER" in str(data["MAC_DESCRIPTION"][lenth])) and data["STATUS"][lenth] == "CHANGED":
                                   if str(data["CURRENT_VALUE"][lenth]) == "nan":
                                        data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Need to check." 
                                   if  str(data["CURRENT_VALUE"][lenth]) == "DETAIL UNKNOWN":
                                        data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM : Need to be Check.Falling into DETAIL UNKNOWN"
                                   if str(data["CURRENT_VALUE"][lenth]) == "PENDING PLACEMENT":
                                        data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM : Need to be Check.Falling into PENDING PLACEMENT" 
                                   if str(data["CURRENT_VALUE"][lenth]) == "DETAIL UNKNOWN" and str(data["CUR_RULE"][lenth]) == "DEFAULT":
                                        data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM : Need to be Check.Falling into DETAIL UNKNOWN and DEFAULT"
                                   if str(data["CURRENT_VALUE"][lenth]) != "nan":
                                        data.at[lenth, "BAU_COMMENTS"] = "OMNI ITEM: Change due to MAC rule change/Value change,Valid."

                                   if "BRAND" in str(data["CURRENT_VALUE"][lenth]) :
                                        data.at[lenth, "BAU_COMMENTS"] = f"OMNI ITEM: Falling in {str(data["CURRENT_VALUE"][lenth])}, Need to Check"
                                   if "MANUFACTURER" in str(data["CURRENT_VALUE"][lenth]):
                                        data.at[lenth, "BAU_COMMENTS"] = f"OMNI ITEM: Falling in {str(data["CURRENT_VALUE"][lenth])}, Need to Check"
                              

                         out = EX_FILE_PATH.get() +"/" +   file_name_to_save.get().upper()+".xlsx"
                         data.to_excel(out, index=False) 
                         label4.configure(text="""
The report has been generated....,
The procedure is finished 100%
                    
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
                         messagebox.showinfo('Success',"The reports have been successfully generated.")
               
# Label
F_location = CTkLabel(master=tab3,text = "Extraction Report",font=('Consolas',13.5),text_color="#DEDEDE")
F_location.place(relx=0.03,rely=0.075)

F_location = CTkLabel(master=tab3,text = "RMS Report",font=('Consolas',13.5),text_color="#DEDEDE")
F_location.place(relx=0.03,rely=0.32)

F_location = CTkLabel(master=tab3,text = "ARS TYPE",font=('Consolas',13.5),text_color="#DEDEDE")
F_location.place(relx=0.03,rely=0.2) 

F_location = CTkLabel(master=tab3,text = "Filename to be Save",font=('Consolas',13.5),text_color="#DEDEDE")
F_location.place(relx=0.03,rely=0.44) 




#Entry
E_report_Entry = CTkEntry(master=tab3,textvariable=EX_File_NAME,width=350,height=25,
                            text_color="#FFCC70",placeholder_text="Extraction Report...",font=("Consolas",14))
E_report_Entry.place(relx=0.27,rely=0.075)

RMS_report_Entry = CTkEntry(master=tab3,textvariable=RMS_File_NAME,width=350,height=25,
                            text_color="#FFCC70",placeholder_text="RMS Report...",font=("Consolas",14))
RMS_report_Entry.place(relx=0.27,rely=0.2)

location_entryss = CTkComboBox(master=tab3,values=["OLD METHOD","NEW METHOD"],font=("Calibri",13),width=140,height=25,text_color="White")
location_entryss.place(relx=0.27,rely=0.32)

location_entrys = CTkEntry(master=tab3,placeholder_text="Save file name...",textvariable=file_name_to_save,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
location_entrys.place(relx=0.27,rely=0.44)

#Buttons
E_report_button = CTkButton(master=tab3,text=" Browse  ",width=10,height=10,command=Extraction_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
E_report_button.place(relx=0.81,rely=0.075)

RMS_report_button = CTkButton(master=tab3,text=" Browse  ",width=10,height=10,command=RMS_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
RMS_report_button.place(relx=0.81,rely=0.44)



generate = CTkButton(master=tab3,text="Run ARS  ",width=20,height=10,command=ars,corner_radius=0,image=CTkImage(light_image=img,dark_image=img))
generate.place(relx=0.6,rely=0.59)

clear = CTkButton(tab3,text=" Clear All ",width=10,height=20,command=clear_all,corner_radius=0)
clear.place(relx=0.762,rely=0.6)

exits = CTkButton(tab3,text="  Exit  ",width=10,height=20,command=exit,corner_radius=0,fg_color="#F95A45")
exits.place(relx=0.88,rely=0.6)


label4 = CTkLabel(master=ARS_frame,text = "",width=0,height=0,text_color="#F5AF61",font=('Consolas',13))
label4.pack(anchor="center",expand=True)

KAD_FILE_PATH = StringVar()
KAD_File_NAME = StringVar()
KAD_FILE = StringVar()


def KAD_FILES ():
     kad_file_path =  filedialog.askopenfilename()
     kadfilename = os.path.split(kad_file_path)
     KAD_FILE_PATH.set(kadfilename[0]) #--- File path only
     KAD_File_NAME.set(kadfilename[1]) #--- File name only
     KAD_FILE.set(kad_file_path) #--- Full file path and file name
     label5.configure(text = "")

def KAD ():
     if KAD_FILE.get() == "":
          
          label5.configure(text="""
                    Insert the files, please...
""",anchor="center")
     else:
          
          try:
               file_name = pd.read_csv(KAD_FILE.get())
          except:
               file_name = pd.read_excel(KAD_FILE.get())

          changed = []
          delete = []
          new = []
          rms = []
          
                         
     
          if omni_type.get() == "OMNISHOPPER":
                    if KAD_BOX.get()== "NEW COUNT":
                         for i in range(file_name.shape[0]):
                              if (str(file_name["ITEM SCOPE"][i]) == "NEWLY ADDED" or "BRAND" in str(file_name["MAC_DESCRIPTION"][i]) or "MANUFACTURER" in str(file_name["MAC_DESCRIPTION"][i])  ):
                                   if str(file_name["STATUS"][i]) == "CHANGED":
                                        changed.append(file_name["ITM_ID"][i])
                                   elif str(file_name["STATUS"][i]) == "NEW":
                                        new.append(file_name["ITM_ID"][i])  
                                   elif str(file_name["ITEM SCOPE"][i]) !="NEWLY ADDED" :
                                        rms.append(file_name["ITM_ID"][i]) 
                              if str(file_name["STATUS"][i]) == "DELETED":
                                        delete.append(file_name["ITM_ID"][i]) 
                    else:
                         for i in range(file_name.shape[0]):
                              if str(file_name["EAN_CODE"][i]) == "nan": 
                                   if str(file_name["STATUS"][i]) == "CHANGED":
                                        changed.append(file_name["ITM_ID"][i])
                                   elif str(file_name["STATUS"][i]) == "NEW":
                                        new.append(file_name["ITM_ID"][i])
                              if str(file_name["STATUS"][i]) == "DELETED":
                                        delete.append(file_name["ITM_ID"][i])
                              elif str(file_name["EAN_CODE"][i]) !="nan" :
                                        rms.append(file_name["ITM_ID"][i]) 

          else:
               for i in range(file_name.shape[0]):
                    if str(file_name["STATUS"][i]) == "CHANGED":
                         changed.append(file_name["ITM_ID"][i])
                    elif str(file_name["STATUS"][i]) == "DELETED":
                         delete.append(file_name["ITM_ID"][i])
                    elif str(file_name["STATUS"][i]) == "NEW":
                         new.append(file_name["ITM_ID"][i])
                         
          label5.configure(text=f"""NEW :  {len(set(new))}  ||  CHANGED  :  {len(set(changed))}  ||  DELETED :  {len(set(delete))}  
                              
RMS :  {len(set(rms))} | TOTAL :  { int(len(changed) + len(delete) + len(new) + len(rms))} | OVERALL :  {file_name.shape[0]}  
                    
or any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
     
def KAD_clear_all():
     if KAD_File_NAME.get() != "" :
          KAD_File_NAME.set("")
          label5.configure(text="""
                              Cleared!"""
,anchor="center")

# Label
KAD_location = CTkLabel(master=tab4,text = "FILE",font=('Consolas',13.5),text_color="#DEDEDE")
KAD_location.place(relx=0.03,rely=0.1) 

omni_type_VALUE = CTkLabel(master=tab4,text = "OMNI TYPE ",font=('Consolas',13.5),text_color="#DEDEDE")
omni_type_VALUE.place(relx=0.03,rely=0.26) 

omni_type = CTkComboBox(master=tab4,values=["OMNISALES","OMNISHOPPER"],font=("Calibri",13),width=140,height=25,text_color="White")
omni_type.place(relx=0.26,rely=0.26)

KAD_TYPE = CTkLabel(master=tab4,text = "COUNT TYPE ",font=('Consolas',13.5),text_color="#DEDEDE")
KAD_TYPE.place(relx=0.03,rely=0.40) 

KAD_BOX = CTkComboBox(master=tab4,values=["OLD COUNT","NEW COUNT"],font=("Calibri",13),width=140,height=25,text_color="White")
KAD_BOX.place(relx=0.26,rely=0.40)

kad_report_Entry = CTkEntry(master=tab4,placeholder_text="Changes Report...",textvariable=KAD_File_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
kad_report_Entry.place(relx=0.26,rely=0.1)



#Buttons
kad_report_button = CTkButton(master=tab4,text=" Browse  ",width=10,height=10,command=KAD_FILES,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
kad_report_button.place(relx=0.81,rely=0.1)

generate_KAD = CTkButton(master=tab4,text="Run KAD  ",width=20,height=10,command=KAD,corner_radius=0,image=CTkImage(light_image=img,dark_image=img))
generate_KAD.place(relx=0.6,rely=0.39)

clear = CTkButton(tab4,text=" Clear All ",width=10,height=20,command=KAD_clear_all,corner_radius=0)
clear.place(relx=0.762,rely=0.4)

exits = CTkButton(tab4,text="  Exit  ",width=10,height=20,command=exit,corner_radius=0,fg_color="#F95A45")
exits.place(relx=0.88,rely=0.4)

label5 = CTkLabel(master=frame4,text = "",width=0,height=0,text_color="#F5AF61",font=('Consolas',15))
label5.pack(anchor="center",expand=True)


tab5 = CTkLabel(tabview.tab(" FILE SPLIT "),text="",width=1400,height=1000)
tab5.pack(padx=20,pady=20)

SPLIT_FRAME = CTkFrame(master=tab5,width=450,height=45,border_width=.6)
SPLIT_FRAME.place(relx=0.18,rely=.75)

SP_FILE_PATH =StringVar()
SP_File_NAME =StringVar()
SP_FILE = StringVar()
ROWS_COUNT = StringVar()

def load_split_file():
     global SPLIT_DB
     File_filepath = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")]) 
     sp_filename = os.path.split(File_filepath)
     SP_FILE_PATH.set(sp_filename[0]) #--- File path only
     SP_File_NAME.set(sp_filename[1]) #--- File name only
     SP_FILE.set(File_filepath) #--- Full file path and file name
     with zipfile.ZipFile(SP_FILE.get(), 'r') as zip_ref:
          with zip_ref.open(zip_ref.namelist()[0]) as file:
               try:
                    SPLIT_DB = pd.read_csv(file)
               except:
                    SPLIT_DB = pd.read_excel(file)
               label7.configure(text=f"""
               Total Number of Rows - {(SPLIT_DB.shape[0])}"""
,anchor="center")

def header():
     global SPLIT_DB
     all_headers = SPLIT_DB.iloc[0].tolist()
     headers = [h for h in all_headers if  str(h).endswith(']') and h != 0]
     SPLIT_DB = SPLIT_DB.iloc[1:, :]
     columns = SPLIT_DB.columns.tolist()

     new_df_columns = []
     for col in columns:
          if col.startswith("char"):
               try:
                    new_df_columns.append(headers.pop(0))  # Use new columns one by one
               except:
                    new_df_columns.append("")
          else:
               if col != '':
                    new_df_columns.append(col)
     SPLIT_DB.columns = new_df_columns
     SPLIT_DB = SPLIT_DB.iloc[1:].reset_index(drop=True)



def split_file():
     global SPLIT_DB
     if Column_yes.get() == "NO" and  SPLIT_AUTO_YES.get() =="AUTOMATIC":
          rows = 749999
          num_of_files = (SPLIT_DB.shape[0] - 1) // rows + 1
          if SPLIT_DB.shape[0] > 750000:
               for i in range(num_of_files):
                    starting_index = i * rows
                    ending_index = min((i + 1) * rows, SPLIT_DB.shape[0])               
                    split_file = SPLIT_DB.iloc[starting_index:ending_index]
                    output_file = os.path.join(SP_FILE_PATH.get(),f"SPLIT FILE {i+1}.csv")
                    split_file.to_csv(output_file,index=False)
                    label7.configure(text="""
The report has been generated....,
The procedure is finished 100%
               
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
               messagebox.showinfo('Success',"The reports have been successfully generated.")

               
          else:
               
               label7.configure(
text=f"""
     The file has only {(SPLIT_DB.shape[0])} rows, 
     while 800000 rows are required for auto splitting. 
     You can try Manual mode instead."""
,anchor="center")
               

     elif Column_yes.get() == "NO" and  SPLIT_AUTO_YES.get() =="MANUAL":
          rows = int(ROWS_COUNT.get())
          num_of_files = (SPLIT_DB.shape[0] - 1) // rows + 1
          for i in range(num_of_files):
               starting_index = i * rows
               ending_index = min((i + 1) * rows, SPLIT_DB.shape[0])               
               split_file = SPLIT_DB.iloc[starting_index:ending_index]
               output_file = os.path.join(SP_FILE_PATH.get(),f"SPLIT FILE {i+1}.csv")
               split_file.to_csv(output_file,index=False)  
               label7.configure(text="""
The report has been generated....,
The procedure is finished 100%
               
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
          messagebox.showinfo('Success',"The reports have been successfully generated.")

     elif Column_yes.get() == "YES" and  SPLIT_AUTO_YES.get() =="MANUAL":
          rows = int(ROWS_COUNT.get())
          num_of_files = (SPLIT_DB.shape[0] - 1) // rows + 1
          
          header()
          for i in range(num_of_files):
               starting_index = i * rows
               ending_index = min((i + 1) * rows, SPLIT_DB.shape[0])               
               split_file = SPLIT_DB.iloc[starting_index:ending_index]
               output_file = os.path.join(SP_FILE_PATH.get(),f"SPLIT FILE {i+1}.csv")
               split_file.to_csv(output_file,index=False)
               label7.configure(text="""
The report has been generated....,
The procedure is finished 100%
               
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
          messagebox.showinfo('Success',"The reports have been successfully generated.")

     elif Column_yes.get() == "YES" and  SPLIT_AUTO_YES.get() =="AUTOMATIC":
          rows = 749999
          num_of_files = (SPLIT_DB.shape[0] - 1) // rows + 1
          if SPLIT_DB.shape[0] > 750000:
               header()
               for i in range(num_of_files):
                    starting_index = i * rows
                    ending_index = min((i + 1) * rows, SPLIT_DB.shape[0])               
                    split_file = SPLIT_DB.iloc[starting_index:ending_index]
                    output_file = os.path.join(SP_FILE_PATH.get(),f"SPLIT FILE {i+1}.csv")
                    split_file.to_csv(output_file,index=False)
                    label7.configure(text="""
The report has been generated....,
The procedure is finished 100%
               
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
               messagebox.showinfo('Success',"The reports have been successfully generated.")
               
          else:
               label7.configure(text=f"""
     The file has only {(SPLIT_DB.shape[0])} rows, 
     while 800000 rows are required for auto splitting. 
     You can try MANUAL mode instead."""
,anchor="center")
               

          
           

          

File = CTkLabel(master=tab5,text = "FILE",font=('Consolas',13.5),text_color="#DEDEDE")
File.place(relx=0.04,rely=0.06)

File_entry = CTkEntry(master=tab5,textvariable=SP_File_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
File_entry.place(relx=0.282,rely=0.06)

Load_file = CTkButton(master=tab5,text=" Load  ",width=10,height=10,command=load_split_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
Load_file.place(relx=0.81,rely=0.06)




Columns_type = CTkLabel(master=tab5,text = "COLUMN FORMATING",font=('Consolas',13.5),text_color="#DEDEDE")
Columns_type.place(relx=0.04,rely=0.21)

Column_yes = CTkComboBox(master=tab5,values=["YES","NO"],font=("Calibri",13),width=140,height=25,text_color="White")
Column_yes.place(relx=0.28,rely=0.21)



SPLIT_AUTO = CTkLabel(master=tab5,text = "MODE",font=('Consolas',13.5),text_color="#DEDEDE")
SPLIT_AUTO.place(relx=0.04,rely=0.35)

SPLIT_AUTO_YES = CTkComboBox(master=tab5,values=["AUTOMATIC","MANUAL"],font=("Calibri",13),width=140,height=25,text_color="White")
SPLIT_AUTO_YES.place(relx=0.28,rely=0.35)


ROW_COUNT = CTkLabel(master=tab5,text = "ENTER THE NO.OF ROWS",font=('Consolas',13.5),text_color="#DEDEDE")
ROW_COUNT.place(relx=0.04,rely=0.49)

ROWWS = CTkEntry(master=tab5,textvariable=ROWS_COUNT,width=140,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
ROWWS.place(relx=0.28,rely=0.49)



Split_button = CTkButton(master=tab5,text="   SPLIT    ",width=20,height=10,command=split_file,corner_radius=0,image=CTkImage(light_image=img,dark_image=img))
Split_button.place(relx=0.78,rely=0.49)

label7 = CTkLabel(master=SPLIT_FRAME,text = "",width=0,height=0,text_color="#F5AF61",font=('Consolas',15))
label7.pack(anchor="center",expand=True)



# IVD

tab6 = CTkLabel(tabview.tab(" IVD "),text="",width=1400,height=1000)
tab6.pack(padx=20,pady=20)

IVD_FRAME = CTkFrame(master=tab6,width=450,height=45,border_width=.6)
IVD_FRAME.place(relx=0.18,rely=.75)
label8 = CTkLabel(master=IVD_FRAME,text = "",width=0,height=0,text_color="#F5AF61",font=('Consolas',15))
label8.pack(anchor="center",expand=True)

IVD_FILE_PATH =StringVar()
IVD_File_NAME =StringVar()
IVD_FILE = StringVar()


IV_FILE_PATH =StringVar()
IV_File_NAME =StringVar()
IV_FILE_4101 = StringVar()


Exctra_column = StringVar()




def load_IVD_file():
     global SPLIT_DB
     ivd_filepath = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")]) 
     ivd_filename = os.path.split(ivd_filepath)
     IVD_FILE_PATH.set(ivd_filename[0]) #--- File path only
     IVD_File_NAME.set(ivd_filename[1]) #--- File name only
     IVD_FILE.set(ivd_filepath) #--- Full file path and file name

def load_IV_file():
     global SPLIT_DB
     iv_filepath = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")]) 
     iv_filename = os.path.split(iv_filepath)
     IV_FILE_PATH.set(iv_filepath[0]) #--- File path only
     IV_File_NAME.set(iv_filename[1]) #--- File name only
     IV_FILE_4101.set(iv_filepath) #--- Full file path and file name


def IVD():
          
          with zipfile.ZipFile(IVD_FILE.get(), 'r') as zip_ref:
               with zip_ref.open(zip_ref.namelist()[0]) as file:
                    try:
                         IVD_DB = pd.read_csv(file,index_col=2)
                    except:
                         IVD_DB = pd.read_excel(file,index_col=2)

          with zipfile.ZipFile(IV_FILE_4101.get(), 'r') as zip_ref:
               with zip_ref.open(zip_ref.namelist()[0]) as file:
                    try:
                         IV_DB = pd.read_csv(file,index_col=None)
                    except:
                         IV_DB = pd.read_excel(file)
          

                    all_IV_headers = IV_DB.iloc[0].tolist()
                    headersS = [h for h in all_IV_headers if  str(h).endswith(']') and h != 0]
                    IV_DB = IV_DB.iloc[1:, :]
                    columnss = IV_DB.columns.tolist()

                    new_df_columns = []

                    for col in columnss:
                              if col.startswith("char"):
                                   try:
                                        new_df_columns.append(headersS.pop(0))  # Use new columns one by one
                                   except:
                                        new_df_columns.append("")
                              else:
                                   if col != '':
                                        new_df_columns.append(col)
                    IV_DB.columns = new_df_columns
                    IV_DB = IV_DB.iloc[1:].reset_index(drop=True)
                    

          

          IVD_DB = IVD_DB.drop(['SEQN', 'CHAR_ID','CHAR_TYPE'], axis=1)
          IV_DB.rename(columns = {"nan_key" : "ITEM_ID"}, inplace=True)

          orginal_ivd = IVD_DB.T.reset_index(drop=True)

          iv = orginal_ivd

          needed_columns = ["#US LOC OMNI DEPARTMENT MASTER ",
                    "#US LOC OMNI SUPER CATEGORY MASTER ",
                    "#US LOC OMNI CATEGORY MASTER ",
                    "#US LOC OMNI SUB CATEGORY MASTER ",
                    "#US LOC OMNI SEGMENT MASTER ",
                    "#US LOC BRAND OWNER ",
                    "#US LOC DERIVED BRAND HIGH ",
                    "#US LOC DERIVED BRAND OWNER HIGH ",
                    "#US LOC BRAND ",
                    "ITEM_ID",'MODULE '
                    ]

          
          iv['ITEM_ID'] = iv['ITEM_ID'].astype(int)
          
          for i in iv.columns:
               if str(i.split("[")[0]) not in list(needed_columns):
                    iv.drop(i, axis=1, inplace=True)

          merge_data = pd.merge(iv, IV_DB, on="ITEM_ID", how="left")

          for ii in merge_data.columns:
               try:
                    if ii.split("[")[0] not in needed_columns :
                         merge_data.drop(ii,axis=1,inplace=True)
               except:
                    pass
          #sorted_iv.to_csv(r"C:\Users\kura2020\Desktop\MAINTENACE\WK 47\FLOUR\CC\beforemerge.csv")
          sorted_iv = merge_data.sort_index(axis=1, ascending=True, inplace=False)
          sorted_iv.drop("ITEM_ID", axis = 1,inplace=True)
          
          sorted_iv.insert(0,"ITEM_ID",orginal_ivd["ITEM_ID"])

          IVD_data = sorted_iv
     
          columns = IVD_data.columns
          for col in columns:
               if col.endswith("_x"):
                    # Get the base column name without the `_x` suffix
                    base_name = col[:-2]
                    corresponding_y = base_name + "_y"
                    
                    # Check if the corresponding `_y` column exists
                    if corresponding_y in columns:
                         # Create a comparison column
                         comparison_column = 'STATEMENT OF ' + base_name 
                         IVD_data[comparison_column] = IVD_data[col] == IVD_data[corresponding_y]
                         
                         # Convert Boolean to TRUE/FALSE
                         IVD_data[comparison_column] = IVD_data[comparison_column].apply(lambda x: 'TRUE' if x else 'FALSE')
                         
                         # Insert the comparison column right after the `_y` column
                         insert_position = IVD_data.columns.get_loc(corresponding_y) + 1
                         IVD_data.insert(insert_position, comparison_column, IVD_data.pop(comparison_column))

          for cc in IVD_data.columns:
               if cc.endswith("_x") :
                    IVD_data.rename(columns = {str(cc) : str(cc.split("_")[0] + " CURRENT WEEK")}, inplace=True)
               elif cc.endswith("_y") :
                    IVD_data.rename(columns = {str(cc): str(cc.split("_")[0])}, inplace=True)
          label8.configure(text="""
     The report has been generated....,
     The procedure is finished 100%
                    
     For any technical queries, reach out to Ram.kumar@nielseniq.com....
     """,anchor="center")
          messagebox.showinfo('Success',"The reports have been successfully generated.")

          output_file_ivd = os.path.join(IVD_FILE_PATH.get(),"IVD OUTPUT FILE.csv")
          IVD_data.to_csv(output_file_ivd,index=False)

# IVD Report 1st
IVD_File = CTkLabel(master=tab6,text = "IVD Report",font=('Consolas',13.5),text_color="#DEDEDE")
IVD_File.place(relx=0.04,rely=0.06)

ivd_entry = CTkEntry(master=tab6,textvariable=IVD_File_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
ivd_entry.place(relx=0.282,rely=0.06)

Load_IVD_file = CTkButton(master=tab6,text=" Load  ",width=10,height=10,command=load_IVD_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
Load_IVD_file.place(relx=0.81,rely=0.06)



# IV Report 1st
IV_File = CTkLabel(master=tab6,text = "IV Report",font=('Consolas',13.5),text_color="#DEDEDE")
IV_File.place(relx=0.04,rely=0.21)

iv_entry = CTkEntry(master=tab6,textvariable=IV_File_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
iv_entry.place(relx=0.28,rely=0.21)

Load_IV_file = CTkButton(master=tab6,text=" Load  ",width=10,height=10,command=load_IV_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
Load_IV_file.place(relx=0.81,rely=0.21)


"""IVD_add = CTkLabel(master=tab6,text = "ADD EXCTRA COLUMN",font=('Consolas',13.5),text_color="#DEDEDE")
IVD_add.place(relx=0.04,rely=0.35)

Added_column = CTkEntry(master=tab6,textvariable=Exctra_column,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
Added_column.place(relx=0.28,rely=0.35)"""


IVD_4101 = CTkLabel(master=tab6,text = "REPORT",font=('Consolas',13.5),text_color="#DEDEDE")
IVD_4101.place(relx=0.04,rely=0.35)

IVD_4101_AUTO_YES = CTkComboBox(master=tab6,values=["IVD","Blank"],font=("Calibri",13),width=140,height=25,text_color="White")
IVD_4101_AUTO_YES.place(relx=0.28,rely=0.35)




IVD_button = CTkButton(master=tab6,text="   RUN    ",width=20,height=10,command=IVD,corner_radius=0,image=CTkImage(light_image=img,dark_image=img))
IVD_button.place(relx=0.78,rely=0.49)



############################################   MAP VALUE ################################

tab7 = CTkLabel(tabview.tab(" MAP VALUE "),text="",width=1400,height=1000)
tab7.pack(padx=20,pady=20)

MAP_FRAME = CTkFrame(master=tab7,width=450,height=45,border_width=.6)
MAP_FRAME.place(relx=0.18,rely=.75)



map_FILE_PATH =StringVar()
map_File_NAME =StringVar()
map_FILE_4101 = StringVar()

MAP1_FILE_PATH =StringVar()
MAP1_File_NAME =StringVar()
map1_FILE_4101 = StringVar()

def load_map_file(): #---- PRE WEEK

     map_filepath = filedialog.askopenfilename() 
     map_filename = os.path.split(map_filepath)
     map_FILE_PATH.set(map_filename[0]) #--- File path only
     map_File_NAME.set(map_filename[1]) #--- File name only
     map_FILE_4101.set(map_filepath) #--- Full file path and file name

def load_map1_file(): # - -- CURRENT WEEK

     map1_filepath = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")]) 
     map1_filename = os.path.split(map1_filepath)
     MAP1_FILE_PATH.set(map1_filename[0]) #--- File path only
     MAP1_File_NAME.set(map1_filename[1]) #--- File name only
     map1_FILE_4101.set(map1_filepath) #--- Full file path and file name

def map_values():
     with zipfile.ZipFile(map1_FILE_4101.get(), 'r') as zip_ref:
               with zip_ref.open(zip_ref.namelist()[0]) as file:
                    try:
                         map_value = pd.read_csv(file,skiprows=1)
                    except:
                         map_value = pd.read_excel(file,skiprows=1)
     try:
          map_value_pre = pd.read_excel(map_FILE_4101.get())
     except:
          map_value_pre = pd.read_csv(map_FILE_4101.get())

     concat = []
     file_save = map1_FILE_4101.get()
     for m in range(map_value.shape[0]):
          con = str(map_value["VALUE"][m]) + str(map_value["VALUE ID"][m])
          concat.append(str(con))
     
     map_value.insert(5,"CONCAT",concat)
     
     VLOOKUP = []
     for v in range(map_value.shape[0]):
          if str(map_value["VALUE ID"][v]) in str(list(map_value_pre["VALUE ID"])):
               spec_value = map_value["VALUE ID"][v]
               name  = map_value_pre.loc[map_value_pre['VALUE ID'] == spec_value, 'VLOOKUP'].values[0]
               VLOOKUP.append(name)
          else:
               VLOOKUP.append("#NA")

     map_value.insert(6,"VLOOKUP",VLOOKUP)

     for drop_col in map_value.columns:
          if "Unnamed" in drop_col :
               print(drop_col)
               map_value.drop(drop_col, axis= 1 , inplace = True)

     output_file_map = os.path.join(MAP1_FILE_PATH.get(),"MAP VALUE Report Output.xlsx")
     map_value.to_excel(output_file_map,index=False)
     label9.configure(text="""
The report has been generated....,
The procedure is finished 100%
               
For any technical queries, reach out to Ram.kumar@nielseniq.com....
""",anchor="center")
     messagebox.showinfo('Success',"The reports have been successfully generated.")



"""                                  ###################### FRONT END FOR MAP VALUE #####################                  """

map_File = CTkLabel(master=tab7,text = "Current week Report",font=('Consolas',13.5),text_color="#DEDEDE")
map_File.place(relx=0.04,rely=0.06)

map_entry = CTkEntry(master=tab7,textvariable=MAP1_File_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
map_entry.place(relx=0.282,rely=0.06)

Load_map_file = CTkButton(master=tab7,text=" Load  ",width=10,height=10,command=load_map1_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
Load_map_file.place(relx=0.81,rely=0.06)



map_pre_File = CTkLabel(master=tab7,text = "Pre Week Report",font=('Consolas',13.5),text_color="#DEDEDE")
map_pre_File.place(relx=0.04,rely=0.21)

map_pre_entry = CTkEntry(master=tab7,textvariable=map_File_NAME,width=350,height=25,
                            text_color="#FFCC70",font=("Consolas",14))
map_pre_entry.place(relx=0.28,rely=0.21)

Load_map_pre_file = CTkButton(master=tab7,text=" Load  ",width=10,height=10,command=load_map_file,image=CTkImage(light_image=img2,dark_image=img2,size=(20,20)),corner_radius=45,fg_color="green")
Load_map_pre_file.place(relx=0.81,rely=0.21)

MAP_button = CTkButton(master=tab7,text="   RUN    ",width=20,height=10,command=map_values,corner_radius=0,image=CTkImage(light_image=img,dark_image=img))
MAP_button.place(relx=0.78,rely=0.49)

label9 = CTkLabel(master=MAP_FRAME,text = "",width=0,height=0,text_color="#F5AF61",font=('Consolas',15))
label9.pack(anchor="center",expand=True)

app.mainloop()

# 0008001009009