

import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import PySimpleGUI as sg
import time
import xlwings as xw
import glob
import pyautogui as pg
import subprocess
import traceback
import warnings
import re
import numpy as np
import pyodbc
import pyxlsb

warnings.filterwarnings("ignore")
sg.theme('DarkBlue')

layout = [
        
        [sg.Text("Choose Your Client",s=45,justification="l"),sg.Listbox(values=['Magnit', 'Dixy', 'X5', 'RW', 'Bristol'], size=(45, 5), select_mode='single', key='-CLIENT-')],
        

        [sg.Submit( )]
    ]

window = sg.Window('Top 5 APP', layout)
event, values = window.read()
client=str(values['-CLIENT-'][0])
    
window.close()

if client=="Magnit":
    sg.theme('DarkBlue')

    layout = [
        [sg.T("Input File with all URLs:", s=45,justification="l"), sg.I(key="-IN-", s=45), sg.FileBrowse(file_types=(("Excel Files", "*.xls*")))],
        [sg.Text("Choose data modification type",s=45,justification="l"),sg.Listbox(values=['Monthly', 'Weekly'], size=(45, 2), select_mode='single', key='-DESTINATION-')],
        [sg.Text("Input week or month num (Will be searched in files)",s=45,justification="l"),sg.InputText( key='-DATE-')],
        
        [sg.T("Input Storage Folder:", s=45,justification="l"), sg.I(key="-FOLDER-"), sg.FolderBrowse()],


        [sg.Submit( )]
    ]

    window = sg.Window('Программа обработки продаж Магнит', layout)

    event, values = window.read()

    user_path=str(values['-IN-'])
    selection=str(values['-DESTINATION-'][0])
    period=str(values['-DATE-'])
    result_file=str(values['-FOLDER-'])
    window.close()

    user_path=user_path.replace("\\","\\\\")
    result_file=result_file.replace("\\","\\\\")

    paths_data=pd.read_excel(user_path)
    paths_data['Type']=paths_data['Type'].str.lower()

    start_time=time.time()
    #---------------------------------------------------------------------------------------
    monthly_f_path_i=paths_data[paths_data.Type.str.contains("monthly folder")]['Path'].index[0]
    weekly_f_path_i=paths_data[paths_data.Type.str.contains("weekly folder")]['Path'].index[0]
    geo_path_i=paths_data[paths_data.Type.str.contains("geo")]['Path'].index[0]
    result_f_path_i=paths_data[paths_data.Type.str.contains("result")]['Path'].index[0]
    pbi_path_i=paths_data[paths_data.Type.str.contains("pbi")]['Path'].index[0]
    in_m_st_i=paths_data[paths_data.Type.str.contains("monthly stock")]['Path'].index[0]
    in_m_rrp_i=paths_data[paths_data.Type.str.contains("monthly sales rrp")]['Path'].index[0]
    in_m_rmc_i=paths_data[paths_data.Type.str.contains("monthly sales rmc")]['Path'].index[0]
    in_m_dc_i=paths_data[paths_data.Type.str.contains("monthly dc")]['Path'].index[0]
    in_w_sales_i=paths_data[paths_data.Type.str.contains("weekly sales")]['Path'].index[0]
    in_w_stock_i=paths_data[paths_data.Type.str.contains("weekly stock")]['Path'].index[0]
    in_w_dc_i=paths_data[paths_data.Type.str.contains("weekly dc")]['Path'].index[0]
    #---------------------------------------------------------------------------------------
    monthly_f_path=paths_data[paths_data.Type.str.contains("monthly folder")]['Path'][monthly_f_path_i]

    weekly_f_path=paths_data[paths_data.Type.str.contains("weekly folder")]['Path'][weekly_f_path_i]

    result_file_geo=paths_data[paths_data.Type.str.contains("geo")]['Path'][geo_path_i]

    result_file=paths_data[paths_data.Type.str.contains("result")]['Path'][result_f_path_i]

    pbi_path=paths_data[paths_data.Type.str.contains("pbi")]['Path'][pbi_path_i]

    in_m_st=paths_data[paths_data.Type.str.contains("monthly stock")]['Path'][in_m_st_i]

    in_m_rrp=paths_data[paths_data.Type.str.contains("monthly sales rrp")]['Path'][in_m_rrp_i]

    in_m_rmc=paths_data[paths_data.Type.str.contains("monthly sales rmc")]['Path'][in_m_rmc_i]

    in_m_dc=paths_data[paths_data.Type.str.contains("monthly dc")]['Path'][in_m_dc_i]

    in_w_sales=paths_data[paths_data.Type.str.contains("weekly sales")]['Path'][in_w_sales_i]

    in_w_stock=paths_data[paths_data.Type.str.contains("weekly stock")]['Path'][in_w_stock_i]

    in_w_dc=paths_data[paths_data.Type.str.contains("weekly dc")]['Path'][in_w_dc_i]
    in_m_st=in_m_st.replace("\u202a","")
    in_m_rrp=in_m_rrp.replace("\u202a","")
    in_m_rmc=in_m_rmc.replace("\u202a","")
    in_m_dc=in_m_dc.replace("\u202a","")
    in_w_stock=in_w_stock.replace("\u202a","")
    in_w_sales=in_w_sales.replace("\u202a","")
    in_w_dc=in_w_dc.replace("\u202a","")
    pbi_file='cmd /K ' +'"cd ' +pbi_path +' & magnit.pbix"'

    if selection=="Monthly":
        etl_file=monthly_f_path
    else:
        etl_file=weekly_f_path

    print("Extracting Data")
    print("----")

    all_files=glob.glob(etl_file+"/*.xls*")
    files_list=[]
    results_list=[]

    for filename in all_files:
        filename=filename.lower()

        if selection=="Monthly":
            #MONTHLY------------------------------------------------------------------------------    
                if "продажи rmc" in filename and period in filename:
                    print("Нашли ",filename)
                    try:
                        #--------------------------
                        print("Начали обработку")
                        print("----")
                        #--------------------------
                        sales=pd.read_excel(filename,index_col=None,header=0)
                        header_row=sales.index[sales.iloc[:,0] == 'Магазин'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        sales=sales[sales['Магазин'].notnull()]
                        print("Прочитали содержимое файла")
                        print("----")
                        #------------------------------------------------------------------------

                        geo=sales[['Магазин','Формат','Филиал','РЦ (ОС)']]
                        geo=geo.rename(columns={"Магазин": "Наименование ТТ", "РЦ (ОС)":"РЦ"})
                        geo=geo.drop_duplicates()
                        geo=geo[geo['Наименование ТТ']!="Grand Total"]
                        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
                        geo=geo[geo['Наименование ТТ']!="Общий итог"]
                        sales=sales.drop(['Формат','Филиал','РЦ (ОС)'],axis=1)





                        #-----------------------------------------------------------------------
                        sales=pd.melt(sales,id_vars='Магазин',var_name="SKU", value_name='sales')
                        sales=sales[sales.sales.notnull()]
                        sales=sales[sales['Магазин']!="Grand Total"]
                        sales=sales[sales['Магазин']!="Общий Итог"]
                        sales=sales[sales['Магазин']!="Общий итог"]
                        sales['SKU']=sales['SKU'].str.lower()
                        filename=filename.replace(".xlsm","")
                        sales["DATE"]="01/"+filename[-4:-2]+"/2022"

                        #--------------------------
                        print("Транспонировали и присвоили дату")
                        print("----")
                        print(sales['DATE'].unique()[0], " - Присвоенная дата")
                        #--------------------------
                        print("----")
                        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
                        sales['mrp2']=sales['mrp2'].str.replace(" ","")
                        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
                        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
                        sales['mrp2']=sales.mrp2.fillna(0)
                        sales=sales.rename(columns={"mrp2": "MRP"})
                        sales=sales.drop(['mrp1'],axis=1)
                        #--------------------------
                        print("Добавили МРЦ")
                        print("----")
                        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
                        print("Обработали ",filename)
                        print("----")
                        print("так выглядит обработанный массив")
                        print("----")
                        print(sales.head(5).to_string(index=False))
                        print("----")
                        #--------------------------
                        print(filename[-4:-2]," Полученная сумма продаж без учета блоков: ",sales.sales.sum())
                        m_rmc_result=filename+" Полученная сумма продаж без учета блоков: "+str(sales.sales.sum())
                        results_list.append(m_rmc_result)
                        #--------------------------
                    except Exception as e:
                        m_rmc_result=filename+" НЕ ПОЛУЧИЛОСЬ"
                        results_list.append(m_rmc_result)
                        traceback.print_exception(e)
                        wait_for_it = input('Press enter to close the terminal window')
                        print(filename,"НЕ ПОЛУЧИЛОСЬ")
                        
                        pass
                        
                    frame=sales 
                    frame = frame.astype({"Магазин": str})
                    wb = xw.Book(result_file_geo)
                    ws = wb.sheets["GEO Monthly RMC"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "Monthly_RMC"

                    #Append Non-Mapped GEO
                    ws = wb.sheets["GEO"]
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    rows=geo_init.shape[0]
                    ws = wb.sheets["Monthly RMC"]
                    wb.api.RefreshAll()
                    time.sleep(35)
                    df = ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    ws = wb.sheets["GEO"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = df
                    wb.save()
                    wb.close()
                    #DONE

                    #--------------------------    
                    print("Теперь соединим файл с исходным")
                    print("----")
                    #current table
                    current_df=pd.read_csv(in_m_rmc , delimiter="\t")
                    current_df = current_df.astype({"Магазин": str})
                    new_total=pd.concat([frame,current_df])
                    new_total=new_total[new_total['SKU'].notnull()]
                    print("Прочитали Файл")
                    print("----")
                    print("Добавили файл к исходному")
                    #--------------------------
                    new_total.to_csv(result_file+"\\Current_RMC.txt", index=None, sep='\t', mode='w+')
                    print("Сохранили файл")
                    print("----")
                    print("Можно закрывать программу")
                
                
                
                elif "продажи rrp" in filename and period in filename:

                    print("Нашли ",filename)

                    try:
                        #--------------------------
                        print("Начали обработку")
                        print("----")
                        #--------------------------
                        sales=pd.read_excel(filename,index_col=None,header=0)
                        
                        header_row=sales.index[sales.iloc[:,0] == 'Магазин'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        print("Прочитали содержимое файла")
                        print("----")
                        #------------------------------------------------------------------------

                        geo=sales[['Магазин','Формат','Филиал','РЦ (ОС)']]
                        geo=geo.rename(columns={"Магазин": "Наименование ТТ", "РЦ (ОС)":"РЦ"})
                        geo=geo.drop_duplicates()
                        geo=geo[geo['Наименование ТТ']!="Grand Total"]
                        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
                        geo=geo[geo['Наименование ТТ']!="Общий итог"]
                        

                        sales=sales.drop(['Формат','Филиал','РЦ (ОС)'],axis=1)





                        #-----------------------------------------------------------------------
                        sales=pd.melt(sales,id_vars='Магазин',var_name="SKU", value_name='sales')
                        sales=sales[sales.sales.notnull()]
                        sales=sales[sales['Магазин']!="Grand Total"]
                        sales=sales[sales['Магазин']!="Общий Итог"]
                        sales=sales[sales['Магазин']!="Общий итог"]
                        sales['SKU']=sales['SKU'].str.lower()
                        filename=filename.replace(".xlsm","")
                        sales["DATE"]="01/"+filename[-4:-2]+"/2022"
                        #--------------------------
                        print("Транспонировали и присвоили дату")
                        print("----")
                        print(sales['DATE'].unique()[0], " - Присвоенная дата")
                        print("----")
                        #--------------------------

                        print("Обработали ",filename)
                        print("----")
                        print("так выглядит обработанный массив")
                        print("----")
                        print(sales.head(5).to_string(index=False))
                        print("----")
                        print(filename[-4:-2],"  Полученная сумма продаж без учета блоков: ",sales.sales.sum())
                        m_rrp_result=filename+" Полученная сумма продаж без учета блоков: "+str(sales.sales.sum())
                        results_list.append(m_rrp_result)
                    except Exception as e:
                        m_rrp_result=filename+" НЕ ПОЛУЧИЛОСЬ"
                        results_list.append(m_rrp_result)
                        traceback.print_exception(e)
                        wait_for_it = input('Press enter to close the terminal window')
                        print(filename,"НЕ ПОЛУЧИЛОСЬ")
                        
                        pass

                    frame=sales
                    frame = frame.astype({"Магазин": str})
                    #----------------------------------------------------------------------------------
                    wb = xw.Book(result_file_geo)
                    ws = wb.sheets["GEO Monthly RRP"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "Monthly_RRP"

                    #Append Non-Mapped GEO
                    ws = wb.sheets["GEO"]
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    rows=geo_init.shape[0]
                    ws = wb.sheets["Monthly RRP"]
                    wb.api.RefreshAll()
                    time.sleep(35)
                    df = ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    ws = wb.sheets["GEO"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = df
                    wb.save()
                    wb.close()
                    #DONE

                    #--------------------------   
                    #--------------------------    
                    print("Теперь соединим файл с исходным")
                    #current table
                    current_df=pd.read_csv(in_m_rrp , delimiter="\t")
                    current_df = current_df.astype({"Магазин": str})
                    print("Прочитали Файл")
                    new_total=pd.concat([frame,current_df])
                    new_total=new_total[new_total['SKU'].notnull()]
                    print("Добавили файл к исходному")
                    #--------------------------
                    new_total.to_csv(result_file+"\\Current_RRP.txt", index=None, sep='\t', mode='w+')
                    print("Сохранили файл")
                    print("----")
                    print("Можно закрывать файл")      
                elif "тт" in filename and period in filename:  
                    print("Нашли ",filename) 
                    try:
                        #--------------------------
                        print("Начали обработку")
                        print("----")
                        #--------------------------
                        sales=pd.read_excel(filename,index_col=None,header=0)
                        header_row=sales.index[sales.iloc[:,0] == 'Наименование ТТ'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        print("Прочитали содержимое файла")
                        print("----")
                        sales=sales[sales['Наименование ТТ'].notnull()]
                        #------------------------------------------------------------------------

                        geo=sales[['Наименование ТТ',"Формат","РЦ"]]
                    
                        geo=geo.drop_duplicates()
                        geo=geo[geo['Наименование ТТ']!="Grand Total"]
                        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
                        geo=geo[geo['Наименование ТТ']!="Общий итог"]
                        sales=sales.drop(['Формат','РЦ'],axis=1)





                        #-----------------------------------------------------------------------
                        sales=pd.melt(sales,id_vars='Наименование ТТ',var_name="SKU", value_name='stock')
                        sales=sales[sales.stock.notnull()]
                        sales=sales[sales['Наименование ТТ']!="Grand Total"]
                        sales=sales[sales['Наименование ТТ']!="Общий Итог"]
                        sales=sales[sales['Наименование ТТ']!="Общий итог"]
                        sales['SKU']=sales['SKU'].str.lower()
                        filename=filename.replace(".xlsm","")
                        sales["DATE"]="01/"+filename[-4:-2]+"/2022"
                        #--------------------------
                        print("Транспонировали и присвоили дату")
                        print("----")
                        print(sales['DATE'].unique()[0], " - Присвоенная дата")
                        print("----")
                        #--------------------------
                        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
                        sales['mrp2']=sales['mrp2'].str.replace(" ","")
                        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
                        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
                        sales['mrp2']=sales.mrp2.fillna(0)
                        sales=sales.rename(columns={"mrp2": "MRP"})
                        sales=sales.drop(['mrp1'],axis=1)
                        #--------------------------
                        print("Добавили МРЦ")
                        print("----")
                        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
                        print("Обработали ",filename)
                        print("----")
                        print("так выглядит обработанный массив")
                        print("----")
                        print(sales.head(5).to_string(index=False))
                        print("----")
                        
                        print(filename[-4:-2],"  Полученная сумма продаж без учета блоков: ",sales.stock.sum())
                        m_st_result=filename+" Полученная сумма продаж без учета блоков: "+str(sales.stock.sum())
                        results_list.append(m_st_result)
                    except Exception as e:
                        m_st_result=filename +" НЕ ПОЛУЧИЛОСЬ"
                        results_list.append(m_st_result)
                        traceback.print_exception(e)
                        wait_for_it = input('Press enter to close the terminal window')
                        
                        print(filename,"НЕ ПОЛУЧИЛОСЬ")
                        
                        pass
                    frame=sales
                    frame = frame.astype({"Наименование ТТ": str})
                    #----------------------------------------------------------------------------------
                    wb = xw.Book(result_file_geo)
                    ws = wb.sheets["GEO Monthly Stock"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "Monthly_Stock"
                    
                    #Append Non-Mapped GEO
                    ws = wb.sheets["GEO"]
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    rows=geo_init.shape[0]
                    ws = wb.sheets["Monthly Stock"]
                    wb.api.RefreshAll()
                    time.sleep(35)
                    df = ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    ws = wb.sheets["GEO"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = df
                    wb.save()
                    wb.close()
                    #DONE
                    #----------------------------------------------------------------------------------
                    #--------------------------    
                    print("Теперь соединим файл с исходным")
                    #current table
                    current_df=pd.read_csv(in_m_st , delimiter="\t")
                    current_df = current_df.astype({"Наименование ТТ": str})
                    print("Прочитали Файл")
                    new_total=pd.concat([frame,current_df])
                    new_total=new_total[new_total['SKU'].notnull()]
                    print("Добавили файл к исходному")
                    new_total.to_csv(result_file+"\\Current_RMC_Stock.txt", index=None, sep='\t', mode='w+')
                    print("Сохранили файл")
                    #--------------------------
                    print("----")
                    print("Можно закрывать файл")
                elif "рц" in filename and period in filename:
                    print("Нашли ",filename)
                    try:
                        #--------------------------
                        print("Начали обработку")
                        #--------------------------


                        sales=pd.read_excel(filename,index_col=None,header=0)
                        header_row=sales.index[sales.iloc[:,0] == 'РЦ'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        sales=sales[sales['РЦ'].notnull()]
                        print("Прочитали содержимое файла")
                        print("----")
                        sales=pd.melt(sales,id_vars='РЦ',var_name="SKU", value_name='stock')
                        sales=sales[sales.stock.notnull()]
                        sales=sales[sales['РЦ']!="Grand Total"]
                        sales=sales[sales['РЦ']!="Общий Итог"]
                        sales=sales[sales['РЦ']!="Общий итог"]
                        sales['SKU']=sales['SKU'].str.lower()
                        filename=filename.replace(".xlsm","")
                        sales["DATE"]="01/"+filename[-4:-2]+"/2022"
                        
                        #--------------------------
                        print("Транспонировали и присвоили дату")
                        print(sales['DATE'].unique())
                        #--------------------------
                        print("----")

                        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
                        sales['mrp2']=sales['mrp2'].str.replace(" ","")
                        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
                        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
                        sales['mrp2']=sales.mrp2.fillna(0)
                        sales=sales.rename(columns={"mrp2": "MRP"})
                        sales=sales.drop(['mrp1'],axis=1)
                        #--------------------------
                        
                        
                        print("----")
                        print("Добавили МРЦ")
                        print("----")
                        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
                        print("Обработали ",filename)
                        print("----")
                        print("так выглядит обработанный массив")
                        print("----")
                        print(sales.head(5).to_string(index=False))
                        print("----")
                        print(filename[-2:]," Сумма без учета блоков: ",sales.stock.sum())
                        print("----")
                        m_dc_result=filename+" Полученная сумма продаж без учета блоков: "+str(sales.stock.sum())
                        results_list.append(m_dc_result)
                    except Exception as e:
                        m_dc_result=filename +"НЕ ПОЛУЧИЛОСЬ"
                        results_list.append(m_dc_result)
                        traceback.print_exception(e)
                        wait_for_it = input('Press enter to close the terminal window')
                        
                        print(filename,"НЕ ПОЛУЧИЛОСЬ")
                        
                        pass
                    frame=sales
                    #--------------------------
                    print("Теперь соединим файл с исходным")
                    #current table
                    current_df=pd.read_csv(in_m_dc , delimiter="\t")
                    
                    print("Прочитали Файл")
                    new_total=pd.concat([frame,current_df])
                    new_total=new_total[new_total['SKU'].notnull()]
                    print("Добавили файл к исходному")
                    #--------------------------
                    new_total.to_csv(result_file+"\\Current_monthly_Stock_DC.txt", index=None, sep='\t', mode='w+')
                    print("Сохранили файл")
                    print("----")
                    print("Можно закрывать файл")
            #MONTHLY------------------------------------------------------------------------------ 

        elif selection=="Weekly":
            #WEEKLY-------------------------------------------------------------------------------
            
                if "рц" in filename and period in filename and "купоны" not in filename:
                    print("Нашли ",filename)
                
                    try:
                        #--------------------------
                        print("Начали обработку")
                        #--------------------------


                        sales=pd.read_excel(filename,index_col=None,header=0)
                        header_row=sales.index[sales.iloc[:,0] == 'РЦ'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        sales=sales[sales['РЦ'].notnull()]
                        print("Прочитали содержимое файла")
                        print("----")
                        sales=pd.melt(sales,id_vars='РЦ',var_name="SKU", value_name='stock')
                        sales=sales[sales.stock.notnull()]
                        sales=sales[sales['РЦ']!="Grand Total"]
                        sales=sales[sales['РЦ']!="Общий Итог"]
                        sales=sales[sales['РЦ']!="Общий итог"]
                        sales['SKU']=sales['SKU'].str.lower()
                        filename=filename.replace(".xlsm","")
                        week = int(filename[-2:])
                        year = 2022
                        date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
                        #--------------------------
                        print("Транспонировали и присвоили дату")
                        #--------------------------
                        print("----")

                        sales["DATE"]=date
                        sales["week"]=week
                        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
                        sales['mrp2']=sales['mrp2'].str.replace(" ","")
                        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
                        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
                        sales['mrp2']=sales.mrp2.fillna(0)
                        sales=sales.rename(columns={"mrp2": "MRP"})
                        sales=sales.drop(['mrp1'],axis=1)
                        #--------------------------
                        
                        print("Присвоили дату: ",date)
                        print("----")
                        print("Добавили МРЦ")
                        print("----")
                        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
                        print("Обработали ",filename)
                        print("----")
                        print("так выглядит обработанный массив")
                        print("----")
                        print(sales.head(5).to_string(index=False))
                        print("----")
                        print(filename[-2:]," Сумма без учета блоков: ",sales.stock.sum())
                        print("----")
                        w_dc_result=filename+" Полученная сумма продаж без учета блоков: "+str(sales.stock.sum())
                        results_list.append(w_dc_result)
                    except Exception as e:
                        w_dc_result= filename + " НЕ ПОЛУЧИЛОСЬ"
                        results_list.append(w_dc_result)
                        traceback.print_exception(e)
                        wait_for_it = input('Press enter to close the terminal window')
                        print(filename,"НЕ ПОЛУЧИЛОСЬ")
                        
                        pass
                    frame=sales
                    #--------------------------
                    print("Теперь соединим файл с исходным")
                    #current table
                    current_df=pd.read_csv(in_w_dc , delimiter="\t")
                    
                    print("Прочитали Файл")
                    new_total=pd.concat([frame,current_df])
                    new_total=new_total[new_total['SKU'].notnull()]
                    print("Добавили файл к исходному")
                    #--------------------------
                    new_total.to_csv(result_file+"\\Current_weekly_Stock_DC.txt", index=None, sep='\t', mode='w+')
                    print("Сохранили файл")
                    print("----")
                    print("Можно закрывать файл")
                elif "тт" in filename and period in filename and "купоны" not in filename:
                    print("Нашли ",filename)
                    
                    try:
                        #--------------------------
                        print("Начали обработку")
                        #--------------------------
                        sales=pd.read_excel(filename,index_col=None,header=0)
                        header_row=sales.index[sales.iloc[:,0] == 'Наименование ТТ'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        sales=sales[sales['Наименование ТТ'].notnull()]
                        print("Прочитали содержимое файла")
                        print("----")
                        #ADD GEO TABLE-----------------------------------------------------------
                        geo=sales[['Наименование ТТ','РЦ','Филиал','Формат']]
                        
                        geo=geo.drop_duplicates()
                        geo=geo[geo['Наименование ТТ']!="Grand Total"]
                        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
                        geo=geo[geo['Наименование ТТ']!="Общий итог"]
                        sales=sales.drop(['РЦ','Филиал','Формат'],axis=1)
                        #ADD GEO TABLE-----------------------------------------------------------
                        sales=pd.melt(sales,id_vars='Наименование ТТ',var_name="SKU", value_name='stock')
                        sales=sales[sales.stock.notnull()]
                        sales=sales[sales['Наименование ТТ']!="Grand Total"]
                        sales=sales[sales['Наименование ТТ']!="Общий Итог"]
                        sales=sales[sales['Наименование ТТ']!="Общий итог"]
                        sales['SKU']=sales['SKU'].str.lower()
                        filename=filename.replace(".xlsm","")
                        week = int(filename[-2:])
                        year = 2022
                        date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
                        
                        sales["DATE"]=date
                        sales["week"]=week
                        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
                        sales['mrp2']=sales['mrp2'].str.replace(" ","")
                        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
                        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
                        sales['mrp2']=sales.mrp2.fillna(0)
                        sales=sales.rename(columns={"mrp2": "MRP"})
                        sales=sales.drop(['mrp1'],axis=1)
                        #--------------------------
                        print("Транспонировали и присвоили дату")
                        print("----")
                        #--------------------------
                        print("Присвоили дату: ",date)
                        print("----")
                        print("Добавили МРЦ")
                        print("----")
                        print("SKU без МРЦ : ",sales[sales["MRP"]==0].SKU.unique())
                        print("----")
                        print("Обработали ",filename)
                        print("----")
                        print("так выглядит обработанный массив")
                        print("----")
                        
                        print(sales.head(5).to_string(index=False))
                        
                        print("----")
                        print(filename[-2:]," Сумма без учета блоков: ",sales.stock.sum())
                        print("--------------------------------------")
                        w_pos_result=filename+" Полученная сумма продаж без учета блоков: "+str(sales.stock.sum())
                        results_list.append(w_pos_result)
                    except Exception as e:
                        w_pos_result=filename +" НЕ ПОЛУЧИЛОСЬ"
                        results_list.append(w_pos_result)
                        traceback.print_exception(e)
                        wait_for_it = input('Press enter to close the terminal window')
                        print(filename,"НЕ ПОЛУЧИЛОСЬ")
                        
                    frame=sales
                    frame = frame.astype({"Наименование ТТ": str})
                    wb = xw.Book(result_file_geo)
                    ws = wb.sheets["GEO Weekly ST"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "Weekly_ST"
                    #Append Non-Mapped GEO
                    ws = wb.sheets["GEO"]
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    rows=geo_init.shape[0]
                    ws = wb.sheets["Weekly ST"]
                    wb.api.RefreshAll()
                    time.sleep(35)
                    df = ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    ws = wb.sheets["GEO"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = df
                    
                    wb.save()
                    wb.close()
                    #DONE
                    
                    #--------------------------
                    print("Теперь соединим файл с исходным")
                    #current table
                    current_df=pd.read_csv(in_w_stock , delimiter="\t")
                    current_df = current_df.astype({"Наименование ТТ": str})
                    print("Прочитали Файл")
                    print("----")
                    new_total=pd.concat([frame,current_df])
                    new_total=new_total[new_total['SKU'].notnull()]
                    print("Добавили файл к исходному")
                    print("----")
                    #--------------------------
                    new_total.to_csv(result_file+"\\Current_weekly_Stock.txt", index=None, sep='\t', mode='w+')
                    print("Сохранили файл")
                    #--------------------------
                    print("Можно закрывать программу")
                    #--------------------------
                elif "продажи" in filename and period in filename and "купоны" not in filename:
                    print("Нашли ",filename)
                    
                    try:
                        #--------------------------
                        print("Начали обработку")
                        #--------------------------
                        sales=pd.read_excel(filename,index_col=None,header=0)
                        header_row=sales.index[sales.iloc[:,0] == 'Магазин'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        sales=sales[sales['Магазин'].notnull()]
                        print("Прочитали содержимое файла")
                        print("----")
                        #ADD GEO TABLE-----------------------------------------------------------
                        #------------------------------------------------------------------------
                        geo=sales[['Магазин','FRMT','Филиал','РЦ (ОС)']]
                        geo=geo.rename(columns={"Магазин": "Наименование ТТ", "FRMT":"Формат","РЦ (ОС)":"РЦ"})
                        geo=geo.drop_duplicates()
                        geo=geo[geo['Наименование ТТ']!="Grand Total"]
                        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
                        geo=geo[geo['Наименование ТТ']!="Общий итог"]

                        sales=sales.drop(['FRMT','Филиал','РЦ (ОС)'],axis=1)
                        #------------------------------------------------------------------------
                        #ADD GEO TABLE-----------------------------------------------------------
                        sales=pd.melt(sales,id_vars='Магазин',var_name="SKU", value_name='sales')
                        sales=sales[sales.sales.notnull()]
                        sales=sales[sales['Магазин']!="Grand Total"]
                        sales=sales[sales['Магазин']!="Общий Итог"]
                        sales=sales[sales['Магазин']!="Общий итог"]
                        sales['SKU']=sales['SKU'].str.lower()
                        filename=filename.replace(".xlsm","")
                        
                        week = int(filename[-2:])
                        year = 2022
                        date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
                        
                        sales["DATE"]=date
                        sales["week"]=week
                        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
                        sales['mrp2']=sales['mrp2'].str.replace(" ","")
                        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
                        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
                        sales['mrp2']=sales.mrp2.fillna(0)
                        sales=sales.rename(columns={"mrp2": "MRP"})
                        sales=sales.drop(['mrp1'],axis=1)
                        #--------------------------
                        print("Транспонировали и присвоили дату")
                        print("----")
                        #--------------------------
                        print("Присвоили дату: ",date)
                        print("----")
                        
                        print("Добавили МРЦ")
                        print("----")
                        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
                        print("----")
                        print("Обработали ",filename)
                        
                        print("----")
                        print(filename[-2:]," Сумма без учета блоков: ",sales.sales.sum())
                        print("----")
                        w_sales_result=filename+" Полученная сумма продаж без учета блоков: "+str(sales.sales.sum())
                        results_list.append(w_sales_result)
                    except Exception as e:
                        w_sales_result=filename+ " НЕ ПОЛУЧИЛОСЬ"
                        results_list.append(w_sales_result)
                        traceback.print_exception(e)
                        wait_for_it = input('Press enter to close the terminal window')
                        print(filename,"НЕ ПОЛУЧИЛОСЬ")
                        
                    frame=sales
                    frame = frame.astype({"Магазин": str})
                    wb = xw.Book(result_file_geo)
                    ws = wb.sheets["GEO Weekly Sales"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "Weekly_Sales"
                    #Append Non-Mapped GEO
                    ws = wb.sheets["GEO"]
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    rows=geo_init.shape[0]
                    ws = wb.sheets["Weekly Sales"]
                    wb.api.RefreshAll()
                    time.sleep(35)
                    df = ws.range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
                    ws = wb.sheets["GEO"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = df
                    wb.save()
                    wb.close()
                    #DONE
                    #--------------------------
                    print("Теперь соединим файл с исходным")
                    #current table
                    
                    current_df=pd.read_csv(in_w_sales , delimiter="\t")
                    current_df = current_df.astype({"Магазин": str})
                    print("Прочитали Файл")
                    new_total=pd.concat([frame,current_df])
                    new_total=new_total[new_total['SKU'].notnull()]
                    print("Добавили файл к исходному")
                    #--------------------------
                    new_total.to_csv(result_file+"\\Current_weekly_Sales.txt", index=None, sep='\t', mode='w+')
                    print("Сохранили файл")
            #WEEKLY-------------------------------------------------------------------------------
        



    results= pd.DataFrame(results_list)

    results.to_csv(result_file+"\\PROGRAM_RESULTS.txt", index=None, sep='\t', mode='w+')
    print("TOTAL EXECUTION TIME ",time.time()-start_time)
    sg.popup("GOOD JOB, TOTAL EXECUTION TIME ",time.time()-start_time) 

    pbi=input("Будем обновлять POWER BI ?  [y/n] ") 
    if pbi=="y":
        try:
            subprocess.run(pbi_file,timeout=70)
        except:
            print("Started Refreshing")
            pg.moveTo(890,139)
            pg.doubleClick(890,139)
            pg.doubleClick(890,139)
            print("Please wait 2 minutes")
            time.sleep(120)
            pg.moveTo(41,28)
            pg.doubleClick(41,28)
            pg.doubleClick(41,28)
            pg.doubleClick(41,28)
            time.sleep(50)
            print("POWER BI File Successfully Saved! ")
            
    else:
        "..."
    sg.popup("PBI FILE SUCCESSFULLY UPDATED!")

if client=="RW":
    sg.theme('DarkBlue')

    layout = [
        [sg.T("Input data Storage folder:", s=45,justification="l"), sg.I(key="-FOLDER-"), sg.FolderBrowse()],
        [sg.Text("Input month num",s=45,justification="l"),sg.InputText( key='-DATE-')],
        [sg.Text("Choose data modification type",s=45,justification="l"),sg.Listbox(values=['RMC', 'RRP'], size=(45, 2), select_mode='single', key='-ETL-')],
        [sg.T("Input main file with all data:",s=45, justification="l"), sg.I(key="-MAIN-"), sg.FileBrowse(file_types=(("Text Files", "*.txt"),))],
        [sg.T("Input storage to save new data:", s=45,justification="l"), sg.I(key="-DESTINATION-"), sg.FolderBrowse()],
        [sg.Submit( )]
    ]

    window = sg.Window('Программа обработки продаж КиБ', layout)

    event, values = window.read()

    user_path=str(values['-FOLDER-'])
    selection=str(values['-ETL-'][0])
    month=str(values['-DATE-'])
    result_file=str(values['-MAIN-'])
    result_storage=str(values['-DESTINATION-'])
    window.close()

    all_files=glob.glob(user_path+"/*.xls*")
    li=[]

    if selection=="RMC":
        for filename in all_files:
            if "Logic" not in filename:
                try:
                        print("let's start from ",filename)
                        sales=pd.read_excel(filename,sheet_name="TDSheet" )
                        
                        print("Looking for <Магазин> in first column header")
                        header_row=sales.index[sales.iloc[:,1] == 'Магазин'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row-1]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        sales.columns.values[1] = "Магазин"
                        sales = sales.rename(columns={np.nan: 'useless','Всего':'useless',"NaN":"useless" })
                        sales.columns = sales.columns.fillna('useless')
                        sales=sales.drop(['useless'],axis=1)

                        print("Deleting Other Columns")

                        sales['Группа магазина']=pd.to_numeric(sales['Группа магазина'], errors='coerce')

                        print("Deleting Non-Numeric Rows")

                        sales=sales[ (sales['Группа магазина'].notnull())]
                        sales=sales.reset_index()
                        sales = sales.melt( id_vars=['index','Группа магазина','Магазин'],var_name="SKU",value_name="SALES AMOUNT")
                        sales=sales.drop(['index','Группа магазина'],axis=1)
                        sales['SKU']=sales.SKU.str.lower()



                        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
                        sales['mrp2']=sales['mrp2'].str.replace(" ","")
                        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
                        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
                        empty_mrp=sales[sales['mrp2'].isnull()]
                        sales['MRP']=sales['mrp2']

                        print("Extracting MRPs from SKU Name")

                        sales["Магазин"]=sales.Магазин.str.replace("№",'')
                        

                        sales[['address','address2']]=sales.Магазин.str.split('\W',expand=True,regex=True,n=1)
                        sales=sales.drop(['Магазин','mrp1','mrp2','address2'],axis=1)

                        print("Creating Date Column")

                        sales['DATE']="01."+month+".2022"
                        print("Given Date: ",sales.DATE.unique()[0])
                        sales=sales.rename(columns={"address": "ADDRESS"})
                        sales = sales[['DATE', 'ADDRESS', 'SKU', 'MRP', 'SALES AMOUNT']]
                        sales=sales[sales['MRP'].notnull()]
                        sales=sales[sales['SALES AMOUNT'].notnull()]
                        sales['DATE']=sales['DATE']+" 0:00:00"

                        book=xw.Book(filename)
                        ws = book.sheets.active

                        Total_in=ws.range("G10").value
                        book.save()
                        book.close()
                        Total=sales['SALES AMOUNT'].sum()+empty_mrp['SALES AMOUNT'].sum()
                        print("RAW RESULTS (G10 cell value): ",Total_in)
                        print("FINAL RESULTS: ",Total)
                        print("RESULTS WHICH WILL BE LOADED: ",sales['SALES AMOUNT'].sum())
                        print("RESULTS WHICH WON'T BE LOADED: ",empty_mrp['SALES AMOUNT'].sum())
                        print("These are empty mrp skus: ",empty_mrp['SKU'].unique())
                        print("----------")
                        print("These are unique mrps: ",sales['MRP'].unique())
                        if Total_in==Total:
                            print("EVERYTHING'S OKAY :)")
                        else:
                            print("Total Sum is not equal to the Value in G10 cell")
                        print("----------")
                        li.append(sales)
                except Exception as e:
                    traceback.print_exception(e)

        frame=pd.concat(li,axis=0,ignore_index=True)

        frame.to_csv(user_path+"\\SALES_TEST.txt", index=None, sep='\t', mode='w+')
        print("NOW LET'S APPEND EXISTING DATA")
            
        current_df=pd.read_csv(result_file , delimiter="\t")

        current_df=current_df.rename(columns={"CURRENT RMC sell-out[DATE]": "DATE", "CURRENT RMC sell-out[SKU]": "SKU","CURRENT RMC sell-out[MRP]":"MRP","CURRENT RMC sell-out[SALES AMOUNT]":"SALES AMOUNT","CURRENT RMC sell-out[ADDRESS]":"ADDRESS"})

        new_total=pd.concat([frame,current_df])

        new_total.to_csv(result_storage+"\\CURRENT RMC sell-out.txt", index=None, sep='\t', mode='w+')
        print("DONE")

        new_total['M']=new_total.DATE.str.slice(start=3, stop=5, step=None)
        new_total['Y']=new_total.DATE.str.slice(start=6, stop=10, step=None)

        for i in range(2022,2023):
            year=str(i)
            for i in range(1,13):
                if i<10:
                    month="0"+str(i)
                else:
                    month=str(i)
                        
                result=new_total.loc[(new_total['M'] == month) & (new_total['Y'] == year), 'SALES AMOUNT'].sum()
                print(year," ",month, "RESULT: ",result)
        print("DOUBLE CHECK THIS PLEASE")
    if selection=="RRP":
        for filename in all_files:
            if "Logic" in filename:
                    try:
                        print("let's start from ",filename)
                        sales=pd.read_excel(filename,sheet_name="TDSheet" )
                        print("Looking for <Магазин> in first column header")
                        header_row=sales.index[sales.iloc[:,1] == 'Магазин'].tolist()
                        header_row=header_row[0]
                        header=sales.iloc[header_row-1]
                        all_rows=header_row+1
                        sales= sales[all_rows:]
                        sales.columns=header
                        sales.columns.values[1] = "Магазин"
                        sales = sales.rename(columns={np.nan: 'useless','Всего':'useless',"NaN":"useless" })
                        sales.columns = sales.columns.fillna('useless')
                        sales=sales.drop(['useless'],axis=1)
                        sales['Группа магазина']=pd.to_numeric(sales['Группа магазина'], errors='coerce')
                        sales=sales[ (sales['Группа магазина'].notnull())]
                        sales=sales.reset_index()
                        sales = sales.melt( id_vars=['index','Группа магазина','Магазин'],var_name="SKU",value_name="SALES AMOUNT")
                        sales=sales.drop(['index','Группа магазина'],axis=1)
                        sales['SKU']=sales.SKU.str.lower()



                        print("Creating Blank MRP Column")
                        sales['MRP']=""
                        sales["Магазин"]=sales.Магазин.str.replace("№",'')
                        sales[['address','address2']]=sales.Магазин.str.split('\W',expand=True,regex=True,n=1)
                        sales=sales.drop(['Магазин','address2'],axis=1)
                        sales['DATE']="01."+month+".2022"
                        sales=sales.rename(columns={"address": "ADDRESS"})
                        sales = sales[['DATE', 'ADDRESS', 'SKU', 'MRP', 'SALES AMOUNT']]
                        
                        sales=sales[sales['SALES AMOUNT'].notnull()]
                        sales['DATE']=sales['DATE']+" 0:00:00"

                        book=xw.Book(filename)
                        ws = book.sheets.active

                        Total_in=ws.range("G10").value
                        book.save()
                        book.close()
                        Total=sales['SALES AMOUNT'].sum()
                        print("RAW RESULTS (G10 cell value): ",Total_in)
                        print("FINAL RESULTS: ",Total)
                        print("RESULTS WHICH WILL BE LOADED: ",sales['SALES AMOUNT'].sum())
                        
                        
                        print("----------")
                        print("These are unique mrps: ",sales['MRP'].unique())
                        if Total_in==Total:
                            print("EVERYTHING'S OKAY :)")
                        else:
                            print("Total Sum is not equal to the Value in G10 cell")
                        print("----------")
                        li.append(sales)
                    except Exception as e:
                        traceback.print_exception(e)

        frame=pd.concat(li,axis=0,ignore_index=True)

        frame.to_csv(user_path+"\\SALES_TEST.txt", index=None, sep='\t', mode='w+')
        print("NOW LET'S APPEND EXISTING DATA")
        current_df=pd.read_csv(result_file , delimiter="\t")
        current_df=current_df.rename(columns={"CURRENT RRP sell-out[DATE]": "DATE", "CURRENT RRP sell-out[SKU]": "SKU","CURRENT RRP sell-out[MRP]":"MRP","CURRENT RRP sell-out[SALES AMOUNT]":"SALES AMOUNT","CURRENT RRP sell-out[ADDRESS]":"ADDRESS"})
        
        new_total=pd.concat([frame,current_df])

        new_total.to_csv(result_storage+"\\CURRENT RRP sell-out.txt", index=None, sep='\t', mode='w+')
        print("DONE")

        new_total['M']=new_total.DATE.str.slice(start=3, stop=5, step=None)
        new_total['Y']=new_total.DATE.str.slice(start=6, stop=10, step=None)

        for i in range(2022,2023):
            year=str(i)
            for i in range(1,13):
                if i<10:
                    month="0"+str(i)
                else:
                    month=str(i)
                        
                result=new_total.loc[(new_total['M'] == month) & (new_total['Y'] == year), 'SALES AMOUNT'].sum()
                print(year," ",month, "RESULT: ",result)
        print("DOUBLE CHECK THIS PLEASE")
if client=="Dixy":
    sg.theme('DarkBlue')

    layout = [
        [sg.T("Input File with all URLs:", s=45,justification="l"), sg.I(key="-IN-", s=45), sg.FileBrowse(file_types=(("Excel Files", "*.xls*")))],
        
        [sg.Text("Choose data for update",s=45,justification="l"),sg.Listbox(values=['Monthly Stock', 'Monthly Sales','Weekly Sales & Stocks' ,'Etalon'], size=(45, 5), select_mode='multiple', key='-DESTINATION-')],
        [sg.Text("Input week or month num (Will be searched in files)",s=45,justification="l"),sg.InputText( key='-DATE-')],
        
        [sg.T("Input Storage Folder:", s=45,justification="l"), sg.I(key="-FOLDER-"), sg.FolderBrowse()],


        [sg.Submit( )]
    ]

    window = sg.Window('Программа обработки продаж Дикси', layout)

    event, values = window.read()

    user_path=str(values['-IN-'])
    selection=str(values['-DESTINATION-'])
    period=str(values['-DATE-'])
    result_file=str(values['-FOLDER-'])
    window.close()

    user_path=user_path.replace("\\","\\\\")
    result_file=result_file.replace("\\","\\\\")

    paths_data=pd.read_excel(user_path)
    paths_data['Type']=paths_data['Type'].str.lower()

    start_time=time.time()
    #---------------------------------------------------------------------------------------
    monthly_f_path_i=paths_data[paths_data.Type.str.contains("monthly folder")]['Path'].index[0]
    etalon_f_path_i=paths_data[paths_data.Type.str.contains("etalon folder")]['Path'].index[0]
    init_etalon_i=paths_data[paths_data.Type.str.contains("initial etalon")]['Path'].index[0]
    monthly_stocks_f_path_i=paths_data[paths_data.Type.str.contains("monthly stocks folder")]['Path'].index[0]
    weekly_f_path_i=paths_data[paths_data.Type.str.contains("weekly folder")]['Path'].index[0]
    result_f_path_i=paths_data[paths_data.Type.str.contains("result")]['Path'].index[0]
    pbi_path_i=paths_data[paths_data.Type.str.contains("pbi")]['Path'].index[0]
    in_m_rmc_i=paths_data[paths_data.Type.str.contains("monthly sales")]['Path'].index[0]
    in_m_stocks_rmc_i=paths_data[paths_data.Type.str.contains("initial monthly stocks")]['Path'].index[0]
    in_w_sales_i=paths_data[paths_data.Type.str.contains("weekly sales")]['Path'].index[0]
    in_w_stock_i=paths_data[paths_data.Type.str.contains("weekly stock")]['Path'].index[0]
    
    #---------------------------------------------------------------------------------------
    monthly_f_path=paths_data[paths_data.Type.str.contains("monthly folder")]['Path'][monthly_f_path_i]

    etalon_f_path=paths_data[paths_data.Type.str.contains("etalon folder")]['Path'][etalon_f_path_i]
    init_etalon=paths_data[paths_data.Type.str.contains("initial etalon")]['Path'][init_etalon_i]
    monthly_stocks_f_path=paths_data[paths_data.Type.str.contains("monthly stocks folder")]['Path'][monthly_stocks_f_path_i]
    weekly_f_path=paths_data[paths_data.Type.str.contains("weekly folder")]['Path'][weekly_f_path_i]
    result_file=paths_data[paths_data.Type.str.contains("result")]['Path'][result_f_path_i]
    pbi_path=paths_data[paths_data.Type.str.contains("pbi")]['Path'][pbi_path_i]
    in_m_rmc=paths_data[paths_data.Type.str.contains("monthly sales")]['Path'][in_m_rmc_i]
    in_m_stocks_rmc=paths_data[paths_data.Type.str.contains("initial monthly stocks")]['Path'][in_m_stocks_rmc_i]
    in_w_sales=paths_data[paths_data.Type.str.contains("weekly sales")]['Path'][in_w_sales_i]
    in_w_stock=paths_data[paths_data.Type.str.contains("weekly stock")]['Path'][in_w_stock_i]

    in_m_rmc=in_m_rmc.replace("\u202a","")
    in_w_stock=in_w_stock.replace("\u202a","")
    in_w_sales=in_w_sales.replace("\u202a","")

    pbi_file='cmd /K ' +'"cd ' +pbi_path +' & Dixy.pbix"'


    print("Extracting Data")
    print("----")
    

    all_files_m=glob.glob(monthly_f_path+"/*.xls*")
    all_files_w=glob.glob(weekly_f_path+"/*.xls*")
    all_files_stock=glob.glob(monthly_stocks_f_path+"/*.xls*")
    all_files_etalon=glob.glob(etalon_f_path+"/*.xls*")
    sales_li=[]
    stock_li_m=[]
    stock_li=[]
    etalon_li=[]
  
    if "Weekly Sales & Stocks" in selection:
        for filename in all_files_w:
            if period in filename:
                print("Reading Sales & Stocks")
                try:
                    print("Reading Sales Sheet ",filename)
                    sales=pd.read_excel(filename,sheet_name="Продажи" )
                    filename_short=filename.replace(weekly_f_path,"")
                    week=re.findall(r'\d+', filename_short)
                    week="-".join(week)
                    week = int(week)
                    year = 2022
                    date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
                    sales['Date']=date
                    sales['week']=week
                    sales_li.append(sales)
                    print("Given Date ",date)
                    print("----")
                    print("Dataframe Look: ")
                    print("----")
                    print(sales.head(5).to_string(index=False))
                    print("----")
                    print("Reading Stock Sheet ",filename)

                    sales=pd.read_excel(filename,sheet_name="Остаток" )
                    filename_short=filename.replace(weekly_f_path,"")
                    week=re.findall(r'\d+', filename_short)
                    week="-".join(week)
                    week = int(week)
                    year = 2022
                    date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
                    sales['Date']=date
                    sales['week']=week
                    stock_li.append(sales)
                    print("Given Date ",date)
                    print("----")
                    print("Dataframe Look: ")
                    print("----")
                    print(sales.head(5).to_string(index=False))
                    print("----")
                except Exception as e:
                    traceback.print_exception(e)
                
        try:
            print("Finished Preparing Data")
            print("Started Grouping Data")
            frame_sales=pd.concat(sales_li)
            frame_stock=pd.concat(stock_li)
        except:
            print("Error concatenating data")
            pass
        try:
            print("Reading Initial Files")
            current_sales=pd.read_csv(in_w_sales , delimiter="\t")
            current_stock=pd.read_csv(in_w_stock , delimiter="\t")
            print("Finished Reading Initial Files")  
            new_total_sales=pd.concat([current_sales,frame_sales])
            new_total_stock=pd.concat([current_stock,frame_stock])
            print("Appended New Data") 
            new_total_sales.to_csv(result_file+"\\Current_Weekly_Sales.txt", index=None, sep='\t', mode='w+')
            new_total_stock.to_csv(result_file+"\\Current_Weekly_Stock.txt", index=None, sep='\t', mode='w+')
            print("Saved New Data")
        except:
            print("Error saving data")
    if "Etalon" in selection:
        for filename in all_files_etalon:
            if period in filename:

                try:
                    print("Reading Эталон Sheet ",filename)
                    sales=pd.read_excel(filename,sheet_name="Эталон" )
                    filename_short=filename.replace(etalon_f_path,"")
                    week=re.findall(r'\d+', filename_short)
                    week="-".join(week)
                    week = int(week)
                    year = 2022
                    date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
                    sales['Date']=date
                    sales['week']=week
                    etalon_li.append(sales)
                    print("Given Date ",date)
                    print("----")
                    print("Dataframe Look: ")
                    print("----")
                    print(sales.head(5).to_string(index=False))
                    print("----")
                except Exception as e:
                    traceback.print_exception(e)
        try:
            print("Finished Preparing Data")
            print("Started Grouping Data")
            frame_etalon=pd.concat(etalon_li)
                
        except:
            print("Error concatenating data")
            pass
        try:
            print("Reading Initial Files")
            current_etalon=pd.read_csv(init_etalon , delimiter="\t")
            print("Finished Reading Initial Files")  
            new_total_etalon=pd.concat([current_etalon,frame_etalon])
            print("Appended New Data") 
            new_total_etalon.to_csv(result_file+"\\Merged_Etalon.txt", index=None, sep='\t', mode='w+')
            print("Saved New Data")
            
        except:
            print("Error saving data")

    if "Monthly Sales" in selection:
        for filename in all_files_m:
            
            if period in filename:
                print("Starting From Monthly Sales Folder")
                try:
                    print("Reading Sales Sheet",filename)
                    sales=pd.read_excel(filename)
                    filename_short=filename.replace(monthly_f_path,"")
                    month=re.findall(r'\d+', filename_short)
                    month="/".join(month)
                    date="01/"+month
                    sales['Date']=date
                    
                    sales_li.append(sales)
                    print("Given Date ",date)
                    print("----")
                    
                    print("Dataframe Look: ")
                    print("----")
                    print(sales.head(5).to_string(index=False))
                    print("----")
                except Exception as e:
                    traceback.print_exception(e)
                    
        try:
            print("Finished Preparing Data")
            print("Started Grouping Data")
            frame_sales=pd.concat(sales_li)
            print("Reading Initial Files")
            current_sales=pd.read_csv(in_m_rmc , delimiter="\t")
            new_total_sales=pd.concat([current_sales,frame_sales])
            print("Appended New Data") 

            new_total_sales.to_csv(result_file+"\\Current_Sales.txt", index=None, sep='\t', mode='w+')
            print("Saved New Data")
              
        except:
            print("Error Appending Data")
            pass
    if "Monthly Stock" in selection:
        for filename in all_files_stock:
            
            if period in filename:
                print("Moving to Monthly Stock Folder")
                try:
                    print("Reading Sales Sheet",filename)
                    sales=pd.read_excel(filename)
                    filename_short=filename.replace(monthly_stocks_f_path,"")
                    month=re.findall(r'\d+', filename_short)
                    month="/".join(month)
                    date="01/"+month
                    sales['Date']=date
                    
                    stock_li_m.append(sales)
                    print("Given Date ",date)
                    print("----")
                    
                    print("Dataframe Look: ")
                    print("----")
                    print(sales.head(5).to_string(index=False))
                    print("----")
                except Exception as e:
                    traceback.print_exception(e)
                    
        try:
            print("Finished Preparing Data")
            print("Started Grouping Data")
            frame_stocks=pd.concat(stock_li_m)
            print("Reading Initial Files")
            current_stocks=pd.read_csv(in_m_stocks_rmc , delimiter="\t")
            new_total_stocks=pd.concat([frame_stocks,current_stocks])
            print("Appended New Data") 


            new_total_stocks.to_csv(result_file+"\\Current_Stock.txt", index=None, sep='\t', mode='w+')
            print("Saved New Data")
               
        except:
            print("Error Appending Data")
            pass
    pbi=input("Будем обновлять POWER BI (Компьютер должен быть без монитора)?  [y/n] ") 
    if pbi=="y":
        try:
            subprocess.run(pbi_file,timeout=70)
        except:
            print("Started Refreshing")
            pg.moveTo(890,139)
            pg.doubleClick(890,139)
            pg.doubleClick(890,139)
            print("Please wait 2 minutes")
            time.sleep(120)
            pg.moveTo(41,28)
            pg.doubleClick(41,28)
            pg.doubleClick(41,28)
            pg.doubleClick(41,28)
            time.sleep(50)
            print("POWER BI File Successfully Saved! ")
                
    else:
        "..."
    sg.popup("GOOD JOB, TOTAL EXECUTION TIME : ",time.time()-start_time)
       
if client=="Bristol":
    
    sg.theme('DarkBlue2')

    layout = [
        [sg.T("Input File with all URLs:", s=45,justification="l"), sg.I(key="-IN-", s=45), sg.FileBrowse(file_types=(("Excel Files", "*.xls*")))],
        
        [sg.Text("Input Month Num (Will be searched in files)",s=45,justification="l"),sg.InputText( key='-DATE-')],
        [sg.Text("Choose Data For Modification",s=25,justification="r")],
        [sg.Listbox(values=['OMRTK RMC', 'OMRTK RRP', 'OVCHIN RMC', 'OVCHIN RRP','RASTYAPINO RMC','RASTYAPINO RRP','KOOP RMC','KOOP RRP'], size=(20, 8), select_mode='multiple', key='-DESTINATION-')],
        [sg.Submit( )]
    ]

    window = sg.Window('BRISTOL_APP', layout)
    event, values = window.read()

    user_path=str(values['-IN-'])
    selection=str(values['-DATE-'])
    mods=list(values['-DESTINATION-'])
    window.close()
    print("-----------------------")
    print("USER INPUTS WERE SAVED")

    user_path=user_path.replace("\\","\\\\")


    paths_data=pd.read_excel(user_path)
    paths_data['Type']=paths_data['Type'].str.lower()

    start_time=time.time()
    #---------------------------------------------------------------------------------------
    files_i_all=paths_data[paths_data.Type.str.contains("all_files")]['Path'].index[0]
    omrtk_i_etl=paths_data[paths_data.Type.str.contains("omrtk_etl_file")]['Path'].index[0]
    omrtk_i_rrp_etl=paths_data[paths_data.Type.str.contains("omrtk_rrp_etl_file")]['Path'].index[0]

    ovchin_i_etl=paths_data[paths_data.Type.str.contains("ovchin_etl_file")]['Path'].index[0]
    ovchin_i_rrp_etl=paths_data[paths_data.Type.str.contains("ovchin_rrp_etl_file")]['Path'].index[0]


    rastyap_i_etl=paths_data[paths_data.Type.str.contains("rastyap_etl_file")]['Path'].index[0]
    rastyap_i_rrp_etl=paths_data[paths_data.Type.str.contains("rastyap_rrp_etl_file")]['Path'].index[0]


    partners_i_db=paths_data[paths_data.Type.str.contains("partners_db")]['Path'].index[0]
    bristol_i_db=paths_data[paths_data.Type.str.contains("bristol_db")]['Path'].index[0]

    koop_i_etl=paths_data[paths_data.Type.str.contains("koop_etl_file")]['Path'].index[0]
    koop_i_rrp_etl=paths_data[paths_data.Type.str.contains("koop_rrp_etl_file")]['Path'].index[0]
    #---------------------------------------------------------------------------------------

    omrtk_path_etl=paths_data[paths_data.Type.str.contains("omrtk_etl_file")]['Path'][omrtk_i_etl]
    omrtk_path_rrp_etl=paths_data[paths_data.Type.str.contains("omrtk_rrp_etl_file")]['Path'][omrtk_i_rrp_etl]


    ovchin_path_etl=paths_data[paths_data.Type.str.contains("ovchin_etl_file")]['Path'][ovchin_i_etl]
    ovchin_path_rrp_etl=paths_data[paths_data.Type.str.contains("ovchin_rrp_etl_file")]['Path'][ovchin_i_rrp_etl]


    rastyap_path_etl=paths_data[paths_data.Type.str.contains("rastyap_etl_file")]['Path'][rastyap_i_etl]
    rastyap_path_rrp_etl=paths_data[paths_data.Type.str.contains("rastyap_rrp_etl_file")]['Path'][rastyap_i_rrp_etl]


    koop_path_etl=paths_data[paths_data.Type.str.contains("koop_etl_file")]['Path'][koop_i_etl]
    koop_path_rrp_etl=paths_data[paths_data.Type.str.contains("koop_rrp_etl_file")]['Path'][koop_i_rrp_etl]

    files_all=paths_data[paths_data.Type.str.contains("all_files")]['Path'][files_i_all]

    partners_db=paths_data[paths_data.Type.str.contains("partners_db")]['Path'][partners_i_db]
    bristol_db=paths_data[paths_data.Type.str.contains("bristol_db")]['Path'][bristol_i_db]
    print("-----------------------")
    print("CREATING CONNECTIONS WITH DATABASES")
    print("-----------------------")
    # DB CONNECTIONS
    partners_db_path="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="+partners_db+";"
    conn = pyodbc.connect(partners_db_path)
    bristol_db_path="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="+bristol_db+";"
    conn_bristol = pyodbc.connect(bristol_db_path)

    partners_list=[]
    bristol_db=pd.read_sql('select * from bristol', conn_bristol)
    omrtk_db = pd.read_sql('select * from omrtk', conn)
    ovch_db = pd.read_sql('select * from ovchinnikov', conn)
    rastyap_db = pd.read_sql('select * from rastyapino', conn)
    partners_list.append(omrtk_db)
    partners_list.append(ovch_db)
    partners_list.append(rastyap_db)
    partners_list.append(bristol_db)
    partners_frame=pd.concat(partners_list)
    partners_frame['Характеристика']=partners_frame['Характеристика'].str.replace(",",".")
    partners_frame['Характеристика']=partners_frame['Характеристика'].astype('float')
    partners_frame=partners_frame[partners_frame["Характеристика"]>=112]
    partners_frame['Характеристика']=partners_frame['Характеристика'].astype('str')
    partners_frame_gr=partners_frame.groupby(by=["SKU","Характеристика"]).sum()
    partners_frame_gr=partners_frame_gr.reset_index()
    merged_partners=partners_frame_gr


    print("STARTING THE PROGRAM")
    print("-----------------------")




    if "OMRTK RMC" in mods:
        
        # OMRTK
        try:
            all_files_omrtk=glob.glob(files_all+"/*.xls*")
            for filename in all_files_omrtk:
                filename=filename.lower()
                
                if "ploom" not in filename and selection in filename and "омртк" in filename:
                    print("WORKING WITH OMRTK RMC")
                    print("-----------------------")
                    print("Found ", filename)
                    print("-----------------------")
                    # GET MONTH NUM FROM FILENAME
                    init_m=filename.replace(".xlsx","")
                    init_m=init_m[-7:-5]
                    print("Extracted Month from filename: ",init_m)
                    print("-----------------------")
                    # OPEN ETL FILE'S GEO SHEET
                    #GEO
                    wb = xw.Book(omrtk_path_etl)

                    # READING GEO DF
                    ws = wb.sheets["Geo"]
                    ws.clear_formats()
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    # GETTING MAX ROW NUM
                    rows=geo_init.shape[0]
                    # GETTING MAX POS ID
                    max_omrtk_geo=geo_init['New Ad_ID'].str.replace("OMRTK_","")
                    max_omrtk_geo=pd.to_numeric(max_omrtk_geo,errors='coerce')
                    max_omrtk_geo=max_omrtk_geo.max()

                    print("Successfully Read Geo Sheet")
                    print("MAX POS ID : ", max_omrtk_geo)
                    print("-----------------------")

                    # READING INITIAL FILE
                    omrtk_init=pd.read_excel(filename)

                    #for i,item in enumerate(omrtk_init.columns):
                    #    b=omrtk_init.index[omrtk_init.iloc[:,i] == 'Код магазина'].tolist()
                    #   if len(b) > 0:
                    #        column_index=item
                    #header_row=omrtk_init.index[omrtk_init.loc[:,column_index] == 'Код магазина'].tolist()
                    #header_row=header_row[0]
                    #header=omrtk_init.iloc[header_row]
                    #all_rows=header_row+1
                    #omrtk_init= omrtk_init[all_rows:]
                    #omrtk_init.columns=header

                    omrtk_init['Код магазина']=omrtk_init['Код магазина'].str.strip()
                    omrtk_init['Адрес магазина']=omrtk_init['Адрес магазина'].str.strip()
                    omrtk_init['Адрес магазина'] = omrtk_init['Адрес магазина'].replace(r'\s+', ' ', regex=True)
                    omrtk_init["MONTH_DATE"]=str("01/"+init_m+"/2022")
                    omrtk_init["MONTH_DATE"]=pd.to_datetime(omrtk_init["MONTH_DATE"])
                    omrtk_init.MONTH_DATE = omrtk_init.MONTH_DATE.dt.strftime('%d/%m/%Y')

                    print("Successfully Read ",filename)
                    print("-----------------------")
                    
                    # APPENDING ETL FILE WITH INITIAL DATA
                    ws = wb.sheets["Партнеры"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = omrtk_init
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "data"
                    
                    print("Appended <Партнеры> Sheet")
                    print("-----------------------")
                    total_init_sum=omrtk_init['Продажи, шт.'].sum()

                    # MAPPING INITIAL DF WITH GEO
                    geo_mapping=omrtk_init.merge(geo_init, on='Код магазина', how='left')
                    geo_mapping=geo_mapping.drop_duplicates(subset=['Код магазина'])
                    geo_mapping=geo_mapping[geo_mapping['New Ad_ID'].isnull()]
                    for i,item in enumerate( geo_mapping['New Ad_ID']):
                        geo_mapping['New Ad_ID'].iloc[i]="OMRTK_"+str(max_omrtk_geo+i+1)
                    geo_mapping=geo_mapping[['New Ad_ID','Код магазина','Адрес магазина_x']]
                    geo_mapping=geo_mapping.rename(columns={"Адрес магазина_x": "Адрес магазина"})
                    geo_mapping['Сity']="г. Омск"
                    geo_mapping['State']="Omskaya"
                    geo_mapping['Область']="Омская обл."
                    geo_mapping['Region']="SIBERIA"
                    geo_mapping['Долгота']=""
                    geo_mapping['Широта']=""
                    geo_mapping['Siebel Code']=""
                    geo_mapping['Branch']="Albion_Omsk_Express"
                    geo_mapping=geo_mapping[['New Ad_ID','Код магазина','Адрес магазина','Сity','State','Область','Region','Долгота','Широта','Siebel Code','Branch']]
                    ws = wb.sheets["Geo"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = geo_mapping
                    geo_new=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    rows_new=geo_new.shape[0]
                    new_range="A"+str(rows+2)+":"+"A"+str(rows_new+1)
                    
                    
                    for a_cell in ws[new_range]:
                        a_cell.color = (0, 208, 142)

                    print("Added New POS")
                    print("-----------------------")

                    # REFRESHING EXCEL
                    print("Refreshing Excel for 25 sec.")
                    print("-----------------------")
                    wb.api.RefreshAll()
                    time.sleep(20)
                    # READING FOR LOAD DATA
                    ws = wb.sheets["For load"]
                    for_load=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value

                    for_load_total_sum=for_load['Продажи шт'].sum()

                    for_load=for_load[for_load['Код 1С Магазин'].isnull()]
                    blank_geos=for_load.shape[0]
                    if blank_geos>0:
                        sg.popup("MISSING SOME POS, PLEASE CHECK")
                        proceed=input("Willing to proceed?")
                        wb.api.RefreshAll()
                        time.sleep(20)
                        wb.save()
                        
                    else:
                        print("GOOD JOB, No Blank POS Found in <For Load> Sheet_OMRTK_RMC")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.save()
                        print("File Saved. Now Let's Move on!")
        except Exception as e:
            sg.popup("THERE WAS AN ERROR!")
            print("error-message----error-message----error-message")
            traceback.print_exception(e)
            print("error-message----error-message----error-message")
            error_msg=input("There was an ERROR in OMRTK RMC, pls press any key to close the program")                
        # OMRTK RRP
    if "OMRTK RRP" in mods:
        try:
            all_files_omrtk=glob.glob(files_all+"/*.xls*")
            for filename in all_files_omrtk:
                filename=filename.lower()
                
                if ("ploom" in filename) and (selection in filename)  and ("омртк" in filename):
                    print("WORKING WITH OMRTK RRP")
                    print("-----------------------")
                    print("Found ", filename)
                    print("-----------------------")

                    # GET MONTH NUM FROM FILENAME
                    init_m=filename.replace(".xlsx","")
                    init_m=init_m[-7:-5]

                    print("Extracted Month from filename: ",init_m)
                    print("-----------------------")
                    
                    # OPEN ETL FILE'S GEO SHEET
                    #GEO
                    wb = xw.Book(omrtk_path_rrp_etl)
                    
                    # READING INITIAL FILE
                    omrtk_init=pd.read_excel(filename)

                

                    omrtk_init['Код магазина']=omrtk_init['Код магазина'].str.strip()
                    omrtk_init['Адрес магазина']=omrtk_init['Адрес магазина'].str.strip()
                    omrtk_init['Адрес магазина'] = omrtk_init['Адрес магазина'].replace(r'\s+', ' ', regex=True)
                    omrtk_init["MONTH_DATE"]=str("01/"+init_m+"/2022")
                    omrtk_init["MONTH_DATE"]=pd.to_datetime(omrtk_init["MONTH_DATE"])
                    omrtk_init.MONTH_DATE = omrtk_init.MONTH_DATE.dt.strftime('%d/%m/%Y')
                    omrtk_init['Key_Update']=omrtk_init['Код магазина']+omrtk_init['Адрес магазина']
                    total_init_sum=omrtk_init['Продажи, шт.'].sum()
                    print("Successfully Read ",filename)
                    
                    print("-----------------------")
                    
                    # APPENDING ETL FILE WITH INITIAL DATA
                    ws = wb.sheets["Партнеры"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = omrtk_init
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "data"

                    print("Appended <Партнеры> Sheet")
                    print("-----------------------")
                    print("Refreshing for 25 sec")

                    wb.api.RefreshAll()
                    time.sleep(20)
                    
                    # READING GEO DF
                    ws = wb.sheets["Geo"]
                    ws.clear_formats()
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    # GETTING MAX ROW NUM
                    rows=geo_init.shape[0]
                    # GETTING MAX POS ID
                    max_omrtk_geo=geo_init['Код 1С Магазин'].str.replace("OMRTK_","")
                    max_omrtk_geo=pd.to_numeric(max_omrtk_geo,errors='coerce')
                    max_omrtk_geo=max_omrtk_geo.max()
                    print("Reading <Geo> Sheet")
                    print("-----------------------")
                    print("MAX POS ID : ",max_omrtk_geo)

                    
                    
                    
                    # MAPPING INITIAL DF WITH GEO
                    geo_mapping=omrtk_init.merge(geo_init, how='left', left_on="Key_Update",right_on="Custom")
                    geo_mapping=geo_mapping.drop_duplicates(subset=['Код магазина'])
                    geo_mapping=geo_mapping[geo_mapping['Код 1С Магазин'].isnull()]
                    for i,item in enumerate( geo_mapping['Код 1С Магазин']):
                        geo_mapping['Код 1С Магазин'].iloc[i]="OMRTK_"+str(max_omrtk_geo+i+1)
                    
                    geo_mapping=geo_mapping[['Код 1С Магазин','Код магазина','Адрес магазина']]
                    
                    geo_mapping['Сity']="г. Омск"
                    geo_mapping['State']="Omskaya"
                    geo_mapping['Область']="Омская обл."
                    geo_mapping['Region']="SIBERIA"
                    geo_mapping['Долгота']=""
                    geo_mapping['Широта']=""
                    geo_mapping['Siebel Code']=""
                    geo_mapping['Branch']="Albion_Omsk_Express"
                    ws = wb.sheets["New_Geo"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo_mapping
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "New_Geo"
                    print("Appended <New_Geo> Sheet")
                    print("-----------------------")
                    print("Refreshing for 25 sec")
                    # REFRESHING EXCEL
                    wb.api.RefreshAll()
                    time.sleep(20)
                    # READING FOR LOAD DATA
                    ws = wb.sheets["For load"]
                    for_load=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    for_load_total_sum=for_load['Продажи шт'].sum()

                    for_load=for_load[for_load['Код 1С Магазин'].isnull()]
                    blank_geos=for_load.shape[0]
                    if blank_geos>0:
                        print("MISSING SOME POS, PLEASE CHECK")
                        proceed=input("MISSING SOME POS, PLEASE CHECK. Willing to proceed?")
                        wb.api.RefreshAll()
                        time.sleep(20)
                        wb.save()
                        
                    else:
                        print("GOOD JOB, No Blank POS Found in <For Load> Sheet_OMRTK_RRP ")
                        wb.save()
                        print("File Saved. Now Let's Move on!")

        except Exception as e:
            sg.popup("THERE WAS AN ERROR!")
            print("error-message----error-message----error-message")
            traceback.print_exception(e)
            print("error-message----error-message----error-message")
            error_msg=input("There was an ERROR in OMRTK RRP, pls press any key to close the program")
        # OVCHIN
    if "OVCHIN RMC" in mods:
        try:
            all_files_ovchin=glob.glob(files_all+"/*.xls*")
            for filename in all_files_ovchin:
                filename=filename.lower()
                
                if "ploom" not in filename and selection in filename and ("овч" in filename):
                    print("WORKING WITH OMRTK RMC")
                    print("-----------------------")
                    print("Found ", filename)
                    print("-----------------------")
                    # GET MONTH NUM FROM FILENAME
                    init_m=filename.replace(".xlsx","")
                    init_m=init_m[-7:-5]
                    print("Extracted Month from filename: ",init_m)
                    print("-----------------------")
                    # OPEN ETL FILE'S GEO SHEET
                    #GEO
                    wb = xw.Book(ovchin_path_etl)

                    # READING GEO DF
                    ws = wb.sheets["Geo"]
                    ws.clear_formats()
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    # GETTING MAX ROW NUM
                    rows=geo_init.shape[0]
                    # GETTING MAX POS ID
                    max_ovchin_geo=geo_init['Код 1С Магазин'].str.replace("Ovchin_","")
                    max_ovchin_geo=pd.to_numeric(max_ovchin_geo,errors='coerce')
                    max_ovchin_geo=max_ovchin_geo.max()
                    print("Successfully Read Geo Sheet")
                    print("MAX POS ID : ", max_ovchin_geo)
                    print("-----------------------")

                    # READING INITIAL FILE
                    ovchin_init=pd.read_excel(filename)

                    

                    ovchin_init['Код магазина']=ovchin_init['Код магазина'].astype('str')
                    ovchin_init['Код магазина']=ovchin_init['Код магазина'].str.strip()
                    
                    ovchin_init['Код магазина'] = ovchin_init['Код магазина'].replace(re.compile('\.0'), "")
                    ovchin_init['Адрес магазина']=ovchin_init['Адрес магазина'].str.strip()
                    ovchin_init['Адрес магазина'] = ovchin_init['Адрес магазина'].replace(r'\s+', ' ', regex=True)
                    ovchin_init['Key_Update']=ovchin_init['Код магазина']+ovchin_init['Адрес магазина']
                    ovchin_init["MONTH_DATE"]=str("01/"+init_m+"/2022")
                    ovchin_init["MONTH_DATE"]=pd.to_datetime(ovchin_init["MONTH_DATE"])
                    ovchin_init.MONTH_DATE = ovchin_init.MONTH_DATE.dt.strftime('%d/%m/%Y')

                    total_init_sum=ovchin_init['Продажи, шт.'].sum()

                    print("Successfully Read ",filename)
                    print("-----------------------")
                    # APPENDING ETL FILE WITH INITIAL DATA
                    ws = wb.sheets["Партнеры"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = ovchin_init
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "data"
                    print("Appended <Партнеры> Sheet")
                    print("-----------------------")

                    # MAPPING INITIAL DF WITH GEO
                    geo_mapping=ovchin_init.merge(geo_init, on='Key_Update', how='left')
                    geo_mapping=geo_mapping[geo_mapping['Код 1С Магазин'].isnull()]
                    geo_mapping=geo_mapping[['Код магазина_x','Адрес магазина_x']]
                    geo_mapping=geo_mapping.rename(columns={"Адрес магазина_x": "Адрес магазина",'Код магазина_x':'Код магазина'})
                    geo_mapping=geo_mapping.drop_duplicates(subset=['Код магазина'])
                    geo_mapping['Код 1С Магазин']=""
                    for i,item in enumerate( geo_mapping['Код магазина']):
                                geo_mapping['Код 1С Магазин'].iloc[i]="Ovchin_"+str(max_ovchin_geo+i+1)
                    geo_mapping['Сity']=""
                    geo_mapping['State']="Nizhegorodskaya"
                    geo_mapping['Область']="Нижегородская обл."
                    geo_mapping['Region']="VOLGA"
                    geo_mapping['Долгота']=""
                    geo_mapping['Широта']=""
                    geo_mapping['Siebel Code']=""
                    geo_mapping['Branch']="Albion_N.Novgorod_Express"




                    geo_mapping=geo_mapping[['Код 1С Магазин','Код магазина','Адрес магазина','Сity','State','Область','Region','Долгота','Широта','Siebel Code','Branch']]
                    ws = wb.sheets["Geo"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = geo_mapping
                    geo_new=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    rows_new=geo_new.shape[0]
                    new_range="A"+str(rows+2)+":"+"A"+str(rows_new+1)
                    
                    
                    for a_cell in ws[new_range]:
                        a_cell.color = (0, 208, 142)
                    # REFRESHING EXCEL
                    wb.api.RefreshAll()
                    time.sleep(20)
                    # READING FOR LOAD DATA
                    ws = wb.sheets["For load"]
                    for_load=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value

                    for_load_total_sum=for_load['Продажи шт'].sum()

                    for_load=for_load[for_load['Код 1С Магазин'].isnull()]
                    blank_geos=for_load.shape[0]
                    
                    if blank_geos>0:
                        print("MISSING SOME POS, PLEASE CHECK")
                        proceed=input("MISSING SOME POS, PLEASE CHECK. Willing to proceed?")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.api.RefreshAll()
                        time.sleep(20)
                        wb.save()
                    else:
                        print("GOOD JOB, No Blank POS Found in <For Load> Sheet_OVCHIN_RMC")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.save()
                        print("File Saved. Now Let's Move on!")



        except Exception as e:
            sg.popup("THERE WAS AN ERROR!")
            print("error-message----error-message----error-message")
            traceback.print_exception(e)
            print("error-message----error-message----error-message")
            error_msg=input("There was an ERROR in OVCHIN RMC, pls press any key to close the program")           
        # OVCHIN RRP
    if "OVCHIN RRP" in mods:
        try:
            all_files_ovchin=glob.glob(files_all+"/*.xls*")
            for filename in all_files_ovchin:
                filename=filename.lower()
                
                if "ploom" in filename and selection in filename and ("овч" in filename):
                    print("WORKING WITH OVCHIN RRP")
                    print("-----------------------")
                    print("Found ", filename)
                    print("-----------------------")
                    # GET MONTH NUM FROM FILENAME
                    init_m=filename.replace(".xlsx","")
                    init_m=init_m[-7:-5]
                    print("Extracted Month from filename: ",init_m)
                    print("-----------------------")
                    # OPEN ETL FILE'S GEO SHEET
                    #GEO
                    wb = xw.Book(ovchin_path_rrp_etl)
                    
                    wb.api.RefreshAll()
                    time.sleep(20)
                    
                    # READING GEO DF
                    ws = wb.sheets["Geo"]
                    ws.clear_formats()
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    # GETTING MAX ROW NUM
                    rows=geo_init.shape[0]
                    # GETTING MAX POS ID
                    max_ovchin_geo=geo_init['Код 1С Магазин'].str.replace("Ovchin_","")
                    max_ovchin_geo=pd.to_numeric(max_ovchin_geo,errors='coerce')
                    max_ovchin_geo=max_ovchin_geo.max()
                    print("Successfully Read Geo Sheet")
                    print("MAX POS ID : ", max_ovchin_geo)
                    print("-----------------------")
                    # READING INITIAL FILE
                    ovchin_init=pd.read_excel(filename)
                    print("Successfully Read ",filename)
                    print("-----------------------")
                    

                    ovchin_init['Код магазина']=ovchin_init['Код магазина'].astype('str')
                    ovchin_init['Код магазина']=ovchin_init['Код магазина'].str.strip()
                    ovchin_init['Код магазина'] = ovchin_init['Код магазина'].replace(re.compile('\.0'), "")
                    ovchin_init['Адрес магазина']=ovchin_init['Адрес магазина'].str.strip()
                    ovchin_init['Адрес магазина'] = ovchin_init['Адрес магазина'].replace(r'\s+', ' ', regex=True)
                    ovchin_init['Key_Update']=ovchin_init['Код магазина']+ovchin_init['Адрес магазина']
                    ovchin_init["MONTH_DATE"]=str("01/"+init_m+"/2022")
                    ovchin_init["MONTH_DATE"]=pd.to_datetime(ovchin_init["MONTH_DATE"])
                    ovchin_init.MONTH_DATE = ovchin_init.MONTH_DATE.dt.strftime('%d/%m/%Y')
                    
                    # APPENDING ETL FILE WITH INITIAL DATA
                    ws = wb.sheets["Партнеры"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = ovchin_init
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "data"
                    print("Appended <Партнеры> Sheet")
                    print("-----------------------")
                    total_init_sum=ovchin_init['Продажи, шт.'].sum()

                    # # MAPPING INITIAL DF WITH GEO
                    geo_mapping=ovchin_init.merge(geo_init, how='left', left_on="Key_Update",right_on="Custom")
                    geo_mapping=geo_mapping.drop_duplicates(subset=['Код магазина'])
                    geo_mapping=geo_mapping[geo_mapping['Код 1С Магазин'].isnull()]
                    for i,item in enumerate( geo_mapping['Код 1С Магазин']):
                        geo_mapping['Код 1С Магазин'].iloc[i]="Ovchin_"+str(max_ovchin_geo+i+1)
                    
                    
                    geo_mapping['Сity']=""
                    geo_mapping['State']="Nizhegorodskaya"
                    geo_mapping['Область']="Нижегородская обл."
                    geo_mapping['Region']="VOLGA"
                    geo_mapping['Долгота']=""
                    geo_mapping['Широта']=""
                    geo_mapping['Siebel Code']=""
                    geo_mapping['Branch']="Albion_N.Novgorod_Express"
                    geo_mapping=geo_mapping[['Код 1С Магазин','Код магазина','Адрес магазина','Сity','State','Область','Region','Долгота','Широта','Siebel Code','Branch']]
                    
                    ws = wb.sheets["New_Geo"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo_mapping
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "New_Geo"
                    print("Appended <New_Geo> Sheet")
                    print("-----------------------")
                    # REFRESHING EXCEL
                    wb.api.RefreshAll()
                    time.sleep(20)
                    
                    # READING FOR LOAD DATA
                    ws = wb.sheets["For load"]
                    for_load=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value

                    for_load_total_sum=for_load['Продажи шт'].sum()

                    for_load=for_load[for_load['Код 1С Магазин'].isnull()]
                    blank_geos=for_load.shape[0]
                    if blank_geos>0:
                        print("MISSING SOME POS, PLEASE CHECK")
                        proceed=input("Willing to proceed?")
                        wb.api.RefreshAll()
                        time.sleep(20)
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.save()
                    else:
                        print("GOOD JOB, No Blank POS Found in <For Load> Sheet_OVCHIN_RRP")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.save()           
                    
        except Exception as e:
            sg.popup("THERE WAS AN ERROR!")
            print("error-message----error-message----error-message")
            traceback.print_exception(e)
            print("error-message----error-message----error-message")
            error_msg=input("There was an ERROR in OVCHIN RRP, pls press any key to close the program")           
            
        # RASTYAPINO
    if "RASTYAPINO RMC" in mods:
        try:
            all_files_rastyap=glob.glob(files_all+"/*.xls*")
            for filename in all_files_rastyap:
                filename=filename.lower()
                
                if "jti" not in filename and selection in filename and ("растяпино" in filename):
                    print("WORKING WITH RASTYAPINO RMC")
                    print("-----------------------")
                    print("Found ", filename)
                    print("-----------------------")
                    # GET MONTH NUM FROM FILENAME
                    init_m=filename.replace(".xlsx","")
                    init_m=init_m[-7:-5]
                    print("Extracted Month from filename: ",init_m)
                    print("-----------------------")
                    
                    # OPEN ETL FILE'S GEO SHEET
                    #GEO
                    wb = xw.Book(rastyap_path_etl)

                    # READING GEO DF
                    ws = wb.sheets["Geo"]
                    ws.clear_formats()
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    # GETTING MAX ROW NUM
                    rows=geo_init.shape[0]
                    # GETTING MAX POS ID
                    max_rastyap_geo=geo_init['New Ad_ID'].str.replace("Rastyap_R_","")
                    max_rastyap_geo=pd.to_numeric(max_rastyap_geo,errors='coerce')
                    max_rastyap_geo=max_rastyap_geo.max()
                    print("Successfully Read Geo Sheet")
                    print("MAX POS ID : ", max_rastyap_geo)
                    print("-----------------------")
                    

                    # READING INITIAL FILE
                    rastyap_init=pd.read_excel(filename)
                    
                    

                    rastyap_init['адрес']=rastyap_init['адрес'].str.strip()
                    
                    rastyap_init["MONTH_DATE"]=str("01/"+init_m+"/2022")
                    rastyap_init["MONTH_DATE"]=pd.to_datetime(rastyap_init["MONTH_DATE"])
                    rastyap_init.MONTH_DATE = rastyap_init.MONTH_DATE.dt.strftime('%d/%m/%Y')

                    print("Successfully Read ",filename)
                    print("-----------------------")

                    # APPENDING ETL FILE WITH INITIAL DATA
                    ws = wb.sheets["Партнеры"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = rastyap_init
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "data"
                    print("Appended <Партнеры> Sheet")
                    total_init_sum=rastyap_init['продажи, шт.'].sum()
                    print("-----------------------")

                    # MAPPING INITIAL DF WITH GEO
                    geo_mapping=rastyap_init.merge(geo_init,  how='left',left_on='адрес', right_on='Адрес магазина')
                    geo_mapping=geo_mapping[geo_mapping['Код магазина'].isnull()]
                    geo_mapping=geo_mapping[['New Ad_ID','Код магазина','адрес','Сity','State','Region']]
                    for i,item in enumerate( geo_mapping['Код магазина']):
                        geo_mapping['Код магазина'].iloc[i]="R_"+str(max_rastyap_geo+i+1)
                    geo_mapping['Сity']=""
                    geo_mapping['State']="Nizhegorodskaya"
                    geo_mapping['Область']="Нижегородская обл."
                    geo_mapping['Region']="VOLGA"
                    geo_mapping['Долгота']=""
                    geo_mapping['Широта']=""
                    geo_mapping['Siebel Code']=""
                    geo_mapping['Branch']="Albion_Rastyapino"
                    geo_mapping=geo_mapping[['New Ad_ID','Код магазина','адрес','Сity','State','Область','Region','Долгота','Широта','Siebel Code','Branch']]
                    ws = wb.sheets["Geo"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = geo_mapping
                    geo_new=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    rows_new=geo_new.shape[0]
                    new_range="A"+str(rows+2)+":"+"A"+str(rows_new+1)
                    
                    
                    for a_cell in ws[new_range]:
                        a_cell.color = (0, 208, 142)
                    print("Added New POS")
                    print("-----------------------")

                    # REFRESHING EXCEL
                    print("Refreshing Excel for 25 sec.")
                    print("-----------------------")

                    # REFRESHING EXCEL
                    wb.api.RefreshAll()
                    time.sleep(20)
                    # READING FOR LOAD DATA
                    ws = wb.sheets["For load"]
                    for_load=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value

                    for_load_total_sum=for_load['Продажи шт'].sum()

                    for_load=for_load[for_load['Код 1С Магазин'].isnull()]
                    blank_geos=for_load.shape[0]
                    if blank_geos>0:
                        print("MISSING SOME POS, PLEASE CHECK")
                        proceed=input("Willing to proceed?")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.api.RefreshAll()
                        time.sleep(20)
                        wb.save()
                    else:
                        print("GOOD JOB, No Blank POS Found in <For Load> Sheet_RASTYAPINO_RMC")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.save()
        except Exception as e:
            sg.popup("THERE WAS AN ERROR!")
            print("error-message----error-message----error-message")
            traceback.print_exception(e)
            print("error-message----error-message----error-message")
            error_msg=input("There was an ERROR in RASTYAPINO RMC, pls press any key to close the program")            
        # RASTYAPINO RRP
    if "RASTYAPINO RRP" in mods:
        try:
            all_files_rastyap=glob.glob(files_all+"/*.xls*")
            for filename in all_files_rastyap:
                filename=filename.lower()
                
                if "jti" in filename and selection in filename and ("растяпино" in filename):
                    print("WORKING WITH RASTYAPINO RRP")
                    
                    print("-----------------------")
                    print("Found ", filename)
                    print("-----------------------")
                    # GET MONTH NUM FROM FILENAME
                    init_m=filename.replace(".xlsx","")
                    init_m=init_m[-7:-5]
                    print("Extracted Month from filename: ",init_m)
                    print("-----------------------")
                    
                    # OPEN ETL FILE'S GEO SHEET
                    #GEO
                    wb = xw.Book(rastyap_path_rrp_etl)
                    
                    wb.api.RefreshAll()
                    time.sleep(20)
                    
                    # READING GEO DF
                    ws = wb.sheets["Geo"]
                    ws.clear_formats()
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    # GETTING MAX ROW NUM
                    rows=geo_init.shape[0]
                    # GETTING MAX POS ID
                    max_rastyap_geo=geo_init['Код 1С Магазин'].str.replace("Rastyap_R_","")
                    max_rastyap_geo=pd.to_numeric(max_rastyap_geo,errors='coerce')
                    max_rastyap_geo=max_rastyap_geo.max()
                    print("Reading <Geo> Sheet")
                    print("-----------------------")
                    print("MAX POS ID : ",max_rastyap_geo)

                    # READING INITIAL FILE
                    rastyap_init=pd.read_excel(filename)
                    
                    


                    rastyap_init['адрес']=rastyap_init['адрес'].str.strip()
                    
                    rastyap_init["MONTH_DATE"]=str("01/"+init_m+"/2022")
                    rastyap_init["MONTH_DATE"]=pd.to_datetime(rastyap_init["MONTH_DATE"])
                    rastyap_init.MONTH_DATE = rastyap_init.MONTH_DATE.dt.strftime('%d/%m/%Y')
                    print("Successfully Read ",filename)
                    
                    print("-----------------------")
                    
                    # APPENDING ETL FILE WITH INITIAL DATA
                    ws = wb.sheets["Партнеры"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = rastyap_init
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "data"
                    print("Appended <Партнеры> Sheet")
                    print("-----------------------")
                    total_init_sum=rastyap_init['продажи, шт.'].sum()
                    
                    
                    # # MAPPING INITIAL DF WITH GEO
                
                    geo_mapping=rastyap_init.merge(geo_init,  how='left',left_on='адрес', right_on='Адрес Магазина')
                    geo_mapping=geo_mapping[geo_mapping['Код 1С Магазин'].isnull()]
                    geo_mapping=geo_mapping[['Код 1С Магазин','Код Магазина','адрес','State','Region']]
                    for i,item in enumerate( geo_mapping['Код Магазина']):
                        geo_mapping['Код 1С Магазин'].iloc[i]="Rastyap_R_"+str(max_rastyap_geo+i+1)
                        geo_mapping['Код Магазина'].iloc[i]="R_"+str(max_rastyap_geo+i+1)
                        
                    geo_mapping['Сity']=""
                    geo_mapping['State']="Nizhegorodskaya"
                    geo_mapping['Область']="Нижегородская обл."
                    geo_mapping['Region']="VOLGA"
                    geo_mapping['Долгота']=""
                    geo_mapping['Широта']=""
                    geo_mapping['Siebel Code']=""
                    geo_mapping['Branch']="Albion_Rastyapino"
                    
                    geo_mapping=geo_mapping[['Код 1С Магазин','Код Магазина','адрес','Сity','State','Область','Region','Долгота','Широта','Siebel Code','Branch']]
                    ws = wb.sheets["New_Geo"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo_mapping
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "New_Geo"
                    print("Appended <New_Geo> Sheet")
                    print("-----------------------")
                    
                    # READING FOR LOAD DATA
                    ws = wb.sheets["For load"]
                    for_load=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value

                    for_load_total_sum=for_load['Продажи шт'].sum()

                    for_load=for_load[for_load['Код 1С Магазин'].isnull()]
                    blank_geos=for_load.shape[0]
                    if blank_geos>0:
                        print("MISSING SOME POS, PLEASE CHECK")
                        proceed=input("Willing to proceed?")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.api.RefreshAll()
                        time.sleep(20)
                        wb.save()
                    else:
                        print("GOOD JOB, No Blank POS Found in <For Load> Sheet_RASTYAPINO_RRP")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.save()
        except Exception as e:
            sg.popup("THERE WAS AN ERROR!")
            print("error-message----error-message----error-message")
            traceback.print_exception(e)
            print("error-message----error-message----error-message")
            error_msg=input("There was an ERROR in RASTYAPINO RRP, pls press any key to close the program")
        
        # koop
    if "KOOP RMC" in mods:
        
        try:
            all_files_koop=glob.glob(files_all+"/*.xls*")
            for filename in all_files_koop:
                filename=filename.lower()
                
                if "rrp" not in filename and selection in filename and ("кооп" in filename):
                    print("WORKING WITH KOOP RMC")
                    print("-----------------------")
                    print("Found ", filename)
                    print("-----------------------")
                    # GET MONTH NUM FROM FILENAME
                    init_m=filename.replace(".xlsx","")
                    init_m=init_m[-7:-5]
                    print("Extracted Month from filename: ",init_m)
                    print("-----------------------")
                    # OPEN ETL FILE'S GEO SHEET
                    #GEO
                    wb = xw.Book(koop_path_etl)

                    # READING GEO DF
                    ws = wb.sheets["Geo"]
                    
                    ws.clear_formats()
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    # GETTING MAX ROW NUM
                    rows=geo_init.shape[0]
                    # GETTING MAX POS ID
                    max_koop_geo=geo_init['Код 1С Магазин'].str.replace("KOOP_","")
                    max_koop_geo=pd.to_numeric(max_koop_geo,errors='coerce')
                    max_koop_geo=max_koop_geo.max()
                    print("Successfully Read Geo Sheet")
                    print("MAX POS ID : ", max_koop_geo)
                    print("-----------------------")
                    # READING INITIAL FILE
                    koop_init=pd.read_excel(filename)
                    
                    

                    koop_init['Юр лицо']=koop_init['Юр лицо'].str.strip()
                    
                    koop_init['Адрес магазина']=koop_init['Адрес магазина'].str.strip()
                    koop_init['Адрес магазина'] = koop_init['Адрес магазина'].replace(r'\s+', ' ', regex=True)
                    koop_init['Key_Update']=koop_init['Юр лицо']+koop_init['Адрес магазина']
                    koop_init["MONTH_DATE"]=str("01/"+init_m+"/2022")
                    koop_init["MONTH_DATE"]=pd.to_datetime(koop_init["MONTH_DATE"])
                    koop_init.MONTH_DATE = koop_init.MONTH_DATE.dt.strftime('%d/%m/%Y')
                    
                    total_init_sum=koop_init['Продажи в пачках, шт.'].sum()
                    print("Successfully Read ",filename)
                    print("-----------------------")

                    # APPENDING ETL FILE WITH INITIAL DATA
                    ws = wb.sheets["Партнеры"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = koop_init
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "data"
                    print("Appended <Партнеры> Sheet")
                    print("-----------------------")
                    # MAPPING INITIAL DF WITH GEO
                    geo_mapping=koop_init.merge(geo_init, on='Key_Update', how='left')
                    geo_mapping=geo_mapping[geo_mapping['Код 1С Магазин'].isnull()]  
                    geo_mapping=geo_mapping.drop_duplicates(subset=['Key_Update'])
                    geo_mapping['Код 1С Магазин']=""
                    geo_mapping=geo_mapping[['Код 1С Магазин','Юр лицо_x','Адрес магазина_x']]
                    geo_mapping=geo_mapping.rename(columns={'Адрес магазина_x': 'Адрес магазина','Юр лицо_x':'Юр лицо'})
                    for i,item in enumerate( geo_mapping['Код 1С Магазин']):
                                geo_mapping['Код 1С Магазин'].iloc[i]="KООП_"+str(1)
                    geo_mapping['Сity']=""
                    geo_mapping['State']="Nizhegorodskaya"
                    geo_mapping['Область']="Нижегородская обл."
                    geo_mapping['Region']="VOLGA"
                    geo_mapping['Долгота']=""
                    geo_mapping['Широта']=""
                    geo_mapping['Siebel Code']=""
                    geo_mapping['Branch']=""
                    geo_mapping=geo_mapping[['Код 1С Магазин','Юр лицо','Адрес магазина','Сity','State','Область','Region','Долгота','Широта','Siebel Code','Branch']]
                    ws = wb.sheets["Geo"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = geo_mapping
                    geo_new=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    rows_new=geo_new.shape[0]
                    new_range="A"+str(rows+2)+":"+"A"+str(rows_new+1)
                    
                    
                    for a_cell in ws[new_range]:
                        a_cell.color = (0, 208, 142)
                    
                    ws = wb.sheets["PARTNERS_DF"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = merged_partners
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "partners"
                    # REFRESHING EXCEL
                    wb.api.RefreshAll()
                    time.sleep(30)
                    # READING FOR LOAD DATA
                    ws = wb.sheets["For load"]
                    for_load=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value


                    ws = wb.sheets["For load Decimal"]
                    for_load_dec=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value

                    for_load_total_sum=for_load_dec['Продажи шт'].sum()

                    for_load=for_load[for_load['Код 1С Магазин'].isnull()]
                    blank_geos=for_load.shape[0]
                    
                    if blank_geos>0:
                        print("MISSING SOME POS, PLEASE CHECK")
                        proceed=input("Willing to proceed?")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.api.RefreshAll()
                        time.sleep(20)
                        wb.save()
                    else:
                        print("GOOD JOB, No Blank POS Found in <For Load> Sheet_KOOP_RMC")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.save()
        except Exception as e:
            sg.popup("THERE WAS AN ERROR!")
            print("error-message----error-message----error-message")
            traceback.print_exception(e)
            print("error-message----error-message----error-message")
            error_msg=input("There was an ERROR in KOOP RMC, pls press any key to close the program")                
    if "KOOP RRP" in mods: 
        try:                  
            # KOOP RRP
            all_files_koop=glob.glob(files_all+"/*.xls*")
            for filename in all_files_koop:
                filename=filename.lower()
                
                if ("сигареты" not in filename) and (selection in filename) and ("rrp" in filename)and ("кооп" in filename):
                    print("WORKING WITH KOOP RRP")
                    print("-----------------------")
                    print("Found ", filename)
                    print("-----------------------")
                    # GET MONTH NUM FROM FILENAME
                    init_m=filename.replace(".xlsx","")
                    init_m=init_m[-7:-5]
                    print("Extracted Month from filename: ",init_m)
                    print("-----------------------")
                    # OPEN ETL FILE'S GEO SHEET
                    #GEO
                    wb = xw.Book(koop_path_rrp_etl)

                    # READING GEO DF
                    ws = wb.sheets["Geo"]
                    ws.clear_formats()
                    geo_init=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    # GETTING MAX ROW NUM
                    rows=geo_init.shape[0]
                    # GETTING MAX POS ID
                    max_koop_geo=geo_init['Код 1С Магазин'].str.replace("KOOP_","")
                    max_koop_geo=pd.to_numeric(max_koop_geo,errors='coerce')
                    max_koop_geo=max_koop_geo.max()
                    print("Successfully Read Geo Sheet")
                    print("MAX POS ID : ", max_koop_geo)
                    print("-----------------------")
                    # READING INITIAL FILE
                    koop_init=pd.read_excel(filename)


                    koop_init['Юр лицо']=koop_init['Юр лицо'].str.strip()
                    
                    koop_init['Адрес магазина']=koop_init['Адрес магазина'].str.strip()
                    koop_init['Адрес магазина'] = koop_init['Адрес магазина'].replace(r'\s+', ' ', regex=True)
                    koop_init['Key_Update']=koop_init['Юр лицо']+koop_init['Адрес магазина']
                    koop_init["MONTH_DATE"]=str("01/"+init_m+"/2022")
                    koop_init["MONTH_DATE"]=pd.to_datetime(koop_init["MONTH_DATE"])
                    koop_init.MONTH_DATE = koop_init.MONTH_DATE.dt.strftime('%d/%m/%Y')
                    print("Successfully Read ",filename)
                    print("-----------------------")
                    # APPENDING ETL FILE WITH INITIAL DATA
                    ws = wb.sheets["Партнеры"]
                    ws.clear()
                    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = koop_init
                    tbl_range = ws.range("A1").expand('table')
                    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
                    ws.tables[0].name = "data"
                    print("Appended <Партнеры> Sheet")
                    print("-----------------------")
                    total_init_sum=koop_init['Продажи в пачках, шт.'].sum()
                    # MAPPING INITIAL DF WITH GEO
                    geo_mapping=koop_init.merge(geo_init, on='Key_Update', how='left')
                    geo_mapping=geo_mapping[geo_mapping['Код 1С Магазин'].isnull()]  
                    geo_mapping=geo_mapping.drop_duplicates(subset=['Key_Update'])
                    geo_mapping['Код 1С Магазин']=""
                    geo_mapping=geo_mapping[['Код 1С Магазин','Юр лицо_x','Адрес магазина_x']]
                    geo_mapping=geo_mapping.rename(columns={'Адрес магазина_x': 'Адрес магазина','Юр лицо_x':'Юр лицо'})
                    for i,item in enumerate( geo_mapping['Код 1С Магазин']):
                                geo_mapping['Код 1С Магазин'].iloc[i]="KOOP_"+str(max_koop_geo+i+1)
                    geo_mapping['Сity']=""
                    geo_mapping['State']="Nizhegorodskaya"
                    geo_mapping['Область']="Нижегородская обл."
                    geo_mapping['Region']="VOLGA"
                    geo_mapping['Долгота']=""
                    geo_mapping['Широта']=""
                    geo_mapping['Siebel Code']=""
                    geo_mapping['Branch']=""
                    geo_mapping=geo_mapping[['Код 1С Магазин','Юр лицо','Адрес магазина','Сity','State','Область','Region','Долгота','Широта','Siebel Code','Branch']]
                    ws = wb.sheets["Geo"]
                    ws["A"+str(rows+2)].options(pd.DataFrame, header=0, index=False, expand='table').value = geo_mapping
                    geo_new=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value
                    rows_new=geo_new.shape[0]
                    new_range="A"+str(rows+2)+":"+"A"+str(rows_new+1)
                    
                    
                    for a_cell in ws[new_range]:
                        a_cell.color = (0, 208, 142)
                    # REFRESHING EXCEL
                    wb.api.RefreshAll()
                    time.sleep(20)
                    # READING FOR LOAD DATA
                    ws = wb.sheets["For load"]
                    for_load=ws.range('A1').options(pd.DataFrame, 
                    header=1,
                    index=False, 
                    expand='table').value

                    for_load_total_sum=for_load['Продажи шт'].sum()

                    for_load=for_load[for_load['Код 1С Магазин'].isnull()]
                    blank_geos=for_load.shape[0]
                    if blank_geos>0:
                        print("MISSING SOME POS, PLEASE CHECK")
                        proceed=input("Willing to proceed?")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.api.RefreshAll()
                        time.sleep(20)
                        wb.save()
                    else:
                        print("GOOD JOB, No Blank POS Found in <For Load> Sheet_KOOP_RRP")
                        if for_load_total_sum==total_init_sum:
                            print("INITIAL & FOR LOAD <Продажи шт> ARE EQUAL :", for_load_total_sum)
                        else:
                            sg.popup("INITIAL & FOR LOAD <Продажи шт> ARE NOT EQUAL")
                        wb.save()
        except Exception as e:
            sg.popup("THERE WAS AN ERROR!")
            print("error-message----error-message----error-message")
            traceback.print_exception(e)
            print("error-message----error-message----error-message")
            error_msg=input("There was an ERROR in KOOP RRP, pls press any key to close the program") 



    sg.popup("TOTAL EXECUTION TIME ",time.time()-start_time)

if client=="X5":
    sg.theme('DarkBlue')

    layout = [
        [sg.T("Input Data Storage Folder:", s=45,justification="l"), sg.I(key="-FOLDER-"), sg.FolderBrowse()],
        [sg.Text("Input month/week num",s=45,justification="l"),sg.InputText( key='-DATE-')],
        [sg.Text("Choose data modification type",s=45,justification="l"),sg.Listbox(values=['RMC', 'RRP'], size=(45, 2), select_mode='single', key='-ETL-')],
        [sg.T("Input Main file with all data:",s=45, justification="l"), sg.I(key="-MAIN-"), sg.FileBrowse(file_types=(("Text Files", "*.txt"),))],
        [sg.T("Input Storage to Store New Data:", s=45,justification="l"), sg.I(key="-DESTINATION-"), sg.FolderBrowse()],
        [sg.Submit( )]
    ]

    window = sg.Window('Программа обработки продаж X5', layout)

    event, values = window.read()

    user_path=str(values['-FOLDER-'])
    selection=str(values['-ETL-'][0])
    month=str(values['-DATE-'])
    result_file=str(values['-MAIN-'])
    result_storage=str(values['-DESTINATION-'])
    window.close()

    all_files=glob.glob(user_path+"/*.xls*")
    li=[]