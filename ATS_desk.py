#[ ] 127 - all_max_costs_ids; all_max_costs - НЕ ОЧИЩАЮТСЯ ПЕРЕД ДОБАВЛЕНИЕМ НОВЫХ ДАННЫХ - ОЧИСТИТЬ?

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog
from calendar import monthrange
from tkinter.messagebox import showerror
from ids_values import ids_val_arr as ids_val_arr
import requests
import os
from collections import Counter
from bs4 import BeautifulSoup

# Создаем корневой объект - окно
win = tk.Tk() 
win.title('АТС-расчет') 
win.geometry('900x650') 
# icon = tk.PhotoImage(file='calendar.ico')
# win.iconphoto(False, icon)

#Определяем переменные
month_num = {"Январь": 1, "Февраль": 2, "Март ": 3, "Апрель": 4, "Май": 5, "Июнь": 6, "Июль": 7, "Август": 8, "Сентябрь": 9, "Октябрь": 10, "Ноябрь": 11, "Декабрь": 12}
ids_values = ids_val_arr
ids_to_check = []
ids_to_check_var = tk.Variable(value=ids_to_check) 
year_var = tk.StringVar() #year
month_var = tk.StringVar() #month
entry_folder_var = tk.StringVar() 

download_links = []
file_paths = []
all_max_costs = []
all_max_costs_var = tk.IntVar(value=all_max_costs)
all_max_costs_ids = []
max_cost_id_counts = {}
max_cost_dict_arr = []

#Функции
#Функция вызываемая при смене месяца или года в выпадающих списках
def update_dates():
    year = int(year_var.get())
    month = month_var.get()

    # Число дней в выбранном месяца года
    _, num_days = monthrange(year, month_num[month])

    # Возвращает массив номеров дней месяца
    dates = [str(day).zfill(2) for day in range(1, num_days + 1)] # массив дней или month_day_arr (str)
    return dates

#Включение кнопок Рассчитать, Сохранить и Очистить
def update_button_state(*args):
    if year_var.get() and month_var.get() and nodes_listbox.size() != 0 and entry_folder_var.get() != '':
        calculate_button['state'] = tk.NORMAL
    else:
        calculate_button['state'] = tk.DISABLED

def update_button_state_1(*args):
    if all_max_costs_var:
        save_button['state'] = tk.NORMAL
        clear_button['state'] = tk.NORMAL
    else:
        save_button['state'] = tk.DISABLED 
        clear_button['state'] = tk.DISABLED

year_var.trace('w', update_button_state)
month_var.trace('w', update_button_state)
entry_folder_var.trace('w', update_button_state)
ids_to_check_var.trace('w', update_button_state)
all_max_costs_var.trace('w', update_button_state_1)

# Кнопки + и -
def add():
    new_id = spinbox.get()
    if int(new_id) in ids_values and not int(new_id) in ids_to_check:
        ids_to_check.append(int(new_id))
        ids_to_check_var.set(ids_to_check)

def delete():
    selected_index = nodes_listbox.curselection() #удаление по индексу
    if selected_index:
        ids_to_check.pop(selected_index[0])
        ids_to_check_var.set(ids_to_check)
    
#Куда сохранять и откуда считать скачанные xls-файлы
def choose():
    folder_path = filedialog.askdirectory()
    entry_folder.delete(0, tk.END)
    entry_folder.insert(0, folder_path)

#Расчет и скачивание 
def calculate():
    # Массив ссылок для скачивания и пути сохранения скачанных xls-файлов
    for day_num in update_dates():
        file_link = 'https://www.atsenergo.ru/nreport?rname=big_nodes_prices_pub&rdate=' + str(year_var.get()) + str(month_num[month_var.get()]).zfill(2) + day_num # 'YYYYMMDD'
        response = requests.get(file_link, verify=False)
        soup = BeautifulSoup(response.text, 'html.parser') #response_souped   
        if(soup.find(id='aid_files_list').find('a')):
            download_link = "https://www.atsenergo.ru/nreport" + soup.find(id='aid_files_list').find('a').get('href') 
            file_path = entry_folder.get() + '/' + str(year_var.get()) + str(month_num[month_var.get()]).zfill(2) + day_num + '.xls'
            download_links.append(download_link) 
            file_paths.append(file_path) 

    # Скачивание xls по ссылке (с проверкой наличия ранее скачанных xls-файлов)
    if(download_links != []): #ИЗМЕНИТЬ!!
        for download_link, file_path in zip(download_links, file_paths):   
            if(not (str(year_var.get()) + str(month_num[month_var.get()]).zfill(2) + file_path.split('.')[0][-2:] + '.xls' in os.listdir(entry_folder.get()))):     
                # GET-запрос и сохранение файла (файлов) локально
                response_download = requests.get(download_link, verify=False)         
                with open(file_path, 'wb') as file:
                    file.write(response_download.content)  
            
            # Чтение xls-файла в словарь DataFrames
            data = pd.read_excel(file_path, usecols='A:F', header=2, sheet_name=list(range(24)), engine = "xlrd")

            max_cost_ids = {} 
            max_costs = {}  
            
            # Итерация по каждому листу (sheet) в словаре data
            for sheet, df in data.items():   
                max_cost_dict = {}
                
                filtered_df = df[df['Номер узла'].isin(ids_to_check)] # Оставляем строки по ID указанных ids_to_check
                max_cost_row = filtered_df.loc[filtered_df['Цена, руб'].idxmax()] # Определение строки с макс. ценой
                max_cost = max_cost_row['Цена, руб'] # ЦЕНА строки с макс. ценой
                max_cost_id = max_cost_row['Номер узла'] # ID строки с макс. ценой
                
                max_cost_dict['sheet'] = sheet
                max_cost_dict['day_num'] = file_path.split('.')[0][-2:]
                max_cost_dict['cost'] = max_cost                

                max_cost_ids[sheet] = max_cost_id # Добавление ID в max_cost_ids dict
                max_costs[sheet] = max_cost # Добавление ЦЕНЫ в max_costs dict
                
                max_cost_dict_arr.append(max_cost_dict)
            
            all_max_costs_ids.extend(list(max_cost_ids.values())) # массив max_cost_ids для 720ч
            all_max_costs.extend(list(max_costs.values())) # массив max_costs для 720ч

        max_cost_id_counts_new = dict(Counter(all_max_costs_ids))
        max_cost_id_counts.update(max_cost_id_counts_new)

        #Отображение all_max_costs [] в Listbox "costs_values_listbox"
        for max_cost in all_max_costs:
            costs_values_listbox.insert(tk.END, max_cost)

        #Отображение "max_cost_id_counts {}" в ListBox "Cost_ID_dict_listbox"
        for key, value in max_cost_id_counts.items():
            item_text = f"Узел: {key}, Число: {value}"
            Cost_ID_dict_listbox.insert(tk.END, item_text)        
       
    else:
        showerror(title="Ошибка", message="ДАННЫЕ ПО ВЫБРАННОЙ ДАТЕ ОТСУТСТВУЮТ!")

    download_links.clear()
    file_paths.clear()

#Сохранение результатов в блокнот
def save():
    folder_path_result = filedialog.askdirectory()
    if folder_path_result:
        # Save all_max_costs to a file named "Макс цены.xlsx"
        df = pd.DataFrame(max_cost_dict_arr)
        pivot_df = df.pivot(index='sheet', columns='day_num', values='cost') # Pivot the DataFrame to have 'day_num' as columns, 'sheet' as index, and 'cost' as cell data
        output_file = f"{folder_path_result}/Макс цены_{str(year_var.get()) + str(month_num[month_var.get()]).zfill(2)}.xlsx" # # Construct the output file path
        pivot_df.to_excel(output_file)
        print("File saved successfully.")

        # Save max_cost_id_counts to a file named "Статистика по узлам.txt"
        max_ids_file_path = f"{folder_path_result}/Статистика по узлам_{str(year_var.get()) + str(month_num[month_var.get()]).zfill(2)}.txt"
        with open(max_ids_file_path, 'w') as file:
            for key, value in max_cost_id_counts.items():
                file.write(f"{key}: {value}\n")

#Очистить listbox
def clear_list():
    all_max_costs = []
    all_max_costs_var.set(all_max_costs)
    Cost_ID_dict_listbox.delete(0, tk.END)

# Сетка элементов
# Узлы
nodes_label = tk.Label(win, text="Номера узлов")
nodes_label.grid(row=0, column=0, padx=10, pady=(10, 0), sticky=tk.W)

spinbox_var = tk.IntVar(value = 100001)
spinbox = ttk.Spinbox(win, textvariable=spinbox_var, values=ids_values)
spinbox.grid(row=1, column=0, padx=10, pady=(0, 10), sticky=tk.W)

entry_folder_label = tk.Label(win, text = "Папка с сохраненными файлами")
entry_folder_label.grid(row=0, column=1, padx=10, pady=(10, 0), sticky=tk.W)
btn_choose_folder = ttk.Button(win, text="Выбрать", command=choose)
btn_choose_folder.grid(row=1, column=2, padx=10, pady=(0, 10), sticky=tk.W)
entry_folder = ttk.Entry(win, textvariable=entry_folder_var, width = 20)
entry_folder.grid(row=1, column=1, padx=10, pady=(0, 10), sticky=tk.W)

btn_add_node = ttk.Button(win, text='+', width=3, command=add)
btn_add_node.grid(row=2, column=0, padx=10, pady=(0, 10), sticky=tk.W)

btn_remove_node = ttk.Button(win, text='-', width=3, command=delete)
btn_remove_node.grid(row=2, column=0, padx=50, pady=(0, 10), sticky=tk.W)

nodes_listbox_label = tk.Label(win, text = "Выбранные узлы")
nodes_listbox_label.grid(row=3, column=0, padx=10, pady=(10, 0), sticky=tk.W)
nodes_listbox = tk.Listbox(win, listvariable=ids_to_check_var)
nodes_listbox.grid(row=4, column=0, padx=10, pady=(0, 10), sticky=tk.W)

# Выбор даты
date_label = tk.Label(win, text="Год и месяц")
date_label.grid(row=3, column=1, padx=10, pady=(10, 0), sticky=tk.W)

year_dropdown = ttk.Combobox(win, textvariable=year_var, state="readonly", values=list(range(2020, 2051)))
year_dropdown.grid(row=4, column=1, padx=10, sticky=tk.NW)
year_dropdown.bind("<<ComboboxSelected>>", lambda event: update_dates())

month_dropdown = ttk.Combobox(win, textvariable=month_var, values=["Январь", "Февраль", "Март ", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"], state="readonly",)
month_dropdown.grid(row=4, column=1, padx=10, sticky=tk.W)
month_dropdown.bind("<<ComboboxSelected>>", lambda event: update_dates())

# Работа с данными
calculate_button = tk.Button(win, text="Рассчитать", command=calculate, state=tk.DISABLED)
calculate_button.grid(row=4, column=1, padx=10, pady=0, sticky=tk.SW)

clear_button = ttk.Button(win, text="Очистить", command=clear_list, state=tk.DISABLED)
clear_button.grid(row=4, column=2, padx=10, pady=0, sticky=tk.SW)

save_button = ttk.Button(win, text="Сохранить", command=save, state=tk.DISABLED)
save_button.grid(row=4, column=3, padx=10, pady=0, sticky=tk.SW)

costs_values_label = tk.Label(win, text="Максимальные цены")
costs_values_label.grid(row=5, column=0, padx=10, pady=(10, 0), sticky=tk.W)
costs_values_listbox = tk.Listbox(win, listvariable=all_max_costs_var)
costs_values_listbox.grid(row=6, column=0, padx=10, pady=(0, 10), sticky=tk.W)

Cost_ID_dict_label = tk.Label(win, text="Статистика по узлам")
Cost_ID_dict_label.grid(row=5, column=1, padx=10, pady=(10, 0), sticky=tk.W)
Cost_ID_dict_listbox = tk.Listbox(win)
Cost_ID_dict_listbox.grid(row=6, column=1, padx=10, pady=(0, 10), sticky=tk.W)

#Запуск 
win.mainloop()

