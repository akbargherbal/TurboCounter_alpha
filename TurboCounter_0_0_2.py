import xlsxwriter as xl
import subprocess
import regex as re
import os
import PySimpleGUI as sg


# HLPER FUNCTION: Get Time Now:
from datetime import datetime
def time_now():
    '''Get Current Time'''
    
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    print("Current Time =", current_time)
    return now



# HELPER FUNCTION: Remove Unwanted Characters:
pat_ar_msa = re.compile(r'[^ءاأإآؤئبةتثجحخدذرزسشصضطظعغفقكلمنهوىي\s]')
pat_en = re.compile(r'[^a-zA-ZÀ-ÿ-\s]')
pat_depunct = re.compile(r'[\\…*?+ـ،¦\.\:(؟°@=!؛><[)}/َ{;\'~_,—\"•\]\d»«]') 

def clean_word(word, category):
    '''Given a word; clean it according to 3 categories:
    1) Modern Standard Arabic MSA (Default)
    2) English
    3) Remove Punctuations >>>> Possibly add a choice where some characters can be removed or added.'''

    
    categories = set(list(range(1,4)))
    
    if category in categories:
        if category == 1: # Default English & European
            result = re.sub(pat_en, '', word)
            return result.lower()
        elif category == 2: # Arabic MSA
            result = re.sub(pat_ar_msa, '', word)
            return result
        else:
            result = re.sub(pat_depunct, '', word)
            return result.lower()
    else:
        print('ERROR! Choose between category 1,2 and 3!')


# Generator Function
def read_yield_txt(folder):
    '''
    A generator function that yields a word at a time from TXT file.
    '''
    list_txt_files = [file for file in os.listdir(folder) if file[-4:].lower()=='.txt']
    if list_txt_files:
        for file in list_txt_files:
            file_abs = os.path.join(os.path.abspath(f'{folder}/'), file)
            print(f'Processing {file} ...')
            with open(f'{file_abs}', encoding='utf-8') as file_01:
                for line in file_01.readlines():
                    for word in line.split():
                        yield word

# Iterator Function
def read_gen(folder, category):
    '''
    Takes in a word (string) and yields a dictionary.
    '''

    read_txt = read_yield_txt(folder)

    dict_count = {}

    for word in read_txt:
        word = clean_word(word, category)
        if len(word):
            if word in dict_count:
                dict_count[word] += 1
            else:
                dict_count[word] = 1

    dict_count = dict(sorted(dict_count.items(), key=lambda item: item[1], reverse=True))
    return dict_count

# Dictionary to Excle File
def dictoxl(dict_in, file_name_op):
    '''
    Given a simple dictionary of k:v;
    Create an xlsx file; containing k in column A; v in Column B.
    Word - Freq for example.
    '''

    xl_file = f'{file_name_op}.xlsx'
    workbook = xl.Workbook(xl_file)
    worksheet = workbook.add_worksheet()
    row, col = 0,0

    worksheet.write(0,0, 'WORD')
    worksheet.write(0,1, 'FREQ')
    for k,v in dict_in.items():
        row += 1
        worksheet.write(row, col, k)
        worksheet.write(row, col + 1, v)
    workbook.close()


# Main Function: count_words
def count_words(folder, file_name_op, category):
    if folder and file_name_op and category:
        ############
        print('\nStarting...')
        start = time_now()

        ############
        dict_count_final = read_gen(folder, category)

        dictoxl(dict_count_final, file_name_op)
        ############
        print('Finished...')
        end = time_now()
        ############
        duration = end - start
        duration_min = round(duration.seconds/60, 3)
        if duration_min < 2:
            time_unit = 'minute'
        else:
            time_unit = 'minutes'
        total_words,unique_words = sum([v for v in dict_count_final.values()]), len(dict_count_final)
        print(f'Total Words Count: {total_words:,} Words')
        print(f'Number of Unique Words: {unique_words:,} Words')
        print(f'Total duration is {duration_min} {time_unit}.')
        print('_'*80, '\n')
        return 'SUCCESS'
    else:
        print("You either didn't select files or didn't enter result file name!!!\n")

en = 'English words only.'
ar_msa = 'Arabic words without diacritics (Harakat).'
punc = 'Remove punctuations and special characters.'
layout =    [
    [sg.Text("Choose folder:",  font='Courier 14', auto_size_text=True, size=(22,1), justification='right'), sg.I(key='-INPUT-', size=(55,1)), sg.FolderBrowse()],
    [sg.T("Choose output location:",  font='Courier 14', auto_size_text=True,size=(22,1)), sg.I(key='-OUTPUT-',size=(55,1)), sg.FileSaveAs()],
    [sg.B("Start!", key='-START-', font='Courier 12'), 
        sg.Radio('English', group_id='lang', tooltip=en, key='-ENGLISH-'), 
        sg.Radio('Arabic MSA', group_id='lang', tooltip=ar_msa, key='-ARABIC_MSA-'), 
        sg.Radio('Remove Punct.', group_id='lang', tooltip=punc, key='-DEPUNC-')],
    [sg.Output(size=(80, 18), font='Calibri 14')]        
            ]
window = sg.Window('TurboCounter by Akbar Gherbal', layout, size=(775,450))
sg.Input()

while True:
    event, values = window.read()
    if event is None:
        break
  
    if event == '-START-':
        if values['-ENGLISH-']:
            category = 1
        elif values['-ARABIC_MSA-']:
            category = 2
        else:
            category = 3
        
        folder, file_name_op, category = values['-INPUT-'], values['-OUTPUT-'], category
        operation = count_words(folder, file_name_op, category)
        if operation == 'SUCCESS':
            pop_file = (file_name_op + '.xlsx').replace('/', '\\')
            subprocess.Popen(f'explorer /select,{pop_file}')
window.close()


