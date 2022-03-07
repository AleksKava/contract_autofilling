#!/usr/bin/env python
# coding: utf-8

# ## Download all the necessary libraries

# In[ ]:


import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
import beepy as beep
import os
import re
import easygui
import shutil
import time
import warnings
os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"
from pygame import mixer


# In[ ]:


# from IPython.core.display import display, HTML
# display(HTML("<style>.container { width:100% !important; }</style>"))
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
pd.options.mode.chained_assignment = None  # default='warn'


# In[ ]:


print('Contract is being created...')


# ## Create all the necessary functions

# In[ ]:


#Функция для создания числа прописью
import decimal
units = (
    u'ноль',
    (u'один', u'одна'),
    (u'два', u'две'),
    u'три', u'четыре', u'пять',
    u'шесть', u'семь', u'восемь', u'девять'
)
teens = (
    u'десять', u'одиннадцать',
    u'двенадцать', u'тринадцать',
    u'четырнадцать', u'пятнадцать',
    u'шестнадцать', u'семнадцать',
    u'восемнадцать', u'девятнадцать'
)
tens = (
    teens,
    u'двадцать', u'тридцать',
    u'сорок', u'пятьдесят',
    u'шестьдесят', u'семьдесят',
    u'восемьдесят', u'девяносто'
)
hundreds = (
    u'сто', u'двести',
    u'триста', u'четыреста',
    u'пятьсот', u'шестьсот',
    u'семьсот', u'восемьсот',
    u'девятьсот'
)
orders = (# plural forms and gender
    #((u'', u'', u''), 'm'), # ((u'рубль', u'рубля', u'рублей'), 'm'), # ((u'копейка', u'копейки', u'копеек'), 'f')
    ((u'тысяча', u'тысячи', u'тысяч'), 'f'),
    ((u'миллион', u'миллиона', u'миллионов'), 'm'),
    ((u'миллиард', u'миллиарда', u'миллиардов'), 'm'),
)
minus = u'минус'

def thousand(rest, sex):
    """Converts numbers from 19 to 999"""
    prev = 0
    plural = 2
    name = []
    use_teens = rest % 100 >= 10 and rest % 100 <= 19
    if not use_teens:
        data = ((units, 10), (tens, 100), (hundreds, 1000))
    else:
        data = ((teens, 10), (hundreds, 1000))
    for names, x in data:
        cur = int(((rest - prev) % x) * 10 / x)
        prev = rest % x
        if x == 10 and use_teens:
            plural = 2
            name.append(teens[cur])
        elif cur == 0:
            continue
        elif x == 10:
            name_ = names[cur]
            if isinstance(name_, tuple):
                name_ = name_[0 if sex == 'm' else 1]
            name.append(name_)
            if cur >= 2 and cur <= 4:
                plural = 1
            elif cur == 1:
                plural = 0
            else:
                plural = 2
        else:
            name.append(names[cur-1])
    return plural, name

def num2text(num, main_units=((u'', u'', u''), 'm')):
    """
    http://ru.wikipedia.org/wiki/Gettext#.D0.9C.D0.BD.D0.BE.D0.B6.D0.B5.D1.81.\
    D1.82.D0.B2.D0.B5.D0.BD.D0.BD.D1.8B.D0.B5_.D1.87.D0.B8.D1.81.D0.BB.D0.B0_2
    """
    _orders = (main_units,) + orders
    if num == 0:
        return ' '.join((units[0], _orders[0][0][2])).strip() # ноль
    rest = abs(num)
    ord = 0
    name = []
    while rest > 0:
        plural, nme = thousand(rest % 1000, _orders[ord][1])
        if nme or ord == 0:
            name.append(_orders[ord][0][plural])
        name += nme
        rest = int(rest / 1000)
        ord += 1
    if num < 0:
        name.append(minus)
    name.reverse()
    return ' '.join(name).strip()


# In[ ]:


# Create function for deleting paragraphs
def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


# In[ ]:


# Create function to delete rows from tables:
def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)


# In[ ]:


# Create function for sounds making
def beep(sound):
    mixer.init() 
    sound=mixer.Sound(os.getcwd() + '/System/Sounds/{}.wav'.format(sound))
    sound.play()


# ## Download and preprocess source files

# In[ ]:


try:
    # Load the data for contract filling
    data = pd.read_excel(os.getcwd() + '/Data.xlsx', usecols = [0,1])
    # Correct first payment's sum
    if np.isnan(float(data['Значение'][6])) == True:
        data['Значение'][6] = 0
        
except BaseException  as e:
    beep(4)
    easygui.msgbox('Failed to upload file "Data": \n' + str(e), title='Error!')
    logger.error(str(e))
    sys.exit()


# In[ ]:


try:
    # Download Contract template
    c = DocxTemplate(os.getcwd() + '/System/Template/Contract.docx')

except BaseException  as e:
    beep(4)
    easygui.msgbox('Failed to upload Contract template: \n' + str(e), title='Error!')
    logger.error(str(e))
    sys.exit()


# ## Create an agreement

# In[ ]:


try:
    # Put the data into the template
    context = {'contractor' : data['Значение'][0],
               'contract_number' : data['Значение'][1],
               'date' : data['Значение'][2],
               'start_date' : data['Значение'][3],
               'end_date' : data['Значение'][4],
               'contract_sum' : data['Значение'][5],
               'contract_sum_words' : num2text(float(data['Значение'][5])),
               'first_sum' : data['Значение'][6],
               'first_sum_words' : num2text(int(data['Значение'][6])),
               'second_sum' : str(int(data['Значение'][5]) - int(data['Значение'][6])),
               'second_sum_words' : num2text(float(data['Значение'][5]) - float(data['Значение'][6])),           
               'address' : data['Значение'][7],
               'client' : data['Значение'][10],
               'pass' : data['Значение'][11],
               'authority' : data['Значение'][12],
               'pass_date' : data['Значение'][13],
               'code' : data['Значение'][14],
               'registration_address' : data['Значение'][15]}
    c.render(context)

    # If executor is IP
    if data['Значение'][0] == 'ИП Рабинович':
        # Find the paragraph about RAV
        indx=0
        for p in c.paragraphs:
            if p.text.startswith('Рабинович Ариэль Вальдемарович, в лице'):
                aaindx = indx
            indx+=1
        # And delete it
        delete_paragraph(c.paragraphs[aaindx])
        # Remove corresponding row in contractor's data table
        remove_row(c.tables[0], c.tables[0].rows[2])

    # If executor is RAV
    if data['Значение'][0] == 'Рабинович Ариэль Вальдемарович':
        # Find the paragraph about IP
        indx=0
        for p in c.paragraphs:
            if p.text.startswith('Индивидуальный предприниматель Рабинович,'):
                ipindx = indx
            indx+=1
        # And delete it
        delete_paragraph(c.paragraphs[ipindx])
        # Remove corresponding row in contractor's data table
        remove_row(c.tables[0], c.tables[0].rows[3])

    # If there is no advance payment
    if float(data['Значение'][6]) == 0:
        # Find the paragraph about the advance payment
        indx=0
        for p in c.paragraphs:
            if p.text.startswith('Заказчик оплачивает аванс в размере '):
                avindx = indx
            indx+=1
        # And delete it
        delete_paragraph(c.paragraphs[avindx])
    
except BaseException  as e:
    beep(4)
    easygui.msgbox('Failed to create file "Contract": \n' + str(e), title='Error!')
    logger.error(str(e))
    sys.exit()


# ## Create an agreement data file

# In[ ]:


try:
    # Create new dataframe for saving in excel
    newdata = data.drop(data.index[[9, 11, 12, 13, 14, 15]])
    newdata['Значение'] = newdata.Значение.astype(str)

except BaseException  as e:
    beep(4)
    easygui.msgbox('Failed to create file "Contract data": \n' + str(e), title='Error!')
    logger.error(str(e))
    sys.exit()


# ## Save created files

# In[ ]:


# Create new folder for saving contract and it's data
 # Get contract number
cnr = data['Значение'][1].replace('/', '_')
 # Get client name
cname = ''.join(data['Значение'][10].split()[:2])
 # Get apartment number
apnr = re.findall(r'\d+', data['Значение'][7])[-1]
apnr
 # Get folder's name
fname = cnr+'_кв'+apnr+'_'+cname


# In[ ]:


# Check if the same folder already exists
 # Create choises
ch1 = 'Yes'
ch2 = 'No, try again'
ch3 = 'No, stop program execution'
ch = ch2 
 # Check for choise     
while ch == ch2:
    if os.path.isdir(os.getcwd() + '/{}'.format(fname)):
        beep(4)
        q = easygui.buttonbox('A folder with the same name already exists, do you want to replace it?', 'Attention!!!', (ch1, ch2, ch3))
        if q == ch1:
            ch = ch1
            shutil.rmtree(os.getcwd() + '/{}'.format(fname))
        elif q == ch3:
            ch = ch3
            sys.exit()
    else:
        ch = 0


# In[ ]:


# Create the folder
os.mkdir(os.getcwd() + '/{}'.format(fname))


# In[ ]:


try:
    # Save newdata to excel file
     # Get file's name
    xname = f"Contract_data_{data['Значение'][1].replace('/', '_')}"
     # Save it
    writer = pd.ExcelWriter(os.getcwd() + '/{}/{}.xlsx'.format(fname, xname), engine='xlsxwriter')
    newdata.to_excel(writer, sheet_name='ContractData', index=False)

    def format_col_width(ws):
        ws.set_column('A:A', 35)
        ws.set_column('B:B', 120)

    workbook  = writer.book
    worksheet = writer.sheets['ContractData']
    format_col_width(worksheet)

    writer.close()

except BaseException  as e:
    beep(4)
    easygui.msgbox('Failed to save file "Contract data": \n' + str(e), title='Error!')
    logger.error(str(e))
    sys.exit()


# In[ ]:


try:
    # Save contract to word file
     # Get file's name
    wname = f"Contract_{cnr}"    
    c.save(os.getcwd() + '/{}/{}.docx'.format(fname, wname))
    
except BaseException  as e:
    beep(4)
    easygui.msgbox('Failed to save file "Contract": \n' + str(e), title='Error!')
    logger.error(str(e))
    sys.exit()


# In[ ]:


# Add sound
if data['Значение'][8] > 0:
    beep(data['Значение'][8])


# In[ ]:


print(wname + ' has been drawn up!')
time.sleep(3)

