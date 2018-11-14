# -*- coding: utf-8 -*-
from app import app
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, Side
from datetime import  *


class ProfTheEventXml:

   font = Font(name='Calibri',
                       size=11,
                       bold=False,
                       italic=False,
                       vertAlign=None,
                       underline='none',
                       strike=False,
                       color='FF000000')    
   align_center=Alignment(horizontal='center',
                        vertical='bottom',
                       text_rotation=0,
                       wrap_text=False,
                       shrink_to_fit=False,
                       indent=0)

   number_format = 'General'
   protection = Protection(locked=True,
                        hidden=False)

   wb = Workbook()
   ws = wb.active


   ws.title = u'Лист1'


#u данные для строк
   rows = [
                  [u'Название', u'Язык', u'Время'],
                  ['Ivan', 'PHP', 123],
                  ['Egor', 'Python', 123],
              ]

   for row in rows:
        ws.append(row)
  
#   os.chdir(r'app\tmp')
   wb.save("sample.xlsx")

#   os.startfile(r'sample.xlsx')
#   os.remove(r'sample.xlsx')
#   os.chdir(r'')



