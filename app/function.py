# -*- coding: utf-8 -*-
from app import app


def GetDepartment(depart_id):
    depart = [
         {
              'name' : u'УНПР',
              'id' : 'unpr'
         },
         {
              'name' : u'НТО',
              'id' : 'nto'
         },
         {
              'name' : u'ОНБП',
              'id' : 'onbp'
         },
         {
              'name' : u'ОНГОЗНТЧС',
              'id' : 'ongozntchs'
         },
         {
              'name' : u'ОГСД',
              'id' : 'ogsd'
         },
         {
              'name' : u'ООРД',
              'id' : 'oord'
         },
         {
              'name' : u'ОЛК',
              'id' : 'olk'
         },
         {
              'name' : u'ОНОВПО',
              'id' : 'onovpo'
         },
         {
              'name' : u'ОНТ',
              'id' : 'ont'
         },
         {
              'name' : u'ОПД',
              'id' : 'opd'
         },
         {
              'name' : u'РОНПР',
              'id' : 'ronpr'
         },
    ]
    if depart_id == 'null':
       return depart
    else:
       for val in depart:

         if val['id'] == depart_id:
           return val['name']

      