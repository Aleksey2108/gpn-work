# -*- coding: utf-8 -*-
from flask import render_template, flash, redirect, url_for
from openpyxl import load_workbook
import os
from app import app



@app.route('/')
@app.route('/ongoz')
def home():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       },
       {
            'name' : u'Разовые',
            'url' : '/single'
       },
       {
            'name' : u'Еженедельные',
            'url' : '/weekly'
       },
       {
            'name' : u'Ежемесячные',
            'url' : '/monthly'
       },
       {
            'name' : u'Ежеквартальные',
            'url' : '/quarterly'
       },
       {
            'name' : u'Полугодичные',
            'url' : '/semi-annual'
       },
       {
            'name' : u'Годовые',
            'url' : '/annual'
       },
       {
            'name' : u'Журнал проверок',
            'url' : '/audit-trail'
       },
       {
            'name' : u'Журнал объектов',
            'url' : '/object-log'
       },
       {
            'name' : u'Журнал административной практики',
            'url' : '/journal-admp'
       },

   ]

   return render_template('home.html',  title='Home' , strings = strings)


@app.route('/single')
def single():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }   
   ]
  
   return render_template('home.html',  title='Single', strings = strings)

@app.route('/weekly')
def weekly():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }
   ]

   return render_template('home.html',  title='weekly', strings = strings)

@app.route('/monthly')
def monthly():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }
   ]

   return render_template('home.html',  title='monthly', strings = strings)

@app.route('/quarterly')
def quarterly():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }
   ]

   return render_template('home.html',  title='quarterly', strings = strings)

@app.route('/semi-annual')
def semi_annual():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }
   ]

   return render_template('home.html',  title='semi-annual', strings = strings)

@app.route('/annual')
def annual():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }
   ]

   return render_template('home.html',  title='Annual', strings = strings)

@app.route('/audit-trail')
def audit_trail():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }
   ]

   return render_template('home.html',  title='Audit trail', strings = strings)

@app.route('/object-log')
def object_log():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }
   ]

   return render_template('home.html',  title='Object Log', strings = strings)

@app.route('/journal-admp')
def journal_admp():

   strings = [
       {
            'name' : u'ОНГОЗНТЧС',
            'url' : '/ongoz'
       }
   ]

   return render_template('home.html',  title='Journal admin', strings = strings)