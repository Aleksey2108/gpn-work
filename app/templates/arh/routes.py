# -*- coding: utf-8 -*-
from flask import render_template, flash, redirect, url_for, request
from openpyxl import load_workbook
import os
from app import app



@app.route('/' , methods=['GET', 'POST'])
def home():

   return render_template('home.html',  title='Home')

@app.route('/select_report' , methods=['GET', 'POST'])
def select_report():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

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
         {
              'name' : u'Назад',
              'url' : '/'
         },
     ]

   return render_template('select_report.html',  title='Select Report', depart_id = post , strings = strings)

@app.route('/ongoz', methods=['GET', 'POST'])
def ongoz():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }   
     ]
  
   return render_template('report.html',  title='Ongoz', depart_id = post , strings = strings)


@app.route('/single' , methods=['GET', 'POST'])
def single():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }   
     ]
  
   return render_template('report.html',  title='Single', depart_id = post , strings = strings)

@app.route('/weekly' , methods=['GET', 'POST'])
def weekly():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }
     ]

   return render_template('report.html',  title='weekly', depart_id = post , strings = strings)

@app.route('/monthly' , methods=['GET', 'POST'])
def monthly():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }
     ]

   return render_template('report.html',  title='monthly', depart_id = post , strings = strings)

@app.route('/quarterly' , methods=['GET', 'POST'])
def quarterly():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }
     ]

   return render_template('report.html',  title='quarterly', depart_id = post , strings = strings)

@app.route('/semi-annual' , methods=['GET', 'POST'])
def semi_annual():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }
     ]

   return render_template('report.html',  title='semi-annual', depart_id = post , strings = strings)

@app.route('/annual' , methods=['GET', 'POST'])
def annual():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }
     ]

   return render_template('report.html',  title='Annual', depart_id = post , strings = strings)

@app.route('/audit-trail' , methods=['GET', 'POST'])
def audit_trail():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }
     ]

   return render_template('report.html',  title='Audit trail', depart_id = post , strings = strings)

@app.route('/object-log' , methods=['GET', 'POST'])
def object_log():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }
     ]

   return render_template('report.html',  title='Object Log', depart_id = post , strings = strings)

@app.route('/journal-admp' , methods=['GET', 'POST'])
def journal_admp():

   try:
      request.form['depart_id']
   except NameError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report')
         }
     ]

   return render_template('report.html',  title='Journal admin', depart_id = post , strings = strings)
