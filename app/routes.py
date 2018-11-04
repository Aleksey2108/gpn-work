# -*- coding: utf-8 -*-
from flask import render_template, flash, redirect, url_for, request
import os, datetime
from app import app, db
# from app.xls_gen import ProfTheEventXml
from app.function import GetDepartment, GetLastDay, CreateAuditTrailXls
from app.forms import AuditTrailForm, SelectDateRangeShort
from app.models import AuditTrail


@app.route('/' , methods=['GET', 'POST'])
def home():
     depart_id = 'null'
     strings = GetDepartment(depart_id)

     return render_template('home.html',  title='Home', strings = strings)

@app.route('/select_report' , methods=['GET', 'POST'])
def select_report():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     depart_name = GetDepartment(post)

     strings = [
         {
              'name' : depart_name,
              'url' : '/'
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
              'name' : u'Журнал админ. практики',
              'url' : '/journal-admp'
         },
     ]

   return render_template('select_report.html',  title='Select Report', depart_id = post , strings = strings)

@app.route('/ongoz', methods=['GET', 'POST'])
def ongoz():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }   
     ]
  
   return render_template('report.html',  title='Ongoz', depart_id = post , strings = strings)


@app.route('/single' , methods=['GET', 'POST'])
def single():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:

#     file = ProfTheEventXml()
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }   
     ]
  
   return render_template('report.html',  title='Single', depart_id = post , strings = strings)

@app.route('/weekly' , methods=['GET', 'POST'])
def weekly():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }
     ]

   return render_template('report.html',  title='weekly', depart_id = post , strings = strings)

@app.route('/monthly' , methods=['GET', 'POST'])
def monthly():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }
     ]

   return render_template('report.html',  title='monthly', depart_id = post , strings = strings)

@app.route('/quarterly' , methods=['GET', 'POST'])
def quarterly():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }
     ]

   return render_template('report.html',  title='quarterly', depart_id = post , strings = strings)

@app.route('/semi-annual' , methods=['GET', 'POST'])
def semi_annual():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }
     ]

   return render_template('report.html',  title='semi-annual', depart_id = post , strings = strings)

@app.route('/annual' , methods=['GET', 'POST'])
def annual():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }
     ]

   return render_template('report.html',  title='Annual', depart_id = post , strings = strings)

@app.route('/audit-trail' , methods=['GET', 'POST'])
def audit_trail():

   try:
#      request.form['depart_id']
      request.args['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
#     post = request.form['depart_id']
     depart_id = request.args['depart_id']
     strings = [
           {
                 'name' : u'Добавить запись',
                 'url' : url_for('add_audit_trail'),
                 'method' : 'get'                     
           },
           {
                 'name' : u'Скачать',
                 'url' : url_for('load_audit_trail'),
                 'method' : 'get'                    
           },
           {
                 'name' : u'Назад',
                 'url' : url_for('select_report'),
                 'method' : 'post'  
           }

     ]

     return render_template('report_action.html',  title='Audit trail', depart_id = depart_id , strings = strings)

@app.route('/add-audit-trail', methods=['GET', 'POST'])
def add_audit_trail():
#   try:
#       request.form['depart_id']
#   except KeyError:
#      return redirect(url_for('home'))
#   else:
#     depart_id = request.form['depart_id']

     print 'request.method ='+request.method

     form = AuditTrailForm()
     if request.method == 'POST':
       try:
         request.form['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.form['depart_id']
         print 'POST depart_id ='+depart_id
     elif request.method == 'GET':
       try:
         request.args['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.args['depart_id']
         print 'GET depart_id = '+depart_id

     if form.validate_on_submit():

             objectname = form.objectname.data
             print 'objectname = '+objectname

             add_data = AuditTrail (
               objectname = form.objectname.data,
               objectadres = form.objectadres.data,
               depart_id = depart_id,
               checkdate = form.checkdate.data,
               of_violations = form.of_violations.data,
               of_violations_unscheduled = form.of_violations_unscheduled.data,
               name_employee = u'Иванов Иван',
               other_documents =  form.other_documents.data,
               check_number =  form.check_number.data,
             )
             db.session.add(add_data)
             db.session.commit()

             flash(u'Запись сохранена.')
             return redirect(url_for('audit_trail')+'?depart_id='+ depart_id)
     return render_template('audit_trail_form.html',  title='Audit trail', depart_id = depart_id , form = form)

@app.route('/load_audit_trail', methods=['GET', 'POST'])
def load_audit_trail():

   form = SelectDateRangeShort()
   if request.method == 'POST':
      try:
         request.form['depart_id']
      except KeyError:
         return redirect(url_for('home'))
      else:
         depart_id = request.form['depart_id']
         print 'POST depart_id ='+depart_id
   elif request.method == 'GET':
      try:
         request.args['depart_id']
      except KeyError:
         return redirect(url_for('home'))
      else:
         depart_id = request.args['depart_id']
         print 'GET depart_id = '+depart_id


   if form.validate_on_submit():
      start_date = datetime.date(int(form.year_start.data), int(form.month_start.data) , 1)
      end_date = datetime.date(int(form.year_end.data), int(form.month_end.data) , GetLastDay(int(form.month_end.data),  int(form.year_end.data)))

      if end_date < start_date:
         error = u'Конечная дата не может быть меньше начальной!'
         print 'error'
         return render_template('select_date_range.html',  title='Audit trail', depart_id = depart_id , error = error, form = form)
      result = CreateAuditTrailXls(start_date, end_date)
      if result == 'error':
        error = u'Нет данных за выбранный период'
      return redirect(url_for('audit_trail')+'?depart_id='+ depart_id)
   return render_template('select_date_range.html',  title='Audit trail', depart_id = depart_id , form = form)

@app.route('/object-log' , methods=['GET', 'POST'])
def object_log():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }
     ]

   return render_template('report.html',  title='Object Log', depart_id = post , strings = strings)

@app.route('/journal-admp' , methods=['GET', 'POST'])
def journal_admp():

   try:
      request.form['depart_id']
   except KeyError:
      return redirect(url_for('home'))
   else:
     post = request.form['depart_id']

     strings = [
         {
              'name' : u'Назад',
              'url' : url_for('select_report'),
              'method' : 'post'  
         }
     ]

   return render_template('report.html',  title='Journal admin', depart_id = post , strings = strings)
