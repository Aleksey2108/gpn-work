# -*- coding: utf-8 -*-
from flask import render_template, flash, redirect, url_for, request, send_from_directory
import os, datetime
from app import app, db
from app.function import GetDepartment, GetLastDay, CheckLastDay, CreateAuditTrailXls_CHS, CreateAuditTrailXls_GO, CreateAuditTrailXls_PB
from app.forms import AuditTrailForm, AuditTrailFormCHS, AuditTrailFormPB, AuditTrailFormGO, SelectDateRangeShort
from app.models import AuditTrail, AuditTrail_CHS, AuditTrail_GO, AuditTrail_PB
# from app.models import AuditTrail, AuditTrail_CHS


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
              'url' : '/single',
              'method' : 'get'
         },
         {
              'name' : u'Еженедельные',
              'url' : '/weekly',
              'method' : 'get'
         },
         {
              'name' : u'Ежемесячные',
              'url' : '/monthly',
              'method' : 'get'
         },
         {
              'name' : u'Ежеквартальные',
              'url' : '/quarterly',
              'method' : 'get'
         },
         {
              'name' : u'Полугодичные',
              'url' : '/semi-annual',
              'method' : 'get'
         },
         {
              'name' : u'Годовые',
              'url' : '/annual'
         },
         {
              'name' : u'Журнал проверок',
              'url' : '/audit-trail-sel',
              'method' : 'post'
         },
         {
              'name' : u'Журнал объектов',
              'url' : '/object-log',
              'method' : 'get'
         },
         {
              'name' : u'Журнал админ. практики',
              'url' : '/journal-admp',
              'method' : 'get'
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

@app.route('/audit-trail-sel' , methods=['GET', 'POST'])
def audit_trail_sel():

     if request.method == 'POST':
       try:
         request.form['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.form['depart_id']
     elif request.method == 'GET':
       try:
         request.args['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.args['depart_id']

     strings = [
           {
                 'name' : u'Гражданская оборона',
                 'url' : url_for('audit_trail'),
                 'method' : 'get',
                 'action':'go'                     
           },
           {
                 'name' : u'Пожарная безопасность',
                 'url' : url_for('audit_trail'),
                 'method' : 'get',   
                 'action':'pb'                        
           },
           {
                 'name' : u'Чрезвычайные ситуации',
                 'url' : url_for('audit_trail'),
                 'method' : 'get',  
                 'action':'chs'                  
           },
           {
                 'name' : u'Назад',
                 'url' : url_for('select_report'),
                 'method' : 'post'  
           }

     ]

     return render_template('report_action.html',  title='Audit trail', depart_id = depart_id , strings = strings)

@app.route('/audit-trail' , methods=['GET', 'POST'])
def audit_trail():

   try:
#      request.form['depart_id']
      request.args['depart_id'] and request.args['action']
   except KeyError:
      return redirect(url_for('home'))
   else:
#     post = request.form['depart_id']
     depart_id = request.args['depart_id']
     action = request.args['action']
     strings = [
           {
                 'name' : u'Добавить запись',
                 'url' : url_for('add_audit_trail_'+action),
                 'method' : 'get'                     
           },
           {
                 'name' : u'Скачать',
                 'url' : url_for('load_audit_trail_'+action),
                 'method' : 'get'                    
           },
           {
                 'name' : u'Назад',
                 'url' : url_for('audit_trail_sel'),
                 'method' : 'post'  
           }

     ]

     return render_template('report_action.html',  title='Audit trail', depart_id = depart_id , strings = strings)

@app.route('/add-audit-trail-go', methods=['GET', 'POST'])
def add_audit_trail_go():
#   try:
#       request.form['depart_id']
#   except KeyError:
#      return redirect(url_for('home'))
#   else:
#     depart_id = request.form['depart_id']

     form = AuditTrailFormGO()
     if request.method == 'POST':
       try:
         request.form['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.form['depart_id']
     elif request.method == 'GET':
       try:
         request.args['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.args['depart_id']

     if form.validate_on_submit():

             if form.type_inspection.data == '0':
               type_inspection_t = u'плановая'
             else: 
               type_inspection_t = u'внеплановая'

             add_data = AuditTrail_GO (
               objectname = form.objectname.data,
               objectadres = form.objectadres.data,
               depart_id = depart_id,
               doc_stored = GetDepartment(form.doc_stored.data),
               checkdate = datetime.datetime.now(),
               type_inspection = type_inspection_t,
               start_date = datetime.date(int(form.start_year.data), int(form.start_month.data) , CheckLastDay(int(form.start_day.data), int(form.start_month.data),  int(form.start_year.data))),
               end_date = datetime.date(int(form.end_year.data), int(form.end_month.data) , CheckLastDay(int(form.end_day.data), int(form.end_month.data),  int(form.end_year.data))),
               act_number = form.act_number.data,
               act_date = datetime.date(int(form.act_year.data), int(form.act_month.data) , CheckLastDay(int(form.act_day.data), int(form.act_month.data),  int(form.act_year.data))),
               order_number = form.order_number.data,
               order_date = datetime.date(int(form.order_year.data), int(form.order_month.data) , CheckLastDay(int(form.order_day.data), int(form.order_month.data),  int(form.order_year.data))),
               of_violations = form.of_violations.data,
               of_violations_unscheduled = form.of_violations_unscheduled.data,
               fixed_violations = form.fixed_violations.data,
               name_employee = form.name_employee.data,
               other_documents = form.other_documents.data,
               check_number =  form.check_number.data,
               depart_name = GetDepartment(depart_id),
             )
             db.session.add(add_data)
             db.session.commit()

             flash(u'Запись сохранена.')
#             return redirect(url_for('audit_trail_sel')+'?depart_id='+ depart_id)
             return redirect(url_for('add_audit_trail_go')+'?depart_id='+ depart_id)
     return render_template('audit_trail_GO_form.html',  title='Audit trail', depart_id = depart_id , form = form)

@app.route('/load-audit-trail-go', methods=['GET', 'POST'])
def load_audit_trail_go():

   form = SelectDateRangeShort()
   if request.method == 'POST':
      try:
         request.form['depart_id']
      except KeyError:
         return redirect(url_for('home'))
      else:
         depart_id = request.form['depart_id']
   elif request.method == 'GET':
      try:
         request.args['depart_id']
      except KeyError:
         return redirect(url_for('home'))
      else:
         depart_id = request.args['depart_id']

   if form.validate_on_submit():
      start_date = datetime.date(int(form.year_start.data), int(form.month_start.data) , 1)
      end_date = datetime.date(int(form.year_end.data), int(form.month_end.data) , GetLastDay(int(form.month_end.data),  int(form.year_end.data)))

      if end_date < start_date:
         error = u'Конечная дата не может быть меньше начальной!'
         print 'error'
         return render_template('select_date_range.html',  title='Audit trail', depart_id = depart_id , action = 'go', error = error, form = form)
      result = CreateAuditTrailXls_GO(start_date, end_date)
      if result == 'error':
        error = u'Нет данных за выбранный период'
      else:      
        return redirect(url_for('download_file', filename = result))
      return redirect(url_for('audit_trail')+'?depart_id='+ depart_id)
   return render_template('select_date_range.html',  title='Audit trail', depart_id = depart_id , action = 'go' , form = form)

@app.route('/add-audit-trail-pb', methods=['GET', 'POST'])
def add_audit_trail_pb():
 
     form = AuditTrailFormPB()
     if request.method == 'POST':
       try:
         request.form['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.form['depart_id']
     elif request.method == 'GET':
       try:
         request.args['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.args['depart_id']

     if form.validate_on_submit():

             if form.type_inspection.data == '0':
               type_inspection_t = u'плановая'
             else: 
               type_inspection_t = u'внеплановая'

             add_data = AuditTrail_PB (
               objectname = form.objectname.data,
               objectadres = form.objectadres.data,
               depart_id = depart_id,
               doc_stored = GetDepartment(form.doc_stored.data),
               checkdate = datetime.datetime.now(),
               type_inspection = type_inspection_t,
               start_date = datetime.date(int(form.start_year.data), int(form.start_month.data) , CheckLastDay(int(form.start_day.data), int(form.start_month.data),  int(form.start_year.data))),
               end_date = datetime.date(int(form.end_year.data), int(form.end_month.data) , CheckLastDay(int(form.end_day.data), int(form.end_month.data),  int(form.end_year.data))),
               act_number = form.act_number.data,
               act_date = datetime.date(int(form.act_year.data), int(form.act_month.data) , CheckLastDay(int(form.act_day.data), int(form.act_month.data),  int(form.act_year.data))),
               order_number = form.order_number.data,
               order_date = datetime.date(int(form.order_year.data), int(form.order_month.data) , CheckLastDay(int(form.order_day.data), int(form.order_month.data),  int(form.order_year.data))),
               of_violations = form.of_violations.data,
               of_violations_unscheduled = form.of_violations_unscheduled.data,
               fixed_violations = form.fixed_violations.data,
               name_employee = form.name_employee.data,
               other_documents = form.other_documents.data,
               check_number =  form.check_number.data,
               depart_name = GetDepartment(depart_id),
             )
             db.session.add(add_data)
             db.session.commit()

             flash(u'Запись сохранена.')
             return redirect(url_for('audit_trail_sel')+'?depart_id='+ depart_id)
     return render_template('audit_trail_PB_form.html',  title='Audit trail', depart_id = depart_id , form = form)

@app.route('/load-audit-trail-pb', methods=['GET', 'POST'])
def load_audit_trail_pb():

   form = SelectDateRangeShort()
   if request.method == 'POST':
      try:
         request.form['depart_id']
      except KeyError:
         return redirect(url_for('home'))
      else:
         depart_id = request.form['depart_id']
   elif request.method == 'GET':
      try:
         request.args['depart_id']
      except KeyError:
         return redirect(url_for('home'))
      else:
         depart_id = request.args['depart_id']


   if form.validate_on_submit():
      start_date = datetime.date(int(form.year_start.data), int(form.month_start.data) , 1)
      end_date = datetime.date(int(form.year_end.data), int(form.month_end.data) , GetLastDay(int(form.month_end.data),  int(form.year_end.data)))

      if end_date < start_date:
         error = u'Конечная дата не может быть меньше начальной!'
         print 'error'
         return render_template('select_date_range.html',  title='Audit trail', depart_id = depart_id , action = 'pb' , error = error,  form = form)
      result = CreateAuditTrailXls_PB(start_date, end_date)
      if result == 'error':
        error = u'Нет данных за выбранный период'
      else:      
        return redirect(url_for('download_file', filename = result))
      return redirect(url_for('audit_trail')+'?depart_id='+ depart_id)
   return render_template('select_date_range.html',  title='Audit trail', depart_id = depart_id ,  action = 'pb' ,  form = form)

@app.route('/add-audit-trail-chs', methods=['GET', 'POST'])
def add_audit_trail_chs():

     form = AuditTrailFormCHS()
     if request.method == 'POST':
       try:
         request.form['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.form['depart_id']
     elif request.method == 'GET':
       try:
         request.args['depart_id']
       except KeyError:
         return redirect(url_for('home'))
       else:
         depart_id = request.args['depart_id']

     if form.validate_on_submit():

             if form.type_inspection.data == '0':
               type_inspection_t = u'плановая'
             else: 
               type_inspection_t = u'внеплановая'

             add_data = AuditTrail_CHS (
               objectname = form.objectname.data,
               objectadres = form.objectadres.data,
               depart_id = depart_id,
               doc_stored = GetDepartment(form.doc_stored.data),
               checkdate = datetime.datetime.now(),
               type_inspection = type_inspection_t,
               start_date = datetime.date(int(form.start_year.data), int(form.start_month.data) , CheckLastDay(int(form.start_day.data), int(form.start_month.data),  int(form.start_year.data))),
               end_date = datetime.date(int(form.end_year.data), int(form.end_month.data) , CheckLastDay(int(form.end_day.data), int(form.end_month.data),  int(form.end_year.data))),
               act_number = form.act_number.data,
               act_date = datetime.date(int(form.act_year.data), int(form.act_month.data) , CheckLastDay(int(form.act_day.data), int(form.act_month.data),  int(form.act_year.data))),
               order_number = form.order_number.data,
               order_date = datetime.date(int(form.order_year.data), int(form.order_month.data) , CheckLastDay(int(form.order_day.data), int(form.order_month.data),  int(form.order_year.data))),
               of_violations = form.of_violations.data,
               of_violations_unscheduled = form.of_violations_unscheduled.data,
               fixed_violations = form.fixed_violations.data,
               name_employee = form.name_employee.data,
               check_number =  form.check_number.data,
               depart_name = GetDepartment(depart_id)
             )
             db.session.add(add_data)
             db.session.commit()

             flash(u'Запись сохранена.')
             return redirect(url_for('audit_trail_sel')+'?depart_id='+ depart_id)
     return render_template('audit_trail_CHS_form.html',  title='Audit trail', depart_id = depart_id , form = form)


@app.route('/load-audit-trail-chs', methods=['GET', 'POST'])
def load_audit_trail_chs():

   form = SelectDateRangeShort()
   if request.method == 'POST':
      try:
         request.form['depart_id']
      except KeyError:
         return redirect(url_for('home'))
      else:
         depart_id = request.form['depart_id']
   elif request.method == 'GET':
      try:
         request.args['depart_id']
      except KeyError:
         return redirect(url_for('home'))
      else:
         depart_id = request.args['depart_id']


   if form.validate_on_submit():
      start_date = datetime.date(int(form.year_start.data), int(form.month_start.data) , 1)
      end_date = datetime.date(int(form.year_end.data), int(form.month_end.data) , GetLastDay(int(form.month_end.data),  int(form.year_end.data)))

      if end_date < start_date:
         error = u'Конечная дата не может быть меньше начальной!'
         print 'error'
         return render_template('select_date_range.html',  title='Audit trail', depart_id = depart_id , action = 'chs', error = error, form = form)
      result = CreateAuditTrailXls_CHS(start_date, end_date)
      if result == 'error':
        error = u'Нет данных за выбранный период'
      else:      
        return redirect(url_for('download_file', filename = result))
      return redirect(url_for('audit_trail')+'?depart_id='+ depart_id)
   return render_template('select_date_range.html',  title='Audit trail', depart_id = depart_id , action = 'chs', form = form)

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

@app.route('/uploads/<path:filename>')
def download_file(filename):
#    fold = app.static_folder+ "\uploads" 
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)
#    return send_from_directory(fold, filename)


