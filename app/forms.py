# -*- coding: utf-8 -*-
from flask_wtf import Form
from wtforms import StringField, TextAreaField, TextField, DateField, IntegerField, SelectField, SubmitField
from wtforms.validators import ValidationError, Required



class SelectDate(Form):
    day = SelectField(u'День', choices=[('1', '1'), ('2', '2'), ('3', '3'),('4', '4'), ('5', '5'), ('6', '6'),('7', '7'), ('8', '8'), ('9', '9'),('10', '10'), ('11', '11'), ('12', '12'), ('13', '13'), ('14', '14'), ('15', '15'),('16', '16'), ('17', '17'), ('18', '18'),('19', '19'), ('20', '20'), ('21', '21'),('22', '22'), ('23', '23'), ('24', '24'), ('25', '25'),('26', '26'), ('27', '27'), ('28', '28'),('29', '29'), ('30', '30'), ('31', '31')]) 
    month = SelectField(u'Месяц', choices=[('1', u'январь'), ('2', u'февраль'), ('3', u'март'),('4', u'апрель'), ('5', u'май'), ('6', u'июнь'),('7', u'июль'), ('8', u'август'), ('9', u'сентябрь'),('10', u'октябрь'), ('11', u'ноябрь'), ('12', u'декабрь')]) 
    submit = SubmitField(u'Ввод')

class SelectDateShort(Form):
    month = SelectField(u'Месяц', choices=[('1', u'январь'), ('2', u'февраль'), ('3', u'март'),('4', u'апрель'), ('5', u'май'), ('6', u'июнь'),('7', u'июль'), ('8', u'август'), ('9', u'сентябрь'),('10', u'октябрь'), ('11', u'ноябрь'), ('12', u'декабрь')]) 
    year = SelectField(u'Год', choices=[('2018', u'2018'), ('2019', u'2019'), ('2020', u'2020'), ('2021', u'2021'), ('2022', u'2022'), ('2023', u'2023')])
    submit = SubmitField(u'Ввод')

class SelectDateRangeShort(Form):

    month_start = SelectField(u'Месяц', choices=[('1', u'январь'), ('2', u'февраль'), ('3', u'март'),('4', u'апрель'), ('5', u'май'), ('6', u'июнь'),('7', u'июль'), ('8', u'август'), ('9', u'сентябрь'),('10', u'октябрь'), ('11', u'ноябрь'), ('12', u'декабрь')]) 
    year_start = SelectField(u'Год', choices=[('2018', u'2018'), ('2019', u'2019'), ('2020', u'2020'), ('2021', u'2021'), ('2022', u'2022'), ('2023', u'2023')])
    month_end = SelectField(u'Месяц', choices=[('1', u'январь'), ('2', u'февраль'), ('3', u'март'),('4', u'апрель'), ('5', u'май'), ('6', u'июнь'),('7', u'июль'), ('8', u'август'), ('9', u'сентябрь'),('10', u'октябрь'), ('11', u'ноябрь'), ('12', u'декабрь')]) 
    year_end = SelectField(u'Год', choices=[('2018', u'2018'), ('2019', u'2019'), ('2020', u'2020'), ('2021', u'2021'), ('2022', u'2022'), ('2023', u'2023')])
    submit = SubmitField(u'Ввод')

class AuditTrailForm(Form):
    objectname = TextAreaField(u'Наименование объекта', validators=[Required()])
    objectadres = TextAreaField(u'Адрес фактического осуществления деятельности', validators=[Required()])
    checkdate = DateField(u'Дата проведения проверки',  format='%d/%m/%Y', render_kw={'placeholder': '31/12/2018'},  validators=[Required()])
    of_violations = IntegerField(u'Выявлено нарушений по результатам проведения плановых и внеплановых проверок')
    of_violations_unscheduled = IntegerField(u'Выявлено нарушений по результатам внеплановых проверок, которые не устранены в установленные предписаниями сроки, всего')
    fixed_violations = IntegerField(u'Устранено нарушений в установленные предписаниями сроки по результатам внеплановых проверок, всего')
    other_documents = TextAreaField(u'Наименование, № других документов, составленных по результатам проверки, дата их составления')
    check_number = StringField(u'№ проверки по АС ЕРП')
    submit = SubmitField(u'Сохранить')
