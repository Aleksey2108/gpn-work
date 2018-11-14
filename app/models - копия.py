from app import db


class AuditTrail(db.Model):
    id = db.Column(db.Integer, primary_key = True)
    objectname = db.Column(db.String(256), default = '')
    objectadres = db.Column(db.String(256), default = '')
    depart_id = db.Column(db.String(64), default = '')
    checkdate = db.Column(db.Date(), default = '')
    of_violations = db.Column(db.SmallInteger, default = 0)
    of_violations_unscheduled = db.Column(db.SmallInteger, default = 0)
    fixed_violations = db.Column(db.SmallInteger, default = 0)
    name_employee = db.Column(db.String(128), default = '')
    other_documents = db.Column(db.String(256), default = '')
    check_number = db.Column(db.SmallInteger, default = '')

class AuditTrail_CHS(db.Model):
    id = db.Column(db.Integer, primary_key = True)
    objectname = db.Column(db.String(), default = '')
    objectadres = db.Column(db.String(), default = '')
    depart_id = db.Column(db.String(64), default = '')
    doc_stored = db.Column(db.String(64), default = '')
    checkdate = db.Column(db.Date(), default = '')
    type_inspection = db.Column(db.String(64), default = '')
    start_date = db.Column(db.Date(), index=True)
    end_date = db.Column(db.Date(), index=True)
    act_number = db.Column(db.String(128), default = '')
    act_date = db.Column(db.Date(), default = '')
    order_number = db.Column(db.String(128), default = '')
    order_date = db.Column(db.Date(), default = '')
    of_violations = db.Column(db.SmallInteger, default = 0)
    of_violations_unscheduled = db.Column(db.SmallInteger, default = 0)
    fixed_violations = db.Column(db.SmallInteger, default = 0)
    name_employee = db.Column(db.String(512), default = '')
    check_number = db.Column(db.SmallInteger, default = '')

    
    def __repr__(self):
       return '<AuditTrail_CHS {}>'.format(self.depart_id)
    
class AuditTrail_GO(db.Model):
    id = db.Column(db.Integer, primary_key = True)
    objectname = db.Column(db.String(), default = '')
    objectadres = db.Column(db.String(), default = '')
    depart_id = db.Column(db.String(64), default = '')
    doc_stored = db.Column(db.String(64), default = '')
    checkdate = db.Column(db.Date(), default = '')
    type_inspection = db.Column(db.String(64), default = '')
    start_date = db.Column(db.Date(), index=True)
    end_date = db.Column(db.Date(), index=True)
    act_number = db.Column(db.String(128), default = '')
    act_date = db.Column(db.Date(), default = '')
    order_number = db.Column(db.String(128), default = '')
    order_date = db.Column(db.Date(), default = '')
    of_violations = db.Column(db.SmallInteger, default = 0)
    of_violations_unscheduled = db.Column(db.SmallInteger, default = 0)
    fixed_violations = db.Column(db.SmallInteger, default = 0)
    name_employee = db.Column(db.String(512), default = '')
    other_documents = db.Column(db.String(), default = '')
    check_number = db.Column(db.SmallInteger, default = '')

    def __repr__(self):
       return '<AuditTrail_GO {}>'.format(self.depart_id)

class AuditTrail_PB(db.Model):
    id = db.Column(db.Integer, primary_key = True)
    objectname = db.Column(db.String(), default = '')
    objectadres = db.Column(db.String(), default = '')
    depart_id = db.Column(db.String(64), default = '')
    doc_stored = db.Column(db.String(64), default = '')
    checkdate = db.Column(db.Date(), default = '')
    type_inspection = db.Column(db.String(64), default = '')
    start_date = db.Column(db.Date(), index=True)
    end_date = db.Column(db.Date(), index=True)
    act_number = db.Column(db.String(128), default = '')
    act_date = db.Column(db.Date(), default = '')
    order_number = db.Column(db.String(128), default = '')
    order_date = db.Column(db.Date(), default = '')
    of_violations = db.Column(db.SmallInteger, default = 0)
    of_violations_unscheduled = db.Column(db.SmallInteger, default = 0)
    fixed_violations = db.Column(db.SmallInteger, default = 0)
    name_employee = db.Column(db.String(512), default = '')
    check_number = db.Column(db.SmallInteger, default = '')

    def __repr__(self):
       return '<AuditTrail_PB {}>'.format(self.depart_id)

