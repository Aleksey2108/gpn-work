import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from config import Config

app = Flask(__name__)
app.config.from_object(Config)
db = SQLAlchemy(app)

#UPLOAD_FOLDER = "app/static/uploads"
#UPLOAD_FOLDER = os.path.abspath(os.path.dirname(__file__))+"\static\uploads"
UPLOAD_FOLDER = os.path.abspath(os.path.dirname(__file__))+"/static/uploads"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
MASTER_FOLDER = os.path.abspath(os.path.dirname(__file__))+"/static/master"
app.config['MASTER_FOLDER'] = MASTER_FOLDER

from app import routes,  function
