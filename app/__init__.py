import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from config import Config

app = Flask(__name__)
app.config.from_object(Config)
db = SQLAlchemy(app)

UPLOAD_FOLDER = "app/uploads"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


from app import routes,  function
