from flask import Flask
from flask_bootstrap import Bootstrap
from flask_datepicker import datepicker


app = Flask(__name__)
#datepicker(app)
datepicker(app=app, local=['static/assets/js/jquery-ui.js', 'static/assets/css/jquery-ui.css'])
bootstrap = Bootstrap(app)
app.config['SECRET_KEY'] = 'you-will-never-guess'

from app import routes

