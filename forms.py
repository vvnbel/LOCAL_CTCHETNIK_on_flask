from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, BooleanField, SubmitField, DateField
from wtforms.validators import DataRequired


class DateForm(FlaskForm):
    entrydate_from = DateField('entrydate', format='%Y-%m-%d')
    entrydate_to = DateField('entrydate', format='%Y-%m-%d')



