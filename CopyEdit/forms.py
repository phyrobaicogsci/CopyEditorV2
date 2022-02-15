from flask.helpers import send_file
from flask_wtf import Form
from wtforms import StringField, IntegerField, TextAreaField, SubmitField, BooleanField, RadioField, SelectField
from wtforms import validators, ValidationError
from flask_wtf.file import FileField, FileRequired, FileAllowed

class ContactForm(Form):
   #send_file = RadioField('Email me corrections when done?', choices = [('Yes','Yes'),('No','No')])
   raw_text = TextAreaField("Enter text:")
   doc = FileField("Or upload a .docx file", validators=[FileAllowed(['docx','doc'],'.docx or .doc files only ;)')])
   send_file = BooleanField('Send me the corrections', default = False)
   email = StringField("Email")
   #model = SelectField('Models', choices = [('Gramformer', 'Gramformer'),('T5', 'T5')])
   submit = SubmitField("Send")