from flask import Flask, stream_with_context, redirect,abort,render_template, url_for, request
from werkzeug.utils import secure_filename
from flask_mail import Mail, Message
from flask import Flask, render_template, request, flash
from .forms import ContactForm
import sqlite3 as sql
from .copyeditor import start
from werkzeug.utils import secure_filename
from docx.api import Document
# from waitress import serve
import os
app = Flask(__name__)
app.config['MAIL_SERVER']='smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = 'dummytestce@gmail.com'
app.config['MAIL_PASSWORD'] = 'dummy@123'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)
app.secret_key = os.urandom(12)
@app.route("/")
def hello_world():
   return render_template('home.html')

@app.route('/about')
def about():
   return 'Hey this is CopyEditor! I will help you solve and suggest copy edit <s>errors</s>.'


@app.route('/rule/<int:postID>')
def show_blog(postID):
   return 'Rule Number %d' % postID

@app.route('/result',methods = ['POST', 'GET'])
def result():
      global data,filename
      with sql.connect("database.db") as con:
         #try:
            if data['file'] != '':
               with open(data['file'],'rb') as f:
                  doc_text = f.read()
               with open("input.docx", "w+b") as i:
                  i.write(doc_text)
               filename = data['file']
               corrected_text, suggestions, grammar = start()

            else:
               raw_text = data['raw_text']
               document = Document()
               document.add_paragraph(raw_text)
               document.save('input.docx')
               filename = "raw_text"
               corrected_text, suggestions, grammar = start()
            send_file = "Yes" #request.form['send_file']
            email = data['email']
            model = data['model']
            file = data['file']
            
      #result = request.form
      if data['send_file']!="No":
         msg = Message('Hola!', sender = 'rohansworkid@gmail.com', recipients = [email])
         msg.body = f"""
         Hey I'm copyeditor! Please find the corrections file for your {filename}.
         
         
         Developed and maintained by Swapnil Lonarey, Rahul Bharti and Rohan Sahni.
          """
         with app.open_resource("./corrected.docx") as fp:  
            msg.attach(f"{filename}_corrected.docx","text/docx",fp.read())  
            mail.send(msg)

      dict = {"tags" : corrected_text,"suggestions":suggestions,"grammar":grammar}   
      print(corrected_text)  
      return render_template('result.html', result = dict)

@app.route('/edit',methods = ['POST', 'GET'])
def edit():
   global filename, data
   print("##############################################")
   form = ContactForm()
   if request.method == 'POST':
      print("##############################################")
      #return render_template("result.html",msg = msg)
      print(request.form,form.validate(), request.files)
      file = request.files['doc']
      if form.validate() == False:
         print("The form didn't validate because of ", form.errors)
         flash('All fields are required.')
         return render_template('edit.html', form = form)
      else:
         print("################## Filename: ############# ", form.doc.data)
         if file.filename != "":
            filename = secure_filename(file.filename)
            file.save(filename)
         else:
            filename = ""
         if 'send_file' not in request.form:
            send_file = "No"
            email = ""
         else:
            send_file = request.form['send_file']
            email = request.form['email']
         data = {
            "raw_text": request.form['raw_text'],
            "email": email,
            "send_file": send_file,
            "model": "Gramformer",
            "file": filename
         }
         #form.doc.data.save('uploads/' + filename)
         return redirect(url_for('loading'))
   elif request.method == 'GET':
      return render_template('edit.html', form = form)

    
   
@app.route('/loading')
def loading():
   return render_template('loading.html')
      


# if __name__ == '__main__':
#    app.run(debug=True)

# serve(app, port=8080, host="127.0.0.1")
