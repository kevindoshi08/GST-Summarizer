from flask import Flask, render_template, request, send_file
from summarizer import summarize

app = Flask(__name__)  
 
@app.route('/')  
def upload():  
    return render_template("upload.html")  
 
@app.route('/success', methods = ['POST'])  
def success():  
    if request.method == 'POST':  
        f = request.files['Excel_file']
        summarize(f)
        path = "./Summary.xlsx"
        return send_file(path, as_attachment=True)
  
if __name__ == '__main__':  
    app.run(threaded=True, port=5000)