from flask import Flask, render_template, request, send_file
import pdf2ppt


app = Flask(__name__)


@app.route("/")
@app.route("/home")

def home():
    return render_template("index.html")




@app.route("/submit", methods=["POST"])
def submit():
    checkbox_count = 0
    checkbox_names = []
    
    if "pdf1" in request.files and "pdf2" in request.files:
        pdf1 = request.files["pdf1"]
        pdf2 = request.files["pdf2"]

        print(request.form.getlist("checkbox"))
        
        for value in request.form.getlist("checkbox"):
            checkbox_count += len(request.form.getlist("checkbox"))
            checkbox_names.append(value)

        if checkbox_count >= 1:
            modified_presentation_path = pdf2ppt.run(pdf1, pdf2, checkbox_names)
            return render_template("index.html", output=modified_presentation_path)
        

    error_message = "Please select PDF files and check at least 5 checkboxes."
    return render_template("index.html", error_message=error_message)




@app.route("/download")
def download_presentation():
    modified_presentation_path = "modified_presentation.pptx"
    return send_file(modified_presentation_path, as_attachment=True)



if __name__ == '__main__':
    app.run(debug=True, port=5000)

