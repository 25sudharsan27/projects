from flask import Flask, render_template,request
app=Flask(__name__)

@app.route("/")

def index():
    return render_template("form.html")

@app.route("/Register",methods=["POST","GET"])

def Register():
    if request.method == "POST":
        name=request.form.get("name")
        section=request.form.get("section")
        gender=request.form.get("gender")
        age=request.form.get("age")

        return render_template("result.html",name=name,section=section,gender=gender,age=age)

if "__main__" == __name__:
    app.run(debug=True)
