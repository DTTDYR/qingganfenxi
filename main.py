from flask import Flask
from hander import parse_xlsx

app = Flask(__name__)

@app.route("/")
def hello_world():
    return "<p>Hello, World!</p>"


@app.route("/parse_xlsx")
def get_xlsx():
    return parse_xlsx()


if __name__ == "__main__":
    app.run()
