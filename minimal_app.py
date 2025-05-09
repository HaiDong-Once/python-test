from flask import Flask

app = Flask(__name__)

@app.route('/')
def hello():
    return "Hello, World! 这是一个测试页面。"

if __name__ == '__main__':
    print("启动最小Flask应用")
    app.run(debug=True) 