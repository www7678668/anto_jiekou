from flask import Flask
from flask import render_template,request

# 实例化，可视为固定格式
app = Flask(__name__)

# route()方法用于设定路由；类似spring路由配置
@app.route('/helloworld')
def hello_world():
    return 'Hello, World!'

# 配置路由，当请求get.html时交由get_html()处理得的
@app.route('/get.html')
def get_html():
    # 使用render_template()方法重定向到templates文件夹下查找get.html文件
    return render_template('get.html')

# 配置路由，当请求post.html时交由post_html()处理
@app.route('/post.html')
def post_html():
    # 使用render_template()方法重定向到templates文件夹下查找post.html文件
    return render_template('post.html')

# 配置路由，当请求deal_request时交由deal_request()处理
# 默认处理get请求，我们通过methods参数指明也处理post请求
# 当然还可以直接指定methods = ['POST']只处理post请求, 这样下面就不需要if了
@app.route("/deal_request", methods = ['GET', 'POST'])
def deal_request():
    if request.method == "GET":
        # get通过request.args.get("param_name","")形式获取参数值
        get_q = request.args.get("q","")
        return render_template("result.html", result=get_q)
    elif request.method == "POST":
        # post通过request.form["param_name"]形式获取参数值
        post_q = request.form["q"]
        return render_template("result.html", result=post_q)


if __name__ == '__main__':
    # app.run(host, port, debug, options)
    # 默认值：host="127.0.0.1", port=5000, debug=False
    app.run(host="127.0.0.1", port=5000)

