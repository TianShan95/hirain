# coding:utf-8
from flask import Flask,render_template,request,redirect,url_for
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import webbrowser
import flask
import sys
from datetime import timedelta
import get_csv_01


def runflask():

    if getattr(sys, 'frozen', False):
        ROOT_PATH = os.path.dirname(sys.executable)+r"\static"
        ROOT_PATH = ROOT_PATH.replace('\\', '/')
        # print('a')
    elif __file__:
        ROOT_PATH = os.path.dirname(__file__)+"/static"
        # print('b')

    print("10:42")
    app = Flask(__name__)
    # 允许全局跨域配置
    CORS(app, supports_credentials=True)
    app.config['TEMPLATES_AUTO_RELOAD'] = True
    # 设置静态文件缓存过期时间
    app.send_file_max_age_default = timedelta(seconds=1)

    @app.route('/index')
    def index():
        return render_template('index.html')

    @app.route('/upload', methods=['POST', 'GET'])
    def upload():
        if request.method == 'POST':
            f = request.files['file']
            basepath = os.path.dirname(__file__)  # 当前文件所在路径
            upload_path = os.path.join(basepath, r'static\data', secure_filename(f.filename))  #注意：没有的文件夹一定要先创建，不然会提示没有该路径
            f.save(upload_path)
            print(upload_path)

            # file_path = os.path.join(basepath, r'static\data')
            # for root, dirs, files in os.walk(file_path):
            #     print(root)  # 当前目录路径
            #     print(dirs)  # 当前路径下所有子目录
            #     print(files)  # 当前路径下所有非目录子文件
            print('文件路径：')
            log_path = os.path.join(basepath, r'static\data\top.log')
            csv_path = os.path.join(basepath, r'static\data\top.csv')
            print(log_path)
            print(csv_path)
            p = get_csv_01.GetCoord()
            p.read_file(log_path, csv_path)
            print('文件读取完成')

            return redirect(url_for('index'))
        return render_template('upload.html')

    @app.route('/favicon.ico')
    def favicon():
        return flask.send_from_directory(ROOT_PATH, 'logo.png', mimetype='image/vnd.microsoft.icon')

    # url_for,修改静态文件（js,css,image)时，网页同步修改
    @app.context_processor
    def override_url_for():
        return dict(url_for=dated_url_for)

    def dated_url_for(endpoint, **values):
        filename = None
        if endpoint == 'static':
            filename = values.get('filename', None)
        if filename:
            file_path = os.path.join(app.root_path, endpoint, filename)
            values['v'] = int(os.stat(file_path).st_mtime)
        return url_for(endpoint, **values)

    url = 'http://127.0.0.1:63343/upload'
    webbrowser.open(url, new=1)  # open in new tab
    app.run(debug=False, port=63343)


if __name__ == '__main__':
    runflask()