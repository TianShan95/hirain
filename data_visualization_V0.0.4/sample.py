# coding:utf-8
from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
import os
import sys


# 获取资源路径
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def runflask():
    app = Flask(__name__, static_url_path="", static_folder=resource_path('static'),
                template_folder=resource_path("templates"))

    @app.route('/')
    def index():
        return render_template('upload.html')
    @app.route('/upload', methods=['POST', 'GET'])
    def upload():
        if request.method == 'POST':
            print('进来了')
            f = request.files['file']
            basepath = os.path.dirname(__file__)  # 当前文件所在路径
            upload_path = os.path.join(basepath, 'static/uploads',secure_filename(f.filename))  #注意：没有的文件夹一定要先创建，不然会提示没有该路径
            f.save(upload_path)
            return redirect(url_for('upload'))
        return render_template('upload.html')
    app.run(host='127.0.0.1', port=63339, debug=False)



if __name__ == '__main__':
    runflask()