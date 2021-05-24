from PyInstaller.__main__ import run

if __name__ == '__main__':
    opts = ['main.py',  # 主程序文件
            '-p get_csv_01.py',
            '-p url.py',
            '-n hirain_data_visualization_0.0.4',  # 可执行文件名称
            '-F',  # 打包单文件
            # '-w', #是否以控制台黑窗口运行
            r'--icon=D:/shan.tian/Downloads/20210329/20210427/templates/data.ico',  # 可执行程序图标
            '-y',
            '--clean',
            '--workpath=build',
            '--add-data=templates;templates',  # 打包包含的html页面
            '--add-data=static;static',  # 打包包含的静态资源
            '--distpath=build',
            '--specpath=./'
            ]

    run(opts)

