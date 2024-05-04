# scan-folder-all-file
使用tkinter开发的一款可扫描并删除本地文件敏感词的Windows软件


1、默认状态仅可点击选择文件夹按钮，导出、停止按钮此时不可用

2、逐个扫描并显示已扫描文件和异常文件（包含敏感词）

	2.1、此时会逐个追加到列表，并在列表可忽略或删除操作
 
	2.2、扫描过程中会给已扫描的文件加索引号，删除或忽略时自动重排索引
 
3、扫描过程可中断

4、扫描完成后可导出已扫描文件和异常文件

5、导出的异常文件包含所有敏感词

6、敏感词可配置、可扩展，目前仅能从本地加载，未来可扩展为从云端加载，动态扩展敏感词库

7、删除文件时是真实删除电脑系统文件！，请谨慎操作，不提供恢复方法

	7.1、删除后会先删除电脑文件，再删除列表文件
 
8、点击忽略时将该文件从已扫描列表或者异常文件列表剔除，后续的导出就不会导出这个文件了

9、软件可自由配置标题和logo，后续也可通过云端获取配置的方式加载

10、扫描过程中动态展示已扫描的文件夹数量和已扫描的文件数

11、重新选择一个文件扫描时，上一次的扫描结果会自动清空，且导出、选择文件夹按钮暂时不可用，此时停止扫描按钮变为可用状态

12、界面可以放大缩小，不影响功能和显示

13、可自由打包为exe





更新：
1、支持扫描文件内容

2、支持删除文件内容

3、支持的文件类型：txt、pdf、doc、docx

4、导出的异常文件报告包含该文件包含的所有敏感词



配置：
1、可配置软件标题

	打开consts.py：
 
		更改SOFT_NAME的值
  
2、可配置软件logo

	将本地的logo.ico替换重新运行即可
 
3、更新敏感词库

	打开consts.py：
 
		对ERROR_WORDS操作即可，可新增、删除、修改
  
		如新增一个敏感词为“测试”，在原有基础上：ERROR_WORDS = [
						"木马", "命令", "shell"
				  ]
      
		修改为：ERROR_WORDS = [
					"木马", "命令", "shell", "测试"
				]
    
		注意不要改变产量名或者数据结构！
  


未来展望：

1、敏感词库可以云端管理，软件启动后自动加载云端敏感词库

2、扫描的所有文件名都可以上传云端服务器分析

3、扫描完成可发送短信或邮件通知

4、可扫描文件内容，通过本地读取文件内容方式判断是否为病毒文件或者异常文件




项目运行：
	1、python版本建议3.10
	
	2、建议使用conda虚拟环境，配置好虚拟环境后，使用命令安装项目所需包：
		pip install -r requirements.txt
	
	3、运行：使用命令：
		python main.py 
	或者在pycharm打开的话，点击运行按钮即可一键启动
	
	4、打包为exe可执行文件：
		pyinstaller --noconsole -y --add-data "*;." --add-data "./logo.ico;." -i logo.ico -F -n 扫描文件 main.py
		--noconsole指令表示不显示控制台
		-y指令表示默认选择确定，因为可能重复打包，需要询问是否删除之前打包好的exe
		--add-data 将当前目录下所有内容一起打包
		-i logo.ico 指定logo
		-F 产生一个文件用以部署，没有其余文件
		-n 指定软件名
		
		打包的exe在dist下，二次打包前，最好先删除dist、build两个文件夹和扫描文件.spec文件

代码参考：
    1、解析doc、docx：https://blog.csdn.net/weixin_40449300/article/details/79143971


来一个测试，我在桌面有一个test文件：
	![image](https://github.com/2424004764/scan-folder-all-file/assets/24261680/e8c0979e-5835-41e7-96c7-cd003798e78f)

扫描一下：
![image](https://github.com/2424004764/scan-folder-all-file/assets/24261680/30d729d2-bf12-456a-b68d-ce372e2c4e6e)

显示扫描到来敏感词，我们点击删除：

敏感词被删除了。![image](https://github.com/2424004764/scan-folder-all-file/assets/24261680/a3bc3164-a38a-417f-a32a-c13250494a49)

    

可以在consts.py中配置软件的名称和logo：

![image](https://github.com/2424004764/scan-folder-all-file/assets/24261680/90b53d18-f891-4543-aec2-72c261354c86)

支持检测的敏感词：

![image](https://github.com/2424004764/scan-folder-all-file/assets/24261680/545abe00-6ac1-4e3c-99f5-50fc2a9ac3b8)

如果要支持更多的文件后缀，则需要扩展软件功能：
![image](https://github.com/2424004764/scan-folder-all-file/assets/24261680/1874b8a4-e86a-482e-b4fd-824133c7f7ae)
1、先扩展扫描时的后缀：
![image](https://github.com/2424004764/scan-folder-all-file/assets/24261680/790e3466-38bd-4d9c-b44f-7363abe13591)
2、删除时也要扩展后缀：
![image](https://github.com/2424004764/scan-folder-all-file/assets/24261680/233c9897-ed4a-4b51-96b7-9fa58dcf15e0)


本软件可自由使用。
