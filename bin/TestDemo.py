"""
    Author：
        Zhang Zhi
    功能：
        读取robotframework自动化测试报告output.xml文件，将部分关键字段测试数据写入到excel文件中
    使用方法：
        1、修改存放output.xml文件的路径
        2、修改生成excel文件的目录及文件名称
        3、本地运行TestDemo.py文件
"""
import time,sys,os
#获取util目录的相对路径
utilpath = os.path.abspath("../util")
#添加环境变量--注：必须在引用自己的模块前先加入到环境变量，否则会报“模块找不到的错误”
sys.path.append(utilpath)
from parsexml import parse_robot_results
from parseexcel import write_excel
# report_path = os.path.abspath("../report/Lift&shift Execution Report【TUI】/output.xml")
#Robotframe自动化测试框架生成的output.xml测试报告结果文件路径
path = r'C:\Users\ZhiZhang\Desktop\Lift&shift Execution Report【TUI】\output.xml'
#生成excel的存放目录及文件名称，自行修改即可。
filename = r'C:\Workspace\ParseXml\excel\test_{0}.xlsx'.format(time.strftime("%m%d%H%M%S"),time.localtime())
results,statistics = parse_robot_results(path)
write_excel(results,statistics,filename)