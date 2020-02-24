from robot.api import ExecutionResult
from strftime import time_transfer
def parse_robot_results(xml_path):
    result = []
    #获取测试集合
    suite = ExecutionResult(xml_path).suite
    allTests = suite.statistics.critical
    #开始时间
    StartTime = suite.starttime
    #结束时间
    EndTime = suite.endtime
    #持续时间
    ElapsedTime = suite.elapsedtime
    #格式化持续时间
    Elapsedtime_transfer = time_transfer(ElapsedTime)
    #测试状态
    Status = suite.status
    #总case条数
    Total = allTests.total
    #成功case条数
    PassNum = allTests.passed
    #失败case条数
    FailNum = allTests.failed
    #统计信息
    Statistics = [['Status',Status],['StartTime',StartTime],['EndTime',EndTime],['ElapsedTime',Elapsedtime_transfer],['Total',Total],['PassNum',PassNum],['FailNum',FailNum]]
    # 遍历所有测试集合
    for tests in suite.suites:
        # 遍历每个集合下面的测试case
        for test in tests.tests:
            #case名称
            case_name = test.name
            #标签
            tag = str(test.tags)
            #状态
            status = test.status
            #返回信息
            message = test.message
            #持续时间
            elapsedtime = test.elapsedtime
            #格式化持续时间
            elapsedtime_transfer = time_transfer(elapsedtime)
            result.append([case_name,tag,status,message,elapsedtime_transfer])

    return result,Statistics

if __name__ == '__main__':
    path = r'C:\Users\ZhiZhang\Desktop\reports\output.xml'
    r = parse_robot_results(path)