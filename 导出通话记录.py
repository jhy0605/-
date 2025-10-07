import pymysql
import datetime
import openpyxl
import time


# 数据库操作
class open_sql:
    def __init__(self):
        self.host = '10.10.100.81'
        self.port = 3306
        self.user = 'root'
        self.passwd = 'jinwan_88888'
        self.name = 'line_phone_record'
        self.datetime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.mysqldb = pymysql.connect(host=self.host, port=self.port, user=self.user, password=self.passwd,
                                       database=self.name)  # 打开数据库链接
        self.curs = self.mysqldb.cursor()
        self.file_path = r'白名单{}.xlsx'.format(
            datetime.datetime.now().strftime("%Y%m%d%H%M%S"))

    def call_list(self):
        # 获取当天的数据
        sql_1 = "SELECT callType as '通话类型', callerNo as '主叫', calledNo as '被叫', startTime as '开始时间'," \
                " endTime as '结束时间' FROM line_record_download_down_call_list WHERE DATE(startTime) = CURDATE();"
        # 获取最近三天的数据
        sql_2 = "SELECT callType as '通话类型', callerNo as '主叫', calledNo as '被叫', startTime as '开始时间', " \
                "endTime as '结束时间' FROM line_record_download_down_call_list WHERE DATE(startTime) " \
                "BETWEEN DATE_SUB(CURDATE(), INTERVAL 2 DAY) AND CURDATE();"
        h = str(input(r'请选择需要导出的时间段：（1：今天，2：最近三天）'))
        if h == '1':
            sql_x = sql_1
        else:
            sql_x = sql_2
        # print(sql_x)
        print('正在导出......')
        self.curs.execute(sql_x)  # 执行sql语句
        fetchall = self.curs.fetchall()
        excel_taber = openpyxl.Workbook()  # 创建excel对象
        sheet1 = excel_taber.active
        sheet1.append(['通话类型', '主叫', '被叫', '开始时间', '结束时间'])
        for call in fetchall:
            # print(call)
            sheet1.append(list(call))
        excel_taber.save(self.file_path)
        print('导出完成')
        time.sleep(3)


if __name__ == '__main__':
    os = open_sql()
    os.call_list()
