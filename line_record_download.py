import os
import re
import pymysql
import requests
import datetime
import logging
import time
from dingding import Dingding_Warning

# 配置日志
logging.basicConfig(
    filename=r'E:\联通白名单录音下载\download_errors.log',  # 日志文件路径
    level=logging.ERROR,  # 记录错误及以上级别的日志
    format='%(asctime)s - %(levelname)s - %(message)s'  # 日志格式
)


# 下载联通的录音
class line_download:
    def __init__(self):
        self.str_time = ''
        self.download_path = r'\\10.10.100.203\recordings\网络电话录音\联通云录音'

        self.host = '10.10.100.81'
        self.port = 3306
        self.user = 'root'
        self.passwd = 'jinwan_88888'
        self.name = 'line_phone_record'

    # 查询需要下载录音的联通通话记录
    def select_line_call(self):
        mysqldb = pymysql.connect(host=self.host, port=self.port, user=self.user, password=self.passwd,
                                  database=self.name)  # 打开数据库链接
        curs = mysqldb.cursor()
        record_index_list = []
        # 根据时间区间计算时间戳
        date1 = datetime.datetime.strptime(self.str_time + ' 00:00:00', '%Y-%m-%d %H:%M:%S')  # 开始时间
        date2 = date1 + datetime.timedelta(days=1)  # 结束时间
        print('开始下载{} 至 {}的联通录音'.format(date1, date2))

        sql = """
            SELECT callType,displayNumber, calledNo, startTime, endTime, 
                   ringDuration, isSuccess, recordUrl 
            FROM line_record_download_down_call_list 
            WHERE callStartTime >= %s and callStartTime < %s 
              and LENGTH(recordUrl) > 0 and isSuccess = 1
        """
        curs.execute(sql, (date1, date2))
        list_a = curs.fetchall()
        for call in list_a:
            call = list(call)
            if call[0] == '呼出':
                call[0] = 'OUT'
            else:
                call[0] = 'IN'
            date_str = re.sub(r'\D', '', call[3])
            root = os.path.join(self.download_path, date_str[0:6], date_str[6:8])  # 日期目录
            file = '{}_{}_{}_{}.wav'.format(call[0], call[1], call[2], date_str)
            call += [root, file]
            record_index_list.append(call)
            # print(call)
        curs.close()  # 关闭游标对象
        mysqldb.close()  # 关闭链接
        return record_index_list

    # 下载录音
    def record_download(self, url, root, name):
        max_retries = 3
        connect_timeout = 10  # 连接超时时间
        read_timeout = 30  # 读取超时时间
        retry_delay = 2  # 重试前等待秒数

        filepath = os.path.join(root, name)
        if not os.path.exists(root):
            os.makedirs(root)

        retries = 0
        while retries < max_retries:
            try:
                if not os.path.exists(filepath):
                    head = {"Accept": "application/json, text/plain, */*"}
                    # 分别设置连接和读取超时
                    r = requests.get(url, headers=head, timeout=(connect_timeout, read_timeout))

                    if r.status_code == 200:
                        with open(filepath, 'wb') as f:
                            f.write(r.content)
                        print(f'文件下载成功: {filepath}')
                        return True
                    else:
                        print(f'下载失败，错误代码 {r.status_code}，URL: {url}')
                        retries += 1
                        if retries < max_retries:
                            time.sleep(retry_delay)
                        continue
                else:
                    print(f'文件已存在: {filepath}')
                    return True

            except requests.exceptions.Timeout:
                retries += 1
                print(f'请求超时，正在重试 ({retries}/{max_retries})...')
                if retries < max_retries:
                    time.sleep(retry_delay)

            except requests.exceptions.RequestException as e:
                retries += 1
                print(f'请求异常: {e}，正在重试 ({retries}/{max_retries})...')
                if retries < max_retries:
                    time.sleep(retry_delay)

            except Exception as e:
                retries += 1
                print(f'未知错误: {e}，正在重试 ({retries}/{max_retries})...')
                if retries < max_retries:
                    time.sleep(retry_delay)

        # 达到最大重试次数仍未成功
        error_message = f'重试次数已用尽，下载失败: {url}'
        logging.error(error_message)
        print(f'最终下载失败: {url}')
        return False

    # 插入记录到数据库
    def insert_record(self, data_list):
        if not data_list:
            print("没有新数据需要插入")
            return

        mysqldb = pymysql.connect(host=self.host, port=self.port, user=self.user, password=self.passwd,
                                  database=self.name)  # 打开数据库链接
        curs = mysqldb.cursor()

        # 提取文件名列表用于去重检查
        names = [data[8] for data in data_list]
        placeholders = ','.join(['%s'] * len(names))

        # 检查已存在的记录
        sql_check = f"SELECT name FROM line_record_download_record_list WHERE name IN ({placeholders})"
        curs.execute(sql_check, names)
        existing_names = set(row[0] for row in curs.fetchall())

        # 过滤掉已存在的记录
        new_data = [
            tuple(data[:7] + [os.path.join(data[8], data[9]), data[8]])  # 构造完整路径作为 file_path 字段
            for data in data_list
            if data[8] not in existing_names
        ]

        if not new_data:
            print("所有记录均已存在，无需插入")
            curs.close()
            mysqldb.close()
            return

        # 批量插入
        sql_insert = """
            INSERT IGNORE INTO line_record_download_record_list 
            (context, caller, callee, start_time, end_time, duration, call_type, file_path, name) 
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """
        try:
            curs.executemany(sql_insert, new_data)
            mysqldb.commit()
            print(f'成功批量插入 {len(new_data)} 条数据')
        except Exception as e:
            mysqldb.rollback()
            print(f'批量插入失败: {e}')
            logging.error(f'批量插入失败: {e}')
        finally:
            curs.close()
            mysqldb.close()

    # 主函数
    def run(self):
        log_list = []
        call_list = self.select_line_call()
        total = len(call_list)

        for idx, call in enumerate(call_list, 1):
            print(f"[{idx}/{total}] 正在处理...")
            if self.record_download(call[7], call[8], call[9]):
                log_list.append(call)

        self.insert_record(log_list)
        print('全部下载及入库完成')


def main():
    # 定义开始时间和结束时间
    start_date = datetime.date(2025, 9, 25)  # 开始时间
    end_date = datetime.date(2025, 10, 6)  # 结束时间
    # start_date = datetime.datetime.today() + datetime.timedelta(days=-7)  # 开始时间
    # end_date = datetime.datetime.today() + datetime.timedelta(days=-1)  # 结束时间
    while start_date <= end_date:
        date = (start_date.strftime("%Y-%m-%d"))  # 当天的日期
        start_date += datetime.timedelta(days=1)
        LD = line_download()
        LD.str_time = date
        LD.run()


if __name__ == '__main__':
    try:
        main()
    except Exception as k:
        print(k)
        Dingding_Warning('严重', '程序运行错误', k)
