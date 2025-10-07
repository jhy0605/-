import os
import time
import subprocess
import threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

# 目标 IP 地址
targets = [
    ("Server_1", "10.10.100.206"),
    ("Server_2", "10.10.100.210"),
    ("Public_DNS", "223.5.5.5"),
    ("Baidu", "www.baidu.com")
]

log_file = "log.txt"
duration = 30  # 3 分钟


def ping(host):
    """仅在 Windows 系统中执行 ping 命令，并解析延迟"""
    # 如果是 Windows 系统，使用 ping 命令并设置参数为 -n 1，只执行一次 ping
    command = ["ping", "-n", "1", host]  # Windows 使用 -n 1 来指定只 ping 一次

    try:
        # 使用 subprocess 执行 ping 命令
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        # 读取并输出结果
        output, error = process.communicate()
        # print(output)
        txt = ''
        for line in output.splitlines():
            print(line)
            if "数据包:" in line:
                txt = line.split('(')[1].replace(')', '')
            if "的回复: " in line:
                txt += line.split("的回复: ")[1]
            txt = host + '：' + txt
        return txt

    except Exception as e:
        return f"Error: {e}"


def log_latency():
    """每秒检测所有目标的 ping 延迟并记录到日志"""
    n = 0

    while n < duration:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entries = [f"[{timestamp}]"]

        # 提交 ping 任务并保存 future 对象
        futures = []
        pool = ThreadPoolExecutor(max_workers=len(targets))  # 线程池只需要创建一次
        for ip in targets:
            future = pool.submit(ping, ip[1])  # 提交 ping 任务
            futures.append(future)

        # 等待所有任务完成并获取结果
        t = datetime.now()
        log_txt = t.strftime('%Y-%d-%m %H:%M:%S') + '：'
        for future in futures:
            result = future.result()  # 获取每个 ping 的结果
            log_txt += result + ' | '  # 使用 " | " 分隔结果

        with open(log_file, "a", encoding="utf-8") as log:
            log.write(log_txt + "\n")

        time.sleep(1)
        n += 1
        print(n)


if __name__ == "__main__":
    log_latency()
