import hmac
import hashlib
import base64
import urllib.parse
import requests
import time
import datetime
import json


# 钉钉发消息
def Dingding_Warning(grade, information, details):
    project = '联通云录音下载'
    w_datetime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 警告时间
    timestamp = str(round(time.time() * 1000))  # 生成时间戳
    secret = 'SECe5e20972577ea603ec828831c53e0bfc2ac064c07c259216f72cb0c187915c30'  # 设置机器人模式为加签，此为秘钥

    # 对秘钥进行加密处理
    secret_enc = secret.encode('utf-8')
    string_to_sign = '{}\n{}'.format(timestamp, secret)
    string_to_sign_enc = string_to_sign.encode('utf-8')
    hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
    sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))

    header = {'Content-Type': 'application/json;charset=utf-8'}
    url = 'https://oapi.dingtalk.com/robot/send?access_token=57f0183f439d5946ae2aa8173d9f8c8b2c9fc3ee28d5477e4a4433b22625c391&timestamp={}&sign={}'.format(
        timestamp, sign)
    json_text = {
        "msgtype": "text",
        "text": {"content": '告警项目：{}\n'
                            '告警时间：{}\n'
                            '告警等级：{}\n'  # 缺陷等级一般划分为四个等级：致命、严重、一般、低
                            '告警信息：{}\n'
                            '问题详情：{}'.format(project, w_datetime, grade, information, details)},
        "at": {"atMobiles": [""], "isAtAll": False}
    }
    requests.post(url, json.dumps(json_text), headers=header)
