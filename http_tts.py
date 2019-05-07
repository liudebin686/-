# coding=utf-8
import xlrd, os, time, base64, json, httplib, requests
from xlwt import *
from xlrd import open_workbook  # 导入openworkbook模块
from xlutils.copy import copy  # 导入copy模块

def excel_colour_pass():
    # 创建一个样式----------------------------
    style = XFStyle()
    pattern = Pattern()
    pattern.pattern = Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = Style.colour_map['bright_green']  # 设置PASS单元格背景色为浅绿色
    style.pattern = pattern
    return style
    # -----------------------------------------
def excel_colour_fail():
    # 创建一个样式----------------------------
    style = XFStyle()
    pattern = Pattern()
    pattern.pattern = Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = Style.colour_map['red']  # 设置FAIL单元格背景色为红色
    style.pattern = pattern
    return style
    # -----------------------------------------
def TTS_TEST():
    url = "http://10.0.0.61:8001/v1/{e9f60196209c468e898dd281ac09499b-id}/tts/short-audio"
    data = open_workbook('tts_http_case.xls', formatting_info=True)
    wb = copy(data)
    project = data.sheet_names()[0]
    she = data.sheet_by_name(project)
    config_list = she.row_values(2)
    weight_skip = config_list.index("skip")
    weight_auth = config_list.index("X-Auth-Token")
    weight_text = config_list.index("text")
    weight_config = config_list.index("config")
    weight_except_errorno = config_list.index("Expect_ErrorNo")
    weight_Expect_ResMessage = config_list.index("Expect_ResMessage")
    weight_status = config_list.index("status")
    weight_ErrorNo = config_list.index("ErrorNo")
    weight_ResMessage = config_list.index("ResMessage")
    weight_Expect_status =config_list.index("Expect_status")
    weight_foramt = config_list.index("audio_format")
    weight_speed = config_list.index("speed")
    ws = wb.get_sheet(project)
    for i in range(3, she.nrows):
        case_value = she.row_values(i)
        if case_value[weight_skip] == 'skip':
            continue
        else:
            print ("\033[1;35m Case %s is comming... about %s \033[0m" % (str(int(case_value[0])), case_value[1]))
            headers = {
                "Content-Type": "application/json",
                "X-Auth-Token": ""
            }
            #X-Auth-Token参数设置
            if case_value[weight_auth] == "null":
                headers["X-Auth-Token"] = ""
            elif case_value[weight_auth] == "miss":
                headers.pop("X-Auth-Token")
            else:
                X_Auth_Token = open(case_value[weight_auth], "rb").read()
                headers["X-Auth-Token"] = X_Auth_Token
            body_data = {
                "text":"",
                "config": {
                }
            }
            #text参数设置
            if case_value[weight_text] == "null":
                body_data["text"] = ""
            elif case_value[weight_text] == "miss":
                body_data.pop("text")
            else:
                text = open(case_value[weight_text], "rb").read()
                print text
                body_data["text"] = text
            #configc参数设置
            if case_value[weight_config] == "null":
                body_data["config"] = ""
            elif case_value[weight_config] == "miss":
                body_data.pop("config")
            else:
                pass
            #config string类型参数设置
            for j in range(weight_config + 1, weight_speed):
                if case_value[j] == '':
                    pass
                elif case_value[j] == 'null':
                    body_data["config"][config_list[j]] = ''
                else:
                    body_data["config"][config_list[j]] = case_value[j]
            #config int类型参数设置
            for j in range(weight_speed, weight_Expect_status):
                if case_value[j] == '':
                    pass
                elif case_value[j] == 'null':
                    body_data["config"][config_list[j]] = ''
                elif case_value[j] == 'float':
                    body_data["config"][config_list[j]] = 2.3
                elif case_value[j] == 'string':
                    body_data["config"][config_list[j]] = "asdsadsa"
                else:
                    body_data["config"][config_list[j]] = int(case_value[j])
            print body_data
            body_string = json.dumps(body_data)
            reponse,status,time_used = asr_interface(url, body_string, headers)
            reponse = json.loads(reponse)
            try:
                ws.write(i,weight_status,str(status))
                if status == 200:
                    if case_value[weight_foramt]:
                        format = case_value[weight_foramt]
                    else:
                        format = "wav"
                    voice_data = reponse["result"]["data"]
                    voice_data = base64.b64decode(voice_data)
                    #保存合成的音频数据
                    with open('./result/http_' +  str(i) + "." + format, 'wb') as audio:
                        audio.write(voice_data)

                    if case_value[weight_Expect_status] == "200":
                        ws.write(i, weight_skip+1, label="PASS", style=excel_colour_pass())
                    else:
                        ws.write(i, weight_skip+1, label="FAILED", style=excel_colour_fail())
                else:
                    ErrorNo = reponse["error_code"]
                    ResMessage = reponse["error_msg"]
                    ws.write(i, weight_ErrorNo, str(ErrorNo))
                    ws.write(i, weight_ResMessage, ResMessage)
                    if case_value[weight_Expect_status] == str(status):
                        if case_value[weight_except_errorno] == str(ErrorNo):
                            if case_value[weight_Expect_ResMessage] == ResMessage:
                                ws.write(i, weight_skip + 1, label="PASS", style=excel_colour_pass())
                                continue
                    ws.write(i, weight_skip + 1, label="FAILED", style=excel_colour_fail())
            except Exception, e:
                print e
                print 123
    wb.save("./Result/audio/" + "HTTP_RESULT" + time.strftime('%Y-%m-%d-%H%M%S', time.localtime(
        time.time())) + ".xls")

def asr_interface(url, body, header):
    start_time = time.time()
    response = requests.post(url, body, headers=header, verify=False)
    post_time = time.time() - start_time
    response_body = response.content
    status = response.status_code
    print "返回状态码：" + "\n" + str(status)
    print "返回信息：" + "\n" + response_body
    return response_body, status,post_time

if __name__ == '__main__':
    TTS_TEST()