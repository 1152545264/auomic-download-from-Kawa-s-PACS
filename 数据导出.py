from pywinauto import *
import pywinauto
import time
from pywinauto.controls.common_controls import DateTimePickerWrapper
import win32gui
import os
import pandas as pd

top_title = u'DICOM 查询/检索模块'


def start_process():
    path = r'D:\HIS\doctor\pacs\pacs\RSQM.exe'
    app = Application(backend='uia').start(path)
    time.sleep(2)
    app = Application().connect(path=path)
    window = app[top_title]

    work_station = window.ComboBox
    work_station.print_control_identifiers()
    work_station.select(1)
    # window.print_control_identifiers()
    return app, window


'''
病人id
0000345018
0000380721
'''


def set_start_time(window=None, app=None):
    # 获取顶层窗口句柄
    father_handle = win32gui.FindWindow(None, top_title)
    date_type = 'SysDateTimePick32'
    # 获取日期窗口句柄
    start_date_handle = win32gui.FindWindowEx(father_handle, None, date_type, None)

    # st_date_handle = pywinauto.findwindows.find_windows(title_re=u'.*起始日期.*')

    start_time_window = DateTimePickerWrapper(start_date_handle)
    start_time_window.set_time(year=2016, month=1, day=1)
    # return True


def input_query_information(index=0, window=None, app=None, patient_id='0000345018', ):
    # 1、设置病人id
    edit = window.edit0
    edit.set_edit_text(patient_id)
    # 2、设置起始时间
    set_start_time(window=window, app=app)
    # 3点击查询按钮
    search_button = window.window(title=u'查询(&Q)')
    time.sleep(1)

    retry_times = 2
    counts = 0

    while retry_times:
        retry_times -= 1
        search_button.click()
        search_button.wait('enabled ready', timeout=10)
        # time.sleep(4)
        counts = window.child_window(class_name='SysListView32').item_count()
        if counts > 0:
            break

    # 4。双击检索CT
    query_res = window.child_window(class_name='SysListView32')
    query_res.wait('exists enabled visible ready', timeout=5)

    # result_window = window.window(title_re='.*检索.*', class_name='Static')
    result_window = window.window(control_id=20032, class_name='Static')
    # 获取检索结果
    result_text = result_window.texts()[0]
    # TODO 根据检索结果对数据进行分类处理,此时需要去掉下面的if count==0 语句块

    if counts == 0:
        info = str(index + 1) + '******' + str(patient_id) + '***** ' + result_text + '\n'
        log_failed(info)
        return

    res = deal_all_ct(index, patient_id, query_res, window)
    return res


# 获取所有ct项目
def deal_all_ct(index, patient_id, query_res, window=None):
    counts = query_res.item_count()
    ct_index = []
    res = False
    for i in range(0, counts):
        # listview 的行和列
        item_mod = query_res.get_item(i, 4)
        item_type = query_res.get_item(i, 3)

        # 获取item对应的数据
        item_mod_data = item_mod.item_data()
        item_type_data = item_type.item_data()

        item_mod_text = item_mod_data['text']
        item_type_text = item_type_data['text']
        if item_type_text == 's':
            info = str(index + 1) + '******' + str(patient_id) + '***** ' + '森田CBCT\n'
            log_failed(info)
            res = False
            continue
        else:
            if item_mod_text == 'CT':
                ct_index.append(i)
                res = True

    if len(ct_index) > 1:
        info = str(index + 1) + '******' + str(patient_id) + '***** ' + '照了多次CBCT\n'
        log(info=info)

    elif len(ct_index) == 0:
        info = str(index + 1) + '******' + str(patient_id) + '***** ' + '没有卡瓦CBCT片子\n'
        log_failed(info=info)

    elif len(ct_index) > 0:
        for i in ct_index:
            deal_single_ct(i, query_res)

            # 获得检索按钮
            retrieve = window.window(title_re='.*检索.*', class_name='Button')
            # 等待检索按钮可用
            retrieve.wait('enabled visible ready', timeout=120)
            time.sleep(1)

            result_window = window.window(control_id=20032, class_name='Static')
            # 获取检索结果
            result_text = result_window.texts()[0]
            if '成功' not in result_text:
                info = str(index) + '******' + str(patient_id) + '******' + result_text
                log_failed(info)
        # result_text += '/n'

        # time.sleep(120)

    return res


def deal_single_ct(index, query_res):
    ct_item = query_res.get_item(index)
    time.sleep(1)
    ct_item.click(double=True)


def log_failed(info):
    with open('./failed.txt', 'a+', encoding='utf-8') as f:
        f.writelines(info)
        f.close()
        time.sleep(4)


def log(file='./multi_cbct.txt', info=None):
    with open(file, 'a+', encoding='utf-8') as f:
        f.writelines(info)
        f.close()
        time.sleep(4)


def deal_single_patient(index=0, window=None, app=None, patient_id='0000345018'):
    res = input_query_information(index, window, app=app, patient_id=patient_id)

    return res


id_path = r'./患者ID/患者ID.xlsx'
dicom_path = r'E:\Oral_image_root'


# index = 0


def deal_all(app, window):
    print(os.getcwd())
    df = pd.read_excel(id_path, dtype=object)
    df = df.iloc[:, 0]
    df = df.values
    df = sorted(df, key=lambda x: int(x))
    index = 0
    for patient_id in df:
        if index > 700:
            return
        if os.path.exists(os.path.join(dicom_path, patient_id)):
            index += 1
            print(f'已经检索了{index}个', f'检索id为：{patient_id}')
            continue

        res = deal_single_patient(index, window, app=app, patient_id=patient_id)
        index += 1
        print(f'已经检索了{index}个', f'检索id为：{patient_id}', f'检索结果为{res}')
        # if res:
        #     time.sleep(120)

        # 将检索失败的结果写入文件
        # if res is False:
        #     with open('./failed.txt', 'a+', encoding='utf-8') as f:
        #         f.writelines(str(index) + '******' + str(patient_id) + '\n')
        #         f.close()
        #         time.sleep(4)


app, window = start_process()
deal_all(app, window)
