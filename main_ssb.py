'''
    助力上海2022抗疫，解析随申办核酸结果截图
'''
import pandas as pd
import os
import datetime
import argparse
from tqdm import tqdm
from paddleocr import PaddleOCR
import re
import shutil

def parse_args():
    parser = argparse.ArgumentParser('training')
    parser.add_argument('--names', type=str, default=r'name_list.xlsx', help='select name_list.xlsx')
    parser.add_argument('--images', type=str, default=r'images', help='select image folder')
    return parser.parse_args()

def main():
    args = parse_args()
    # reader = easyocr.Reader(['ch_sim'], gpu=True) # this needs to run only once to load the model into memory

    excel_file = args.names
    img_folder = args.images
    print(f"You select -- file: {args.names},  image folder:  {args.images}")

    # 读取所有用户信息
    all_students = pd.read_excel(excel_file)
    has_id_info = '证件号码' in all_students.columns

    if has_id_info:
        all_students[["证件号码"]] = all_students[["证件号码"]].astype(str)
    else:
        print('\033[1;31;40m Warning： The file does not include 证件号码 column！\033[0m')
    all_students['test'] = 0  # 标记用户是否已经测试
    all_students['采样时间'] = ''
    all_students['检测结果'] = ''
    all_students['from'] = ''  # 记录来源于健康云还是随申办
    all_students['img'] = ''  # 该用户匹配的图片
    all_students['检测证件号码'] = ''  # 检测到的证件号码
    all_students['检测姓名'] = ''  # 检测到的姓名


    # 存储特殊用户信息
    ls_det_failed = []
    ls_jky = []  # 健康云图片
    ls_not_matching = []
    ls_warning = []

    ls = ['姓名', '证件号码', '采样时间', '检测结果']

    mapping = {'姓名': '姓名', '证件号码': '号码', '采样时间': '样时间', '检测结果': '结果'}
    data = pd.DataFrame(columns=ls + ['from', 'img'])  # 添加一列记录img name, 记录来源于健康云还是随申办

    ocr = PaddleOCR(lang='ch', det_db_box_thresh=0.5, use_gpu = False) #,
                            # det_model_dir=r'./model/det/ch/ch_PP-OCRv2_det_infer',
                            # rec_model_dir=r'./model/rec/ch/ch_PP-OCRv2_rec_infer',
                            # cls_model_dir=r'./model/cls/ch_ppocr_mobile_v2.0_cls_infer')  # need to run only once to download and load model into memory
    imgs = os.listdir(img_folder)
    for img_path in tqdm(imgs):
        # 遍历每一张图片
        # img = imgs[0]
        # paddleocr
        result_o = ocr.ocr(os.path.join(img_folder, img_path), cls=False)
        result = [line[1][0] for line in result_o]
        if '样本编码' in result:
            ls_jky.append(img_path)
            continue # 来自于健康云的图片跳过

        # continue
        # easyocr
        # result = reader.readtext(os.path.join(img_folder, img_path), detail = 0)
        cur = {}
        for col in ls:  # ls = ['姓名', '证件号码', '采样时间', '检测结果']
            for i, res in enumerate(result):

                # 采用更加简短的形式匹配，因为极端情况下，字体识别错误
                col_short = mapping[col]

                # 如果没有检测到，则跳过
                if col_short not in res:
                    continue
                try:  # 理想情况下，i+1项就是我们需要的结果
                    if i + 1 == len(result):  # 到达末尾，可以尝试将上一个元素作为结果
                        cur[col] = result[i - 1]
                    else:
                        cur[col] = result[i + 1]

                    # 修复身份证号篡位问题可能出现在证件号码前面的问题
                    if col == '证件号码':
                        if not bool(re.search(r'\d', cur[col])) or bool(re.search(r'[\u4e00-\u9fa5]',cur[col])):  # 如果检测出的证件号码没有数字或者含有汉字
                            if bool(re.search(r'\d', result[i - 1])) and not bool(re.search(r'[\u4e00-\u9fa5]', result[i - 1])):  # 但是i前面有数字且无汉字，则替换为该数字
                                cur[col] = result[i - 1]
                            elif bool(re.search(r'\d', result[i + 2])) and not bool(re.search(r'[\u4e00-\u9fa5]', result[i + 2])):  # 有时竟然可以i+2才可以找到
                                cur[col] = result[i + 2]
                            else:  # 寻找到的i前后都没有数字，则无证件号码,设置为0
                                cur[col] = '000000000000000000'
                                # del cur[col]
                                # raise ValueError
                        cur[col] = cur[col].upper()  # 如果存在x,则转化为大写

                    # 修复检测结果篡位问题
                    if col == '检测结果' and '性' not in cur[col]:
                        if '性' in result[i - 1]:  # 查看result[i - 1]是否符合
                            cur[col] = result[i - 1]
                        else:  # 放弃检索
                            del cur[col]
                            raise ValueError

                    # 已经检查到col项目，开始检测下一个col
                    break

                except (ValueError, IndexError):
                    break

            # 如果没有检测到col，则break
            if col not in cur.keys():
                break

        # 如果当前img无法完全检测到['姓名', '证件号码', '采样时间', '检测结果']， 跳转到下一个
        print(cur, img_path)
        if len(cur) != 4:
            ls_det_failed.append(img_path)
            continue

        # 修正日期：日与时之间无空格
        if cur['采样时间'][10] != ' ':
            hour_start = cur['采样时间'].find(':') - 2
            cur['采样时间'] = cur['采样时间'][:hour_start] + ' ' + cur['采样时间'][hour_start:]

        # 添加当前img信息
        cur['img'] = img_path
        cur['from'] = '健康云' if '样本编码' in result else '随申办'  # 健康云中有'样本编码'

        # 加入dataframe
        data = data.append(cur, ignore_index=True)
        #
    # 记录是否能与name_list信息match
    data['matching'] = 0

    # 开始匹配
    for ind_img, row in data.iterrows():
        # ind, row = next(data.iterrows())
        name = row['姓名']
        id = row['证件号码']

        if has_id_info:
            # 匹配id后4位
            ind_id_postfix = all_students['证件号码'].str.endswith(id[-4:])
        # 匹配name_postfix
        # name_postfix = name.split('*')[-1]
        name_postfix = re.sub(r'[^\u4e00-\u9fa5]', '', name)  # 只保留中文
        ind_name_postfix = all_students['姓名'].str.endswith(name_postfix)

        # 以下每一项中，有且只有一行符合就算符合。有一项符合即可算匹配成功
        # ind_id_postfix： 证件号码后四位匹配，且仅匹配到一位用户
        # ind_name_postfix： 姓名匹配，且仅匹配到一位用户
        if has_id_info:
            matching_ls = [ind_id_postfix, ind_name_postfix]
        else:
            matching_ls = [ind_name_postfix]

        for ind in matching_ls:
            if ind.sum() == 1:
                # 设置该用户已经测试核酸，并附上图片名称
                all_students.loc[ind.to_list().index(True), 'test'] = 1
                all_students.loc[ind.to_list().index(True), '采样时间'] = row['采样时间']
                all_students.loc[ind.to_list().index(True), '检测结果'] = row['检测结果']
                all_students.loc[ind.to_list().index(True), 'from'] = row['from']
                all_students.loc[ind.to_list().index(True), '检测证件号码'] = row['证件号码']
                all_students.loc[ind.to_list().index(True), '检测姓名'] = row['姓名']
                all_students.loc[ind.to_list().index(True), 'img'] = row['img']

                # 设置该img已经匹配
                data.loc[ind_img, 'matching'] = 1
                break

    # 查找还没有测试的用户
    ls_warning = all_students.loc[all_students['test'] == 0, '姓名'].tolist()
    # 查找不match的img
    ls_not_matching = data.loc[data['matching'] == 0, 'img'].tolist()

    # 将ls_det_failed，ls_not_matching对应的图片复制到放置到单独的文件夹
    if len(ls_det_failed) != 0:
        os.mkdir('images_detected_failed')
        for file_name in ls_det_failed:
            shutil.copyfile(os.path.join(img_folder, file_name), os.path.join('images_detected_failed', file_name))

    if len(ls_not_matching) != 0:
        os.mkdir('images_detected_successfully_but_not_matching')
        for file_name in ls_not_matching:
            shutil.copyfile(os.path.join(img_folder, file_name),  os.path.join('images_detected_successfully_but_not_matching', file_name))

    if len(ls_jky) != 0:
        os.mkdir('from_jiankangyun')
        for file_name in ls_jky:
            shutil.copyfile(os.path.join(img_folder, file_name),  os.path.join('from_jiankangyun', file_name))

    # print 检测结果
    print('\033[1;31;40m Detect failed imgs:\033[0m')
    print(ls_det_failed)
    print('\033[1;31;40m Detected, but not found in excel file:\033[0m')
    print(ls_not_matching)
    print('\033[1;31;40m These students deserve a warning:\033[0m')
    print(ls_warning)
    print('\033[1;31;40m These images from 健康云:\033[0m')
    print(ls_jky)

    # 导出文件
    output = f"{os.path.basename(args.names).split('.')[0]}_{datetime.datetime.now().strftime('%Y%m%d%H%M')}.xlsx"
    with pd.ExcelWriter(output) as writer:
        all_students.to_excel(writer, sheet_name='overview')
        data.to_excel(writer, sheet_name='img_detection_res')
        pd.DataFrame({'warning': [ls_warning], 'imgs detect failed': [ls_det_failed], 'detecte successfully but imgs not matching': [ls_not_matching], '这些图片来自健康云': [ls_jky]}).to_excel(writer, sheet_name='warning')
    print('Done.')

    # 祝福
    print('Keep healthy! Good luck to everyone.')

main()