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
    all_students['img'] = ''  # 该用户匹配的图片

    # 存储特殊用户信息
    ls_det_failed = []
    ls_not_matching = []
    ls_warning = []

    ls = ['姓名', '证件号码', '采样时间：', '检测结果：']
    data = pd.DataFrame(columns=ls + ['img'])  # 添加一列记录img name
    imgs = os.listdir(img_folder)

    ocr = PaddleOCR(lang='ch', det_db_box_thresh=0.5)  #,
                            # det_model_dir=r'./model/det/ch/ch_PP-OCRv2_det_infer',
                            # rec_model_dir=r'./model/rec/ch/ch_PP-OCRv2_rec_infer',
                            # cls_model_dir=r'./model/cls/ch_ppocr_mobile_v2.0_cls_infer')  # need to run only once to download and load model into memory

    for img_path in tqdm(imgs):
        # 遍历每一张图片
        # img = imgs[0]
        # paddleocr
        result_o = ocr.ocr(os.path.join(img_folder, img_path), cls=False)
        result = [line[1][0] for line in result_o]
        # continue
        # easyocr
        # result = reader.readtext(os.path.join(img_folder, img_path), detail = 0)
        cur = []
        for col in ls:
            # col = '姓名'
            try:
                ind_ls = result.index(col)   # 只查找第一个
                item = result[ind_ls + 1]
                # ind_ls_end = result.index(col_end)  # 避免检测出杂项,也寻找该项的下一项
                # item = ''.join(result[ind_ls + 1: ind_ls_end])

            except (ValueError, IndexError):
                # 检测结果这一项，容易出问题，特地再次检测
                if col == '检测结果：':
                    for i in result:
                        if '【' in i or '】' in i:
                            item = i
                            break

                else:
                    ls_det_failed.append(img_path)
                    continue
            cur.append(item)

        # # 如果检测失败
        # if len(cur) != len(ls):
        #     ls_det_failed.append(img_path)
        #     continue

        # 修正日期：日与时之间无空格
        if cur[2][10] != ' ':
            hour_start = cur[2].find(':') - 2
            cur[2] = cur[2][:hour_start] + ' ' + cur[2][hour_start:]

        # 加入dataframe
        data = data.append(dict(zip(ls + ['img'], cur + [img_path])), ignore_index=True)

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
        name_postfix = re.sub(r'[^\u4e00-\u9fa5]','',name)  # 只保留中文
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
                all_students.loc[ind.to_list().index(True), 'img'] = row['img']
                all_students.loc[ind.to_list().index(True), '采样时间'] = row['采样时间：']
                all_students.loc[ind.to_list().index(True), '检测结果：'] = row['检测结果：']

                # 设置该img已经匹配
                data.loc[ind_img, 'matching'] = 1
                break
        if data.loc[ind_img, 'matching'] != 1:  # 匹配失败
            ls_not_matching.append(row['img'])

    # 查找还没有测试的用户
    ls_warning = all_students.loc[all_students['test'] == 0, '姓名']

    print('\033[1;31;40m Detect failed imgs:\033[0m')
    print(ls_det_failed)
    print('\033[1;31;40m Detected, but not found in excel file:\033[0m')
    print(ls_not_matching)
    print('\033[1;31;40m These students deserve a warning:\033[0m')
    print(list(ls_warning))

    # 导出文件
    output = f"{os.path.basename(args.names).split('.')[0]}_{datetime.datetime.now().strftime('%Y%m%d%H%M')}.xlsx"
    with pd.ExcelWriter(output) as writer:
        all_students.to_excel(writer, sheet_name='overview')
        data.to_excel(writer, sheet_name='img_detection_res')
        pd.DataFrame({'warning': [list(ls_warning)], 'imgs detect failed': [ls_det_failed], 'detecte successfully but imgs not matching': [ls_not_matching]}).to_excel(writer, sheet_name='warning')
    print('Done.')

    # 祝福
    print('Keep healthy! Good luck to everyone.')

if __name__ == '__main__':
    main()