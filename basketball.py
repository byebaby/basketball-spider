import re
import time

import requests_html
import xlsxwriter as xw
from requests.adapters import HTTPAdapter

headers = {
    'User-Agent': requests_html.user_agent(),
    'Accept': '*/*',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh-HK;q=0.8,en-GB;q=0.6,en-US;q=0.4',
}


def create_execl(play_id, workbook, session):
    for p_key, pid in enumerate(play_id, start=1):
        # 新建工作薄
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 15)
        worksheet.set_column('H:H', 15)
        worksheet.set_column('I:I', 15)
        worksheet.set_column('J:J', 15)
        # 框架url
        url = 'http://nba.win0168.com/cn/Tech/TechTxtLive.aspx?matchid=%s' % pid
        r = session.get(url, timeout=6)
        # ------------总比分 ----------------#
        for tr_key, tr_val in enumerate(r.html.find('table.t_bf > tr'), start=15):
            for td_key, td_val in enumerate(tr_val.find('td'), start=0):
                # 写入数据
                worksheet.write(tr_key, td_key, td_val.text)

        # ------------技术统计----------------#

        # 返回函数 行标
        def count():
            temp = 0

            def counter(num=1):
                nonlocal temp
                temp += num
                return temp

            return counter

        # 光标行定位
        cursor = count()
        # 生成表格框架
        for h2_key, h2_val in enumerate(r.html.find('h2.rs_tit'), start=0):
            # 写入队名
            cell_format = workbook.add_format({'bold': True})
            worksheet.write(cursor(3), 7, h2_val.text, cell_format)
            for tr_key, tr_val in enumerate(r.html.find('#techMainDiv tbody')[h2_key].find('tr'), start=0):
                # 换行
                cursor(1)
                # 合并列 y轴增加
                line = 0
                for td_key, td_val in enumerate(tr_val.find('td')):
                    # 写入列名
                    y = 7 + line + td_key
                    if td_val.attrs.get('colspan'):
                        merge_format = workbook.add_format({'align': 'center'})
                        worksheet.merge_range(cursor(0), y, cursor(0), y + int(td_val.attrs.get('colspan')) - 1,
                                              td_val.text, merge_format)
                        line += int(td_val.attrs.get('colspan')) - 1
                    else:
                        worksheet.write(cursor(0), y, td_val.text)
                if tr_key == 3:
                    # 最多上场12人 手动增加
                    cursor(9)
        # 内容url
        url = 'http://nba.win0168.com/jsData/tech/%s/%s/%s.js?flesh=0.3349029049499843' % (
            pid[0], pid[1:3], pid)
        r = session.get(url, timeout=6)
        # 双方得分统计插入
        score_statistics = r.text[:r.text.find('$')].rsplit('^', 6)
        del score_statistics[0]
        ss_position = ['I44', 'J44', 'I45', 'J45', 'I46', 'J46']
        for ss_key, ss_val in enumerate(score_statistics):
            worksheet.write(ss_position[ss_key], ss_val)
        # 左队插入
        leftteam_data = list(map(lambda x: x[x.find('^') + 1:].split('^'),
                                 r.text[r.text.find('$'):r.text.rfind('$')].split("!")))
        # 队员技术统计
        score_team = leftteam_data[:-2]
        for row_key, row_val in enumerate(score_team, start=6):
            # 数据格式化
            del row_val[1]
            del row_val[1]
            del row_val[1]
            row_val[3] += '-' + row_val[4]
            del row_val[4]
            row_val[4] += '-' + row_val[5]
            del row_val[5]
            row_val[5] += '-' + row_val[6]
            del row_val[6]
            format_val = {'G': '后卫', 'C': '中锋', 'F': '前锋'}
            row_val[1] = format_val.get(row_val[1]) if format_val.get(row_val[1]) else row_val[1]
            for line_key, line_val in enumerate(row_val, start=7):
                worksheet.write(row_key, line_key, line_val)
        # 总计
        total_team = leftteam_data[-2]
        total_team[0] += '-' + total_team[1]
        del total_team[1]
        total_team[1] += '-' + total_team[2]
        del total_team[2]
        total_team[2] += '-' + total_team[3]
        del total_team[3]
        for line_key, line_val in enumerate(total_team, start=10):
            worksheet.write(18, line_key, line_val)
        worksheet.write('K20', leftteam_data[-1][0] + '%')
        worksheet.write('L20', leftteam_data[-1][1] + '%')
        worksheet.write('M20', leftteam_data[-1][2] + '%')
        worksheet.write('U20', '总失误:' + leftteam_data[-1][4])
        # 右队插入
        rightteam_data = list(map(lambda x: x[x.find('^') + 1:].split('^'),
                                  r.text[r.text.rfind('$'):].split("!")))
        # 队员技术统计
        score_team = rightteam_data[:-2]
        for row_key, row_val in enumerate(score_team, start=25):
            # 数据格式化
            del row_val[1]
            del row_val[1]
            del row_val[1]
            row_val[3] += '-' + row_val[4]
            del row_val[4]
            row_val[4] += '-' + row_val[5]
            del row_val[5]
            row_val[5] += '-' + row_val[6]
            del row_val[6]
            format_val = {'G': '后卫', 'C': '中锋', 'F': '前锋'}
            row_val[1] = format_val.get(row_val[1]) if format_val.get(row_val[1]) else row_val[1]
            for line_key, line_val in enumerate(row_val, start=7):
                worksheet.write(row_key, line_key, line_val)
            # 总计
        total_team = rightteam_data[-2]
        total_team[0] += '-' + total_team[1]
        del total_team[1]
        total_team[1] += '-' + total_team[2]
        del total_team[2]
        total_team[2] += '-' + total_team[3]
        del total_team[3]
        for line_key, line_val in enumerate(total_team, start=10):
            worksheet.write(37, line_key, line_val)
        worksheet.write('K39', rightteam_data[-1][0] + '%')
        worksheet.write('L39', rightteam_data[-1][1] + '%')
        worksheet.write('M39', rightteam_data[-1][2] + '%')
        worksheet.write('U39', '总失误:' + rightteam_data[-1][4])
        print('完成', p_key, '场')
        time.sleep(1)


def regular_season(requests_date):
    with requests_html.HTMLSession() as session:
        session.headers = headers
        session.mount('http://', HTTPAdapter(max_retries=5))
        session.mount('https://', HTTPAdapter(max_retries=5))
        # 抓取的常规赛年月
        url_format = 'http://nba.win0168.com/jsData/matchResult/%s/l1_1_20%s_10.js?version=2018112112' % (
            requests_date, requests_date[:2])
        r = session.get(url_format, timeout=6)
        # 年月数据格式化
        year_month = map(lambda x: x.split(','), r.html.search('ymList = [[{}]];')[0].split('],['))
        for ym in year_month:
            if ym == ['2018', '12']:
                return 0
            # 新建excel format: year - month.xlsx
            workbook = xw.Workbook(ym[0] + '-' + ym[1] + '.xlsx')
            url = 'http://nba.win0168.com/jsData/matchResult/%s/l1_1_%s_%s.js?version=2018112112' % (
                requests_date, ym[0], ym[1])
            r = session.get(url, timeout=6)
            # 该年月的比赛id
            play_id = map(lambda x: x.split(',')[0], r.html.search('arrData = [[{}]];')[0].split('],['))
            play_id = list(play_id)
            # 当前场次后无数据 截取list
            if ym == ['2018', '11']:
                play_id = play_id[:play_id.index('325827')]
            create_execl(play_id, workbook, session)
            # 关闭保存
            workbook.close()
            print('完成', ym[0], ym[1])


def playoffs(requests_date):
    with requests_html.HTMLSession() as session:
        session.headers = headers
        session.mount('http://', HTTPAdapter(max_retries=5))
        session.mount('https://', HTTPAdapter(max_retries=5))
        # 抓取季度数据
        r = session.get('http://nba.win0168.com/jsData/matchResult/%s/l1_2.js?version=2018112122' % requests_date,
                        timeout=6)
        # 季度数据格式化
        quarter = list(map(lambda x: x.split(',')[0],
                           re.split(",\[\[|[0-9]\],\[", r.html.search(",[[{}var")[0])))
        # 新建excel format: year - month.xlsx
        workbook = xw.Workbook(requests_date + '季后赛.xlsx')
        # 该年月的比赛id
        play_id = quarter
        create_execl(play_id, workbook, session)
        # 关闭保存
        workbook.close()
        print('完成', requests_date)


if __name__ == '__main__':
    dict_date = ['16-17', '17-18','18-19']
    for date in dict_date:
        regular_season(date)
    dict_date = ['16-17', '17-18']
    for date in dict_date:
        playoffs(date)
    # regular_season()
