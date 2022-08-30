'''按人员名单顺序生成值班表'''
import calendar
import datetime
import json
import os
import requests
import openpyxl.styles
from openpyxl import Workbook, load_workbook, worksheet
import time


class Schedule(object):

    def __init__(self, month: int, personal_file_name: str, xlsx_file_name: str):
        self.current_time = time.localtime()
        self.year = time.strftime("%Y", time.localtime())
        self.month = month
        self.data_path = os.path.abspath(rf"personal_information\{personal_file_name}.txt")
        self.output_xlsx_path = os.path.abspath(f"{xlsx_file_name}.xlsx")

    def set_personal_information(self, func):
        def fun(*args,**kwargs):
            date_list = self.creation_date()
            data_path = self.data_path
            personal_list = []
            result_list = []
            new_file_list = []
            day_total = int(self.creation_date()[-1].split("月")[-1].split("日")[0])
            with open(data_path, 'r+', encoding='utf-8') as f:
                for line in f.readlines():
                    personal_list.append(line.strip())
            f.close()
            with open(data_path.replace("txt", "bak"), 'w', encoding='utf-8') as f:
                f.write('\n'.join(personal_list))
                f.close()
            current_num = len(personal_list)
            while current_num < day_total:
                result_list = personal_list + personal_list
                current_num = len(personal_list)

            num = func(date_list, personal_list)
            last_name = result_list[num]
            pointer = personal_list.index(last_name)
            new_file_list = personal_list[pointer + 1:] + personal_list[0:pointer]
            if new_file_list:
                with open(data_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(new_file_list))
                    f.close()

        return fun

    def creation_date(self):
        year = int(self.year)
        month = self.month
        date_list = []
        for i in range(calendar.monthrange(year, month)[1] + 1)[1:]:
            str1 = f"{year}年{month}月{i}日"
            date_list.append(str1)
        return date_list

    def is_holiday(self, year, month, day):
        response = requests.get(
            f"https://api.apihubs.cn/holiday/get?field=date,week,workday,holiday_recess&year={year}&month={year}{month}&cn=1&size=31")
        day_dict = json.loads(response.text).get("data").get("list")[::-1][day - 1]
        if day_dict.get("workday_cn") == "非工作日" and day_dict.get("holiday_recess_cn") == "假期节假日":
            return [True, day_dict.get("date_cn"), day_dict.get("week_cn")]
        return False

    @set_personal_information()
    def output_xlsx(self, date_list, personal_list):
        # print(date_list)
        # print(personal_list)
        output_xlsx_path = self.output_xlsx_path
        year = int(date_list[0][0:4])
        wb = load_workbook(output_xlsx_path)
        ws = wb.active
        fianl_row_num = int(date_list[-1].split("月")[-1].split("日")[0]) + 2
        id = 1
        for row_num, date, day in zip(range(3, 35), date_list, range(1, 33)):
            ws.cell(row_num, 1, value=date.split("年")[-1]).font = openpyxl.styles.Font(name=u'宋体', size=10,
                                                                                        bold=False, color='000000')
            weekday = datetime.date(year, self.month, int(date.split("月")[-1].split("日")[0])).weekday()
            holiday = self.is_holiday(self.year, self.month, day)
            if not holiday[0]:
                personal = personal_list[day - id]
                if weekday in [0, 1, 2, 3, 4]:
                    prefix = "夜班："
                    font_bold = False
                else:
                    prefix = "全天："
                    font_bold = True
                ws.cell(row_num, 3, value=f"{prefix}{personal}").font = openpyxl.styles.Font(name=u'宋体', size=10,
                                                                                             bold=font_bold,
                                                                                             color='000000')
            else:
                id -= 1
        # print(fianl_row_num)
        if 33 - fianl_row_num > 0:
            [ws.cell(33 - num, 1, value="\\") for num in range(33 - fianl_row_num)]
            [ws.cell(33 - num, 3, value="\\") for num in range(33 - fianl_row_num)]
        wb.save(output_xlsx_path)
        wb.close()
        print(f"{self.month}月值班表已填入excel表格！")
        return id


if __name__ == '__main__':
    month = input("请输入月份后按下enter：")
    file = input("请输入模板excel文件名（不含拓展名）后按下enter：")
    main_function = Schedule(int(month), personal_file_name="personal", xlsx_file_name=file)
    # print(main_function.creation_date())
    # print(main_function.get_personal_information())
    main_function.output_xlsx()
    input("按下回车即可退出")
