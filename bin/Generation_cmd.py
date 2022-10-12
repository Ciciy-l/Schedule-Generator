"""
命令行运行主程序
"""
from src.schedule import Schedule


def generation():
    while True:
        month = input("请输入月份后按下enter：")
        if month not in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]:
            print("输入有误！须输入正确的月份（1~12）")
        else:
            break
    while True:
        skip_holidays = input("请是否自动跳过节假日?请输入(y/n)后按下enter：")
        if skip_holidays not in ["y", "n"]:
            print("输入有误！须输入正确的参数（y or n）")
        else:
            break
    while True:
        file_template = input("请输入模板excel文件名（不含拓展名）后按下enter：")
        file_output = input("请输入生成excel文件名（不含拓展名）后按下enter：")
        try:
            main_function = Schedule(int(month), personal_file_name="personal", xlsx_template_name=file_template, skip=skip_holidays,xlsx_output_name=file_output)
        except PermissionError:
            input("Excel文件已被其他应用占用！请关闭占用软件后按下回车键重试…")
        else:
            break
    print("正在写入Excel，请稍后……")
    while True:
        try:
            main_function.output_xlsx(main_function.creation_date(), main_function.read_personal_information(),
                                      main_function.read_personal_information("leader"))
        except PermissionError:
            input("Excel文件已被其他应用占用！请关闭占用软件后按下回车键重试…")
            print("正在重试……")
        else:
            break
    input("排班表已填写完成，按下回车键即可退出")


if __name__ == '__main__':
    generation()
