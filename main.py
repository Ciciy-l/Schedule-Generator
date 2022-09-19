# -*- coding: UTF-8 -*-
# by:Caiqiancheng
# Date:2022/9/16


from bin.Generation_cmd import generation
from bin.Generation_gui import generation_ui
from src.common import read_config


def get_start_mode():
    """
    从配置文件获取启动方式
    :return: cmd or gui
    """
    start_mode = read_config("start").get("mode")
    return start_mode


if __name__ == '__main__':
    if get_start_mode() == "cmd":
        generation()
    elif get_start_mode() == "gui":
        generation_ui()
