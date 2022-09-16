# -*- coding: UTF-8 -*-
# by:Caiqiancheng
# Date:2022/9/16
import configparser
import os


def read_config(section, filename="config.ini"):
    """
    读取配置文件,返回键值字典
    """
    conf = configparser.ConfigParser()
    with open(file="./conf/{}".format(filename), mode="r", encoding="utf-8") as f:
        conf.read_file(f)
    return dict(conf.items(section))


def get_file_list(directory_path, mode="path"):
    """
    获取指定目录下的文件名列表或文件完整路径列表
    :param directory_path: 指定目录
    :param mode: "path"-返回完整路径列表 "name"-返回文件名列表
    :return: 文件名列表或文件完整路径列表
    """
    filename_list = []
    filepath_list = []
    [filename_list.append(filename) for parent, dirname, filename in os.walk(directory_path, topdown=True)]
    [filepath_list.append(f"{directory_path}{filename}") for filename in filename_list[0]]
    if mode == "path":
        return filepath_list
    elif mode == "name":
        return filename_list
    else:
        return []
