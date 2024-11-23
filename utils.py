import yaml
import time
import os

def check_suffix(filename, suffix):
    return os.path.splitext(filename)[-1] == '.' + suffix

def yaml2dict(file_path, encoding='utf-8'):
    if check_suffix(file_path, 'yaml'):
        with open(file_path, 'r', encoding=encoding) as f:
            return yaml.safe_load(f)
    else:
        raise ValueError("file suffix name not support")

def get_each_level(path):
    levels = []
    while True:
        levels = [os.path.basename(path)] + levels
        path = os.path.dirname(path)
        if levels[0] == '':
            levels = levels[1:]
            return levels

def combine_levels(levels):
    # 从第一个元素开始，逐步添加到路径中
    path = levels[0]  # 从文件名或目录名开始
    for level in levels[1:]:  # 从第二个元素开始，因为第一个元素已经在path中
        path = os.path.join(path, level)  # 将下一个级别添加到路径中
    return path

def remove_top_level(path):
    return combine_levels(get_each_level(path)[1:])

def timestr():
    return str(time.time())

if __name__ == '__main__':
    print(remove_top_level("haha/haha2/haha3.txt"))
