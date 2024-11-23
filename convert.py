import utils
import os
import shutil
import zipfile
import re
import sys

import pandas as pd

CONFIG = None
QUICK = False
MATCH_PREFIX = "PLACE_HOLDER_"

def read_config():
    global config
    while True:
        if CONFIG is None:
            print("请输入配置文件的绝对路径:")
            str = input()
        else:
            str = CONFIG
        if os.path.exists(str):
            try:
                config = utils.yaml2dict(str)
                workPath = os.path.dirname(str)
                config['template'] = os.path.abspath(os.path.join(workPath, config['template']))
                config['data'] = os.path.abspath(os.path.join(workPath, config['data']))
                config['dst'] = os.path.abspath(os.path.join(workPath, config['dst']))
                if config['template'] is None or not os.path.exists(config['template']):
                    print(f"模板路径 {config['template']} 不存在, 请重试:\n")
                    continue
                if config['data'] is None or not os.path.exists(config['data']):
                    print(f"数据路径 {config['data']} 不存在, 请重试:\n")
                    continue
                if config['dst'] is None:
                    print(f"目标目录 {config['dst']} 不存在, 请重试:\n")
                    continue
                os.makedirs(config['dst'], exist_ok=True)
                
                print("请确认以下信息(输入 Y 并回车确认, 输入其它内容取消操作):")
                print(f"模板路径 {config['template']}")
                print(f"数据路径 {config['data']}")
                print(f"目标目录 {config['dst']}")
                print(f"命名方式 {config['filename'] if config['filename'] else '数字 id 命名'}")
            except (KeyError, ValueError) as e:
                if isinstance(e, (KeyError)):
                    print(f"配置文件 {str} 缺少关键项, 请对照标准格式检查")
                if isinstance(e, ValueError):
                    print(f"配置文件 {str} 不是一个 yaml 文件")
                sys.exit()
            if not QUICK:
                r = input()
                if r == 'Y':
                    return 
                else:
                    sys.exit()
            else:
                return 
        else:
            print(f"配置文件路径不存在, 请重试:\n")

def unzip_docx(docx_path):
    print("正在解压 docx 文档")
    if not utils.check_suffix(docx_path, 'docx'):
        print(f"{docx_path} 后缀名不是 .docx, 请检查你的配置文件")
        sys.exit()
    docx_dir = os.path.dirname(docx_path)
    temp_basename = utils.timestr()
    temp_dir = os.path.join(docx_dir, temp_basename)
    temp_path = temp_basename + '.zip'
    shutil.copy(docx_path, temp_path)
    
    with zipfile.ZipFile(temp_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    os.remove(temp_path)
    return temp_dir

def read_xlsx(xlsx_path) -> pd.DataFrame:
    print("正在读取 xlsx 数据表")
    if not utils.check_suffix(xlsx_path, 'xlsx'):
        print(f"{xlsx_path} 后缀名不是 .xlsx, 请检查你的配置文件")
        sys.exit()
    xls = pd.ExcelFile(xlsx_path)
    if len(xls.sheet_names) > 1:
        print(f"{xlsx_path} 的 sheet 数量过多, 请保证它只有一个 sheet")
        sys.exit()
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    
    match_template = re.compile(MATCH_PREFIX + r'([1-9]\d*)')
    for index, key in enumerate(df.keys(), start=1):
        if match_template.match(key) and int(match_template.match(key).group(1)) == index:
            continue
        print(f"数据表格表头有误, 第 {index} 列应该为 PLACE_HOLDER_{index}, 但你的表格是 {key}")
        sys.exit()
    
    return df
    
def read_document_xml(xml_root, encoding='utf-8'):
    document_xml_path = os.path.join(xml_root, 'word', 'document.xml')
    with open(document_xml_path, 'r', encoding=encoding) as f:
        return f.read()

def write_document_xml(xml_root, s, encoding='utf-8'):
    document_xml_path = os.path.join(xml_root, 'word', 'document.xml')
    with open(document_xml_path, 'w', encoding=encoding) as f:
        f.write(s)

def zip_docx(xml_root, docx_path):
    # ziph是zipfile.ZipFile对象
    with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(xml_root):
            for file in files:
                zip_path = os.path.relpath(os.path.join(root, file), os.path.join(xml_root, os.path.pardir))
                path_without_root = utils.remove_top_level(zip_path)
                zipf.write(os.path.join(root, file), path_without_root)

def gen_docx(template, df:pd.DataFrame, dst, filename=None):
    if not filename:
        filename = ""

    for row_index, row in enumerate([row for _, row in df.iterrows()], start=1):
        print(f"正在处理数据表, 第 {row_index} 项")
        xml_root = unzip_docx(template)
        xml_content = read_document_xml(xml_root)
        this_filename = str(row_index) + filename
        for col_index, val in enumerate(row, start=1):
            replace_template = MATCH_PREFIX + str(col_index)
            if xml_content.find(replace_template) == -1:
                print(f"警告, {replace_template} 在模板中没有找到")
            xml_content = xml_content.replace(replace_template, val)
            this_filename = this_filename.replace(replace_template, val)
        write_document_xml(xml_root, xml_content)
        zip_docx(xml_root, os.path.join(dst, this_filename + '.docx'))
        shutil.rmtree(xml_root)

read_config()
df = read_xlsx(config['data'])
gen_docx(config['template'], df, config['dst'], config['filename'])


