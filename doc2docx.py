import argparse
import os.path
from multiprocessing import Pool
from docx_func import get_docx_list
from win32com import client as wc  # 导入模块


def doc2docx(file_path):
    if 'docx' not in file_path:
        word = wc.Dispatch("Word.Application")  # 打开word应用程序
        doc = word.Documents.Open(file_path)  # 打开word文件
        doc.SaveAs("{}x".format(file_path), 12)  # 另存为后缀为".docx"的文件，其中参数12指docx文件
        doc.Close()  # 关闭原来word文件
        word.Quit()
        print('[转换完成]', file_path)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Project Cc2333 doc2docx')
    parser.add_argument('-s', '--source', action="store", help='reports dir')
    args = parser.parse_args()

    docx_list = get_docx_list(args.source)
    p = Pool(5)
    for d in docx_list:
        p.apply_async(doc2docx, args=(os.path.join(*d),))
    p.close()
    p.join()
