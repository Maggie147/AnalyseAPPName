#coding:utf-8
import re
import os
import xlwt

count = 0
APP = []

def get_app(buf):
    global count
    global APP
    rule1 = r"<link>(.*?)</link>"
    app_boxs = re.findall(rule1, buf, re.S | re.M)
    for app_box in app_boxs:
        app_tmp = app_box.split('/')[-1]
        app_name = app_tmp.split(']')[0]
        if not app_name:
            continue
        # print app_name
        APP.append(app_name)
        count  = count +1


def get_file(fname):
    try:
        with open(fname, 'r') as fObj:
            raw_data = fObj.read()
    except Exception as e:
        print e
        raw_data = ''
    return raw_data


def get_fileList(fpath):
    xmlfile = []
    for value1, value2, value3 in os.walk(fpath):
        # print value1, value1, value1
        if len(value3)==0:
            continue
        if value3[0].find(".xml")==-1:
            continue
        for i in value3:
            xmlfile.append(value1+"/"+i)
            # xmlfile.append(os.path.join(value1, i))
    return xmlfile


def deal_more_app(applist):
    new_app = []
    for app in applist:
        if app not in new_app:
            new_app.append(app)
    # print new_app
    # new_app = [app for app in applist if app not in new_app]
    return new_app


def write_excel(applist, fname='./app_top_500.xls'):
    # workbook = xlwt.open_workbook(fname)
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('test', cell_overwrite_ok=True)
    sheet.write(0, 0, 'APP_Name')  # 其中的'0-行, 0-列'指定表中的单元，'EnglishName'是向该单元写入的内容
    for i, value in enumerate(applist):
        sheet.write(i+1, 0, value)
    book.save(fname)


def main():
    fileList = get_fileList('./raw_file')

    for file in fileList:
        print "Analyse file: [%s]"% file
        data = get_file(file)
        if data:
            get_app(data)
    # print APP
    print len(APP)
    print count

    new_app = deal_more_app(APP)
    print "new_app len: ", len(new_app)

    write_excel(new_app)


if __name__ == '__main__':
    main()