#coding:utf-8
import re
import os
import xlwt

class DealWithExcel(object):
    def __init__(self, name):
        self.fname = name

    def write_excel(self, applist):
        # workbook = xlwt.open_workbook(self.fname)
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        sheet = book.add_sheet('test', cell_overwrite_ok=True)
        sheet.write(0, 0, 'APP_Name')           # 其中的'0-行, 0-列'指定表中的单元，'EnglishName'是向该单元写入的内容
        for i, value in enumerate(applist):
            sheet.write(i+1, 0, value)
        book.save(self.fname)


class DealWithRawFile(object):
    count = 0
    APP = []

    def __init__(self, fpath):
        self.main_path = fpath

    def _get_filelist(self):
        file_paths = []

        for value1, value2, value3 in os.walk(self.main_path):
            # print value1, '\n',value2, '\n', value3
            # print '--'*20
            if len(value3)==0:
                continue
            if value3[0].find(".xml")==-1:
                continue
            for i in value3:
                file_paths.append(os.path.join(value1, i))
        return file_paths

    def _get_file(self, fname):
        try:
            with open(fname, 'r') as fObj:
                raw_data = fObj.read()
        except Exception as e:
            print e
            raw_data = ''
        return raw_data

    def _analyse_app(self, buf):
        rule1 = r"<link>(.*?)</link>"
        app_boxs = re.findall(rule1, buf, re.S | re.M)
        for app_box in app_boxs:
            app_tmp = app_box.split('/')[-1]
            app_name = app_tmp.split(']')[0]
            if not app_name:
                continue
            DealWithRawFile.count += 1
            DealWithRawFile.APP.append(app_name)


    def _deal_double(self, applist):
        new_app = []
        for app in applist:
            if app not in new_app:
                new_app.append(app)
        # print new_app
        # new_app = [app for app in applist if app not in new_app]
        return new_app

    def get_app(self):
        files = self._get_filelist()
        for file in files:
            print "Analyse file: ", file

            buff = self._get_file(file)
            if not buff:
                print "no data in file: ", file
                continue
            self._analyse_app(buff)

        APP2 = self._deal_double(DealWithRawFile.APP)

        print "count  : ", DealWithRawFile.count
        print "new_len: ", len(APP2)

        return APP2


def main():
    spath = './raw_file'
    dpath = './test.xls'

    app_obj = DealWithRawFile(spath)
    apps = app_obj.get_app()
    if not apps:
        print "get app name failed!!!"

    excel_obj = DealWithExcel(dpath)
    excel_obj.write_excel(apps)


if __name__ == '__main__':
    main()