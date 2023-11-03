# ExcelToPDF

## Excel To PDF


<!--more-->

### 安装

```python
pip install pywin32
```

### office文件 (word, ppt, excel等) 转为pdf

```Python
#-*- coding:utf-8 -*-
import os
from win32com.client import Dispatch, constants, gencache, DispatchEx
 
 
class PDFConverter:
    def __init__(self, pathname, export='.'):
        self._handle_postfix = ['doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx']
        self._filename_list = list()
        self._export_folder = os.path.join(os.path.abspath('.'), 'pdfconver')
        if not os.path.exists(self._export_folder):
                os.mkdir(self._export_folder)
        self._enumerate_filename(pathname)
     
    def _enumerate_filename(self, pathname):
        '''
        读取所有文件名
        '''
        full_pathname = os.path.abspath(pathname)
        if os.path.isfile(full_pathname):
            if self._is_legal_postfix(full_pathname):
                self._filename_list.append(full_pathname)
            else:
                raise TypeError('文件 {} 后缀名不合法！仅支持如下文件类型：{}。'.format(pathname, '、'.join(self._handle_postfix)))
        elif os.path.isdir(full_pathname):
            for relpath, _, files in os.walk(full_pathname):
                for name in files:
                    filename = os.path.join(full_pathname, relpath, name)
                    if self._is_legal_postfix(filename):
                        self._filename_list.append(os.path.join(filename))
        else:
            raise TypeError('文件/文件夹 {} 不存在或不合法！'.format(pathname))
 
    def _is_legal_postfix(self, filename):
        return filename.split('.')[-1].lower() in self._handle_postfix and not os.path.basename(filename).startswith('~')
     
    def run_conver(self):
        '''
        进行批量处理，根据后缀名调用函数执行转换
        '''
        print('需要转换的文件数：', len(self._filename_list))
        for filename in self._filename_list:
            postfix = filename.split('.')[-1].lower()
            funcCall = getattr(self, postfix)
            print('原文件：', filename)
            funcCall(filename)
        print('转换完成！')
     
    def doc(self, filename):
        '''
        doc 和 docx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = os.path.join(self._export_folder, name)
        print('保存 PDF 文件：', exportfile)
        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
        w = Dispatch("Word.Application")
        doc = w.Documents.Open(filename)
        doc.ExportAsFixedFormat(exportfile, constants.wdExportFormatPDF,
                Item=constants.wdExportDocumentWithMarkup,
                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
         
        w.Quit(constants.wdDoNotSaveChanges)
     
    def docx(self, filename):
        self.doc(filename)
     
    def xls(self, filename):
        '''
        xls 和 xlsx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = os.path.join(self._export_folder, name)
        xlApp = DispatchEx("Excel.Application")
        xlApp.Visible = False   
        xlApp.DisplayAlerts = 0  
        books = xlApp.Workbooks.Open(filename,False)
        books.ExportAsFixedFormat(0, exportfile)
        books.Close(False)
        print('保存 PDF 文件：', exportfile)
        xlApp.Quit()
     
    def xlsx(self, filename):
        self.xls(filename)
     
    def ppt(self, filename):
        '''
        ppt 和 pptx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = os.path.join(self._export_folder, name)
        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
        p = Dispatch("PowerPoint.Application")
        ppt = p.Presentations.Open(filename, False, False, False)
        ppt.ExportAsFixedFormat(exportfile, 2, PrintRange=None)
        print('保存 PDF 文件：', exportfile)
        p.Quit()
     
    def pptx(self, filename):
        self.ppt(filename)
 
 
 
if __name__ == "__main__":
     
    ## 支持文件夹批量导入
    folder = 'tmp'
    pathname = os.path.join(os.path.abspath('.'), folder)
     
    ## 也支持单个文件的转换
    ## pathname = 'test.doc'
 
    pdfConverter = PDFConverter(pathname)
    pdfConverter.run_conver()
```

### excel的不同sheet存为pdf

```Python
#-*- coding:utf-8 -*-
 
import os
from win32com.client import Dispatch, constants, gencache, DispatchEx
import xlrd
 
class PDFConverter:
    def __init__(self, pathname,sheetnum, export='.'):
        self.sheetnum = sheetnum
        self._handle_postfix = ['doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx']
        self._filename_list = list()
        self._export_folder = os.path.join(os.path.abspath('.'), 'pdfconver')
        if not os.path.exists(self._export_folder):
            os.mkdir(self._export_folder)
        self._enumerate_filename(pathname)
 
    def _enumerate_filename(self, pathname):
        '''
        读取所有文件名
        '''
        full_pathname = os.path.abspath(pathname)
        if os.path.isfile(full_pathname):
            if self._is_legal_postfix(full_pathname):
                self._filename_list.append(full_pathname)
            else:
                raise TypeError('文件 {} 后缀名不合法！仅支持如下文件类型：{}。'.format(pathname, '、'.join(self._handle_postfix)))
        elif os.path.isdir(full_pathname):
            for relpath, _, files in os.walk(full_pathname):
                for name in files:
                    filename = os.path.join(full_pathname, relpath, name)
                    if self._is_legal_postfix(filename):
                        self._filename_list.append(os.path.join(filename))
        else:
            raise TypeError('文件/文件夹 {} 不存在或不合法！'.format(pathname))
 
    def _is_legal_postfix(self, filename):
        return filename.split('.')[-1].lower() in self._handle_postfix and not os.path.basename(filename).startswith(
            '~')
 
    def run_conver(self):
        '''
        进行批量处理，根据后缀名调用函数执行转换
        '''
        print('需要转换的文件数：', len(self._filename_list))
        for filename in self._filename_list:
            postfix = filename.split('.')[-1].lower()
            funcCall = getattr(self, postfix)
            print('原文件：', filename)
            funcCall(filename)
        print('转换完成！')
 
    def xls(self, filename):
        '''
        xls 和 xlsx 文件转换
        '''
        xlApp = DispatchEx("Excel.Application")
        xlApp.Visible = False
        xlApp.DisplayAlerts = 0
        books = xlApp.Workbooks.Open(filename, False)
        ## 循环保存每一个sheet
        for i in range(1, self.sheetnum+1):
            sheetName = books.Sheets(i).Name
            xlSheet = books.Worksheets(sheetName)
            name = sheetName + '.pdf'
            exportfile = os.path.join(self._export_folder, name)
            xlSheet.ExportAsFixedFormat(0, exportfile)
            print('保存 PDF 文件：', exportfile)
        books.Close(False)
        xlApp.Quit()
 
    def xlsx(self, filename):
        self.xls(filename)
 
 
if __name__ == "__main__":
    ## 支持单个文件的转换
    pathname = u'原始数据.xlsx'
    ## 获取到文件的sheet数
    b = xlrd.open_workbook(pathname)
    sheetnum = len(b.sheets())
    pdfConverter = PDFConverter(pathname, sheetnum)
    pdfConverter.run_conver()
```

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/python/exceltopdf/  

