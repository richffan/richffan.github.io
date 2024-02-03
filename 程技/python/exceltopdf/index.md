# ExcelToPDF

## Excel To PDF


&lt;!--more--&gt;

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
    def __init__(self, pathname, export=&#39;.&#39;):
        self._handle_postfix = [&#39;doc&#39;, &#39;docx&#39;, &#39;ppt&#39;, &#39;pptx&#39;, &#39;xls&#39;, &#39;xlsx&#39;]
        self._filename_list = list()
        self._export_folder = os.path.join(os.path.abspath(&#39;.&#39;), &#39;pdfconver&#39;)
        if not os.path.exists(self._export_folder):
                os.mkdir(self._export_folder)
        self._enumerate_filename(pathname)
     
    def _enumerate_filename(self, pathname):
        &#39;&#39;&#39;
        读取所有文件名
        &#39;&#39;&#39;
        full_pathname = os.path.abspath(pathname)
        if os.path.isfile(full_pathname):
            if self._is_legal_postfix(full_pathname):
                self._filename_list.append(full_pathname)
            else:
                raise TypeError(&#39;文件 {} 后缀名不合法！仅支持如下文件类型：{}。&#39;.format(pathname, &#39;、&#39;.join(self._handle_postfix)))
        elif os.path.isdir(full_pathname):
            for relpath, _, files in os.walk(full_pathname):
                for name in files:
                    filename = os.path.join(full_pathname, relpath, name)
                    if self._is_legal_postfix(filename):
                        self._filename_list.append(os.path.join(filename))
        else:
            raise TypeError(&#39;文件/文件夹 {} 不存在或不合法！&#39;.format(pathname))
 
    def _is_legal_postfix(self, filename):
        return filename.split(&#39;.&#39;)[-1].lower() in self._handle_postfix and not os.path.basename(filename).startswith(&#39;~&#39;)
     
    def run_conver(self):
        &#39;&#39;&#39;
        进行批量处理，根据后缀名调用函数执行转换
        &#39;&#39;&#39;
        print(&#39;需要转换的文件数：&#39;, len(self._filename_list))
        for filename in self._filename_list:
            postfix = filename.split(&#39;.&#39;)[-1].lower()
            funcCall = getattr(self, postfix)
            print(&#39;原文件：&#39;, filename)
            funcCall(filename)
        print(&#39;转换完成！&#39;)
     
    def doc(self, filename):
        &#39;&#39;&#39;
        doc 和 docx 文件转换
        &#39;&#39;&#39;
        name = os.path.basename(filename).split(&#39;.&#39;)[0] &#43; &#39;.pdf&#39;
        exportfile = os.path.join(self._export_folder, name)
        print(&#39;保存 PDF 文件：&#39;, exportfile)
        gencache.EnsureModule(&#39;{00020905-0000-0000-C000-000000000046}&#39;, 0, 8, 4)
        w = Dispatch(&#34;Word.Application&#34;)
        doc = w.Documents.Open(filename)
        doc.ExportAsFixedFormat(exportfile, constants.wdExportFormatPDF,
                Item=constants.wdExportDocumentWithMarkup,
                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
         
        w.Quit(constants.wdDoNotSaveChanges)
     
    def docx(self, filename):
        self.doc(filename)
     
    def xls(self, filename):
        &#39;&#39;&#39;
        xls 和 xlsx 文件转换
        &#39;&#39;&#39;
        name = os.path.basename(filename).split(&#39;.&#39;)[0] &#43; &#39;.pdf&#39;
        exportfile = os.path.join(self._export_folder, name)
        xlApp = DispatchEx(&#34;Excel.Application&#34;)
        xlApp.Visible = False   
        xlApp.DisplayAlerts = 0  
        books = xlApp.Workbooks.Open(filename,False)
        books.ExportAsFixedFormat(0, exportfile)
        books.Close(False)
        print(&#39;保存 PDF 文件：&#39;, exportfile)
        xlApp.Quit()
     
    def xlsx(self, filename):
        self.xls(filename)
     
    def ppt(self, filename):
        &#39;&#39;&#39;
        ppt 和 pptx 文件转换
        &#39;&#39;&#39;
        name = os.path.basename(filename).split(&#39;.&#39;)[0] &#43; &#39;.pdf&#39;
        exportfile = os.path.join(self._export_folder, name)
        gencache.EnsureModule(&#39;{00020905-0000-0000-C000-000000000046}&#39;, 0, 8, 4)
        p = Dispatch(&#34;PowerPoint.Application&#34;)
        ppt = p.Presentations.Open(filename, False, False, False)
        ppt.ExportAsFixedFormat(exportfile, 2, PrintRange=None)
        print(&#39;保存 PDF 文件：&#39;, exportfile)
        p.Quit()
     
    def pptx(self, filename):
        self.ppt(filename)
 
 
 
if __name__ == &#34;__main__&#34;:
     
    ## 支持文件夹批量导入
    folder = &#39;tmp&#39;
    pathname = os.path.join(os.path.abspath(&#39;.&#39;), folder)
     
    ## 也支持单个文件的转换
    ## pathname = &#39;test.doc&#39;
 
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
    def __init__(self, pathname,sheetnum, export=&#39;.&#39;):
        self.sheetnum = sheetnum
        self._handle_postfix = [&#39;doc&#39;, &#39;docx&#39;, &#39;ppt&#39;, &#39;pptx&#39;, &#39;xls&#39;, &#39;xlsx&#39;]
        self._filename_list = list()
        self._export_folder = os.path.join(os.path.abspath(&#39;.&#39;), &#39;pdfconver&#39;)
        if not os.path.exists(self._export_folder):
            os.mkdir(self._export_folder)
        self._enumerate_filename(pathname)
 
    def _enumerate_filename(self, pathname):
        &#39;&#39;&#39;
        读取所有文件名
        &#39;&#39;&#39;
        full_pathname = os.path.abspath(pathname)
        if os.path.isfile(full_pathname):
            if self._is_legal_postfix(full_pathname):
                self._filename_list.append(full_pathname)
            else:
                raise TypeError(&#39;文件 {} 后缀名不合法！仅支持如下文件类型：{}。&#39;.format(pathname, &#39;、&#39;.join(self._handle_postfix)))
        elif os.path.isdir(full_pathname):
            for relpath, _, files in os.walk(full_pathname):
                for name in files:
                    filename = os.path.join(full_pathname, relpath, name)
                    if self._is_legal_postfix(filename):
                        self._filename_list.append(os.path.join(filename))
        else:
            raise TypeError(&#39;文件/文件夹 {} 不存在或不合法！&#39;.format(pathname))
 
    def _is_legal_postfix(self, filename):
        return filename.split(&#39;.&#39;)[-1].lower() in self._handle_postfix and not os.path.basename(filename).startswith(
            &#39;~&#39;)
 
    def run_conver(self):
        &#39;&#39;&#39;
        进行批量处理，根据后缀名调用函数执行转换
        &#39;&#39;&#39;
        print(&#39;需要转换的文件数：&#39;, len(self._filename_list))
        for filename in self._filename_list:
            postfix = filename.split(&#39;.&#39;)[-1].lower()
            funcCall = getattr(self, postfix)
            print(&#39;原文件：&#39;, filename)
            funcCall(filename)
        print(&#39;转换完成！&#39;)
 
    def xls(self, filename):
        &#39;&#39;&#39;
        xls 和 xlsx 文件转换
        &#39;&#39;&#39;
        xlApp = DispatchEx(&#34;Excel.Application&#34;)
        xlApp.Visible = False
        xlApp.DisplayAlerts = 0
        books = xlApp.Workbooks.Open(filename, False)
        ## 循环保存每一个sheet
        for i in range(1, self.sheetnum&#43;1):
            sheetName = books.Sheets(i).Name
            xlSheet = books.Worksheets(sheetName)
            name = sheetName &#43; &#39;.pdf&#39;
            exportfile = os.path.join(self._export_folder, name)
            xlSheet.ExportAsFixedFormat(0, exportfile)
            print(&#39;保存 PDF 文件：&#39;, exportfile)
        books.Close(False)
        xlApp.Quit()
 
    def xlsx(self, filename):
        self.xls(filename)
 
 
if __name__ == &#34;__main__&#34;:
    ## 支持单个文件的转换
    pathname = u&#39;原始数据.xlsx&#39;
    ## 获取到文件的sheet数
    b = xlrd.open_workbook(pathname)
    sheetnum = len(b.sheets())
    pdfConverter = PDFConverter(pathname, sheetnum)
    pdfConverter.run_conver()
```

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/python/exceltopdf/  

