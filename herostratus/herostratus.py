#!/usr/bin/python3
import os
from os import path
import pathlib
import magic
import re
from docx import Document
from pptx import Presentation
from PyPDF2 import PdfFileReader
import argparse
import warnings
import datetime as dt
from tqdm import tqdm
import dominate
import xml.etree.ElementTree as xee
import xlwt

warnings.filterwarnings('ignore')

def set_date_or_fail(date_time_string, date_time_format='%Y-%m-%d %H:%M:%S'):
    date_time_result = None
    try:
        date_time_result = dt.datetime.strptime(date_time_string, date_time_format)
    except ValueError:
        date_time_result = None
    return date_time_result

def fetch_or_fail(keyword, haystack):
    needle = re.search(keyword + '\s.*?\W+\s', haystack)
    if needle == None:
        return ''
    else:
        return needle.group(0)[len(keyword)]

class DocumentInfo():
    def __init__(self, path=''):
        self.path = path
        self.name = os.path.basename(path)
        self.author = None
        self.author_last = None
        self.date_create = None
        self.date_modified = None
        self.size = os.path.getsize(path)
        self.pages = None
        self.processed = False

    def set_date_create_from_file(self):
        fname = pathlib.Path(self.path)
        self.date_create = dt.datetime.fromtimestamp(fname.stat().st_ctime)

    def set_date_modified_from_file(self):
        fname = pathlib.Path(self.path)
        self.date_modified = dt.datetime.fromtimestamp(fname.stat().st_mtime)

    def to_xml_document(self):
        root = xee.Element("document")
        e_name = xee.SubElement(root, "name")
        e_name.text = self.name
        e_path = xee.SubElement(root, "path")
        e_path.text = self.path
        e_author = xee.SubElement(root, "author")
        e_author.text = self.author
        e_author_last = xee.SubElement(root, "author_last")
        e_author_last.text = self.author_last
        e_date_create = xee.SubElement(root, "date_create")
        e_date_create.text = '' if self.date_create == None else self.date_create.strftime("%m/%d/%Y, %H:%M:%S")
        e_date_modified = xee.SubElement(root, "date_modified")
        e_date_modified.text = '' if self.date_modified == None else self.date_modified.strftime("%m/%d/%Y, %H:%M:%S")
        e_pages = xee.SubElement(root, "pages")
        e_pages.text = str(self.pages)
        e_size = xee.SubElement(root, "size")
        e_size.text = str(self.size)
        return root

    def to_xml_file(self):
        root = xee.Element("file")
        e_name = xee.SubElement(root, "name")
        e_name.text = self.name
        e_path = xee.SubElement(root, "path")
        e_path.text = self.path
        e_size = xee.SubElement(root, "size")
        e_size.text = str(self.size)
        return root

    def to_xml(self):
        if self.processed:
            return self.to_xml_document()
        else:
            return self.to_xml_file()

    def __str__(self):
        return "\nName: {}, Author: {}\nDate_c: {} Date_m: {}\nPages: {} Size: {}\nPath{}".format(
            self.name, self.author, self.date_create, self.date_modified, self.pages, self.size, self.path
        )

def document_info_sort_date_create(e):
    return e.date_create

def document_info_sort_date_modified(e):
    return e.date_create
class Timeline():
    def __init__(self):
        self.processed = []
        self.unprocessed = []

    def add(self, doc):
        if doc.processed:
            self.processed.append(doc)
        else:
            self.unprocessed.append(doc)

    def total(self):
        return len(self.processed) + len(self.unprocessed)

    def sort(self, key=document_info_sort_date_create):
        self.processed.sort(key=key)
        self.unprocessed.sort(key=key)

class MagicProcessor():
    def __init__(self):
        self._data = None

    def process(self, filename):
        doc_info = DocumentInfo(filename)
        file_magic = magic.from_file(filename)
        
        doc_info.author = fetch_or_fail('Author:', file_magic)
        doc_info.author_last = fetch_or_fail('Last Saved By:', file_magic)

        mstr = fetch_or_fail('Create Time/Date:', file_magic)
        doc_info.date_create = set_date_or_fail(mstr)

        mstr = fetch_or_fail('Last Saved Time/Date:', file_magic)
        doc_info.date_modified = set_date_or_fail(mstr)
        
        if doc_info.date_create == None:
            doc_info.set_date_create_from_file()
        if doc_info.date_modified == None:
            doc_info.set_date_modified_from_file()

        doc_info.pages = fetch_or_fail('Number of Pages:', file_magic)
        doc_info.processed = True
        return doc_info

class XlsProcessor():
    def __init__(self):
        self._data = None

    def process(self, filename):
        doc_info = DocumentInfo(filename)
        return doc_info

class DocxProcessor():
    def __init__(self):
        self._data = None

    def process(self, filename):
        doc_info = DocumentInfo(filename)
        file = open(filename, 'rb')
        document = Document(file)
        core_props = document.core_properties;
        doc_info.author = core_props.author
        doc_info.author_last = core_props.last_modified_by
        doc_info.date_create = core_props.created
        doc_info.date_modified = core_props.modified
        if doc_info.date_create == None:
            doc_info.set_date_create_from_file()
        if doc_info.date_modified == None:
            doc_info.set_date_modified_from_file()
        file.close()
        doc_info.processed = True
        return doc_info

class PptxProcessor():
    def __init__(self):
        self._data = None

    def process(self, filename):
        doc_info = DocumentInfo(filename)
        file = open(filename, 'rb')
        document = Presentation(file)
        core_props = document.core_properties;
        doc_info.author = core_props.author
        doc_info.author_last = core_props.last_modified_by
        doc_info.date_create = core_props.created
        doc_info.date_modified = core_props.modified
        if doc_info.date_create == None:
            doc_info.set_date_create_from_file()
        if doc_info.date_modified == None:
            doc_info.set_date_modified_from_file()
        file.close()
        doc_info.processed = True
        return doc_info

class PdfProcessor():
    def __init__(self):
        self._data = None

    def process(self, filename):
        # print("PDF: {}".format(filename))
        doc_info = DocumentInfo(filename)
        doc_info.author_last = None
        with open(filename, 'rb') as f:
            pdf = PdfFileReader(f)
            info = pdf.documentInfo
            xmp = pdf.getXmpMetadata()
            doc_info.pages = pdf.getNumPages()
            if info:
                doc_info.author = info.author
            if xmp:
                doc_info.date_create = xmp.xmp_createDate
                doc_info.date_modified = xmp.xmp_modifyDate
        
        if doc_info.date_create == None:
            doc_info.set_date_create_from_file()
        if doc_info.date_modified == None:
            doc_info.set_date_modified_from_file()

        doc_info.processed = True
        return doc_info

class DefaultProcessor():
    def __init__(self):
        self._data = None
    def process(self, filename):
        doc_info = DocumentInfo(filename)
        doc_info.author = None
        doc_info.author_last = None
        doc_info.set_date_create_from_file()
        doc_info.set_date_modified_from_file()
        doc_info.processed = False;
        return doc_info
class DocumentProcessorFactory():
    def __init__(self):
        self._processors = {}

    def register_mime(self, mime, processor):
        self._processors[mime] = processor

    def get_processor(self, mime):
        processor = self._processors.get(mime)
        if not processor:
            processor = DefaultProcessor
        return processor()

processor_factory = DocumentProcessorFactory()
processor_factory.register_mime('application/msword', MagicProcessor)
processor_factory.register_mime('application/vnd.ms-excel', MagicProcessor)
processor_factory.register_mime('application/vnd.ms-powerpoint', MagicProcessor)
processor_factory.register_mime('text/rtf', MagicProcessor)

processor_factory.register_mime('application/vnd.openxmlformats-officedocument.wordprocessingml.document', DocxProcessor)
processor_factory.register_mime('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', MagicProcessor)
processor_factory.register_mime('application/vnd.openxmlformats-officedocument.presentationml.presentation', PptxProcessor)

processor_factory.register_mime('application/pdf', PdfProcessor)

class Crawler():

    def __init__(self):
        self.supported = ['application/msword', 'docx', 'xls', 'xlx', 'ppt', 'pptx', 'pdf']

    def discover(self, target_path="/tmp"):
        filelist = []
        for dirName, subdirList, fileList in os.walk(target_path):
            for fname in fileList:
                filename = os.path.join(dirName, fname)
                filelist.append(filename)
        return filelist

    def create_document_info_from_file(self, filename):
        file_magic = magic.from_file(filename, mime=True)
        document_info = None
        try:
            processor = processor_factory.get_processor(file_magic)
        except ValueError:
            print("File: [{}] is not supported.".format(filename))            
        else:
            document_info = processor.process(filename)
        return document_info

    def collect_timeline(self, target_path="/tmp")-> Timeline:
        timeline = Timeline()
        files = self.discover(target_path)
        for file in files:
            file_docu_info = self.create_document_info_from_file(file)
            timeline.add(file_docu_info)
        print("Documents discovered: [{}]".format(timeline.total()))
        timeline.sort(key=document_info_sort_date_create)
        return timeline

    def write_html_processed_document(self, document):
        div = dominate.tags.div(_class='document')
        dominate.tags.div(
            dominate.tags.a(document.name, href='%s' % document.path),
            _class='header'
        )
        with dominate.tags.div(_class='content'):
            with dominate.tags.ul():
                dominate.tags.li('Author: %s' % document.author)
                dominate.tags.li('Create date: %s' % document.date_create)
                dominate.tags.li('Last editor: %s' % document.author_last)
                dominate.tags.li('Modified date: %s' % document.date_modified)
                dominate.tags.li('Pages: %s' % document.pages)
                dominate.tags.li('Size: %d' % document.size)
        return div;

    def write_html_processed(self, documents):
        div = dominate.tags.div(_class='processed')
        for doc in documents:
            self.write_html_processed_document(doc)
        return div 

    def write_html_unprocessed_file(self, file):
        li = dominate.tags.li(
            dominate.tags.a(file.name, href='%s' % file.path)
        )
        return li;

    def write_html_unprocessed(self, documents):
        div = dominate.tags.div(_class='unprocessed')
        with dominate.tags.ul():
            for doc in documents:
                self.write_html_unprocessed_file(doc)
        return div 

    def write_timeline_html(self, path, filename, timeline):
        print(
            "Writing [{}] documents HTML timeline.\n\tFilename: [{}]\n\tPath: [{}]"
            .format(timeline.total(), filename, path)
        )
        html_document = dominate.document(path)
        with html_document.head:
            dominate.tags.link(rel='stylesheet', href='style.css')
            dominate.tags.script(type='text/javascript', src='script.js')
            dominate.tags.h1(path)
            self.write_html_processed(timeline.processed)
            self.write_html_unprocessed(timeline.unprocessed)
        with open(filename, 'w') as f:
            f.write(html_document.render())

    def write_timeline_xml(self, path, filename, timeline):
        print(
            "Writing [{}] documents XML timeline.\n\tFilename: [{}]\n\tPath: [{}]"
            .format(timeline.total(), filename, path)
        )
        root = xee.Element("path")
        e_path = xee.SubElement(root, "path")
        e_path.text = path

        processed = xee.Element("processed")
        root.append(processed)
        for doc in timeline.processed:
            doc_xml = doc.to_xml()
            processed.append(doc_xml)

        unprocessed = xee.Element("unprocessed")
        root.append(unprocessed)
        for file in timeline.unprocessed:
            file_xml = file.to_xml()
            unprocessed.append(file_xml)

        tree = xee.ElementTree(root)
        with open(filename, 'wb') as f:
            tree.write(f)

    def write_xls_headers(self, sheet, headers):
        style_header = xlwt.easyxf('font: bold 1') 
        for column, header in enumerate(headers):
            sheet.write(1, column, header, style_header)

    def write_xls_processed_header(self, sheet, path):
        style_path = xlwt.easyxf('font: bold 1, color blue;') 
        sheet.write(0, 0, path, style_path)
        headers = ['#', 'name', 'path', 'date_create', 'author', 'date_modified', 'author_last', 'pages', 'size']
        self.write_xls_headers(sheet, headers)

    def write_xls_unprocessed_header(self, sheet, path):
        style_path = xlwt.easyxf('font: bold 1, color red;') 
        sheet.write(0, 0, path, style_path)
        headers = ['#', 'name', 'path', 'date_create', 'size']
        self.write_xls_headers(sheet, headers)

    def write_xls_document(self, sheet, cursor, document):
        sheet.write(cursor, 0, '')
        sheet.write(cursor, 1, document.name)
        sheet.write(cursor, 2, 'file:/{}'.format(document.path))
        sheet.write(cursor, 3, document.date_create.strftime("%m/%d/%Y, %H:%M:%S"))
        sheet.write(cursor, 4, document.author)
        sheet.write(cursor, 5, document.date_modified.strftime("%m/%d/%Y, %H:%M:%S"))
        sheet.write(cursor, 6, document.author_last)
        sheet.write(cursor, 7, document.pages)
        sheet.write(cursor, 8, document.size)

    def write_xls_processed_documents(self, sheet, documents):
        row_start = 2
        for row, document in enumerate(documents):
            self.write_xls_document(sheet, row_start + row, document)

    def write_xls_file(self, sheet, cursor, file):
        sheet.write(cursor, 0, cursor)
        sheet.write(cursor, 1, file.name)
        sheet.write(cursor, 2, 'file:/{})'.format(file.path))
        sheet.write(cursor, 3, file.date_create.strftime("%m/%d/%Y, %H:%M:%S"))
        sheet.write(cursor, 4, file.size)

    def write_xls_unprocessed_files(self, sheet, files):
        for row, file in enumerate(files):
            self.write_xls_file(sheet, row + 2, file)

    def write_timeline_xls(self, path, filename, timeline):
        print(
            "Writing [{}] documents XLS timeline.\n\tFilename: [{}]\n\tPath: [{}]"
            .format(timeline.total(), filename, path)
        )
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('processed')
        self.write_xls_processed_header(sheet, path)
        self.write_xls_processed_documents(sheet, timeline.processed)
        
        sheet = workbook.add_sheet('unprocessed')
        self.write_xls_unprocessed_header(sheet, path)
        self.write_xls_unprocessed_files(sheet, timeline.unprocessed)

        workbook.save(filename)

if __name__ == "__main__":
    filename_xml = ''
    filename_html = ''

    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    parser.add_argument("filename")
    args = parser.parse_args()
 
    # get the arguments value
    if args.path == None or not os.path.isdir(args.path):
        print("Invalid target path: {}".format(args.path))
    
    if args.filename == None:
        print("Invalid filename: {}".format(args.filename))
    
    filename = args.filename
    filename_xml = os.path.join(os.getcwd(), filename + '.xml')
    filename_html = os.path.join(os.getcwd(), filename + '.html')
    filename_xls = os.path.join(os.getcwd(), filename + '.xls')
    if os.path.isfile(filename_xls) or os.path.exists(filename_xml) or os.path.exists(filename_xls):
        print(
            "Files: {} or {} or {} already exist."
            .format(filename_xml, filename_html, filename_xls)
        )

    print('Target path: {}'.format(args.path))
    print(
        'HTML: {}\nXML: {}\nXLS: {}'
        .format(filename_html, filename_xml, filename_xls)
    )
    crawler = Crawler()    
    timeline = crawler.collect_timeline(args.path)
    crawler.write_timeline_html(args.path, filename_html, timeline)
    crawler.write_timeline_xml(args.path, filename_xml, timeline)
    crawler.write_timeline_xls(args.path, filename_xls, timeline)    
