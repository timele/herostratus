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
        self.pages = 0

    def set_date_create_from_file(self):
        fname = pathlib.Path(self.path)
        self.date_create = dt.datetime.fromtimestamp(fname.stat().st_ctime)

    def set_date_modified_from_file(self):
        fname = pathlib.Path(self.path)
        self.date_modified = dt.datetime.fromtimestamp(fname.stat().st_mtime)

    def __str__(self):
        return "\nName: {}, Author: {}\nDate_c: {} Date_m: {}\nPages: {} Size: {}\nPath{}".format(
            self.name, self.author, self.date_create, self.date_modified, self.pages, self.size, self.path
        )

def document_info_sort_date_create(e):
    return e.date_create

def document_info_sort_date_modified(e):
    return e.date_create

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
        document = Document(filename)
        core_props = document.core_properties;
        doc_info.author = core_props.author
        doc_info.author_last = core_props.last_modified_by
        doc_info.date_create = core_props.created
        doc_info.date_modified = core_props.modified

        if doc_info.date_create == None:
            doc_info.set_date_create_from_file()
        if doc_info.date_modified == None:
            doc_info.set_date_modified_from_file()
            
        doc_info.pages = 0
        return doc_info

class PptxProcessor():
    def __init__(self):
        self._data = None

    def process(self, filename):
        doc_info = DocumentInfo(filename)
        document = Presentation(filename)
        core_props = document.core_properties;
        doc_info.author = core_props.author
        doc_info.author_last = core_props.last_modified_by
        doc_info.date_create = core_props.created
        doc_info.date_modified = core_props.modified

        if doc_info.date_create == None:
            doc_info.set_date_create_from_file()
        if doc_info.date_modified == None:
            doc_info.set_date_modified_from_file()

        doc_info.pages = 0
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

        return doc_info

class DocumentProcessorFactory():
    def __init__(self):
        self._processors = {}

    def register_mime(self, mime, processor):
        self._processors[mime] = processor

    def get_processor(self, mime):
        processor = self._processors.get(mime)
        if not processor:
            raise ValueError(mime)
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
            if len(subdirList) > 0:
                del subdirList[0]
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

    def collect_timeline(self, target_path="/tmp"):
        timeline = []
        files = self.discover(target_path)
        for file in files:
            file_docu_info = self.create_document_info_from_file(file)
            if file_docu_info != None:
                timeline.append(file_docu_info)
        print("Files discovered: [{}]".format(len(timeline)))
        timeline.sort(key=document_info_sort_date_create)
        return timeline

    def rebuild_timeline_by_date_modified(self, timeline): 
        return sorted(timeline, key=document_info_sort_date_modified)

    def write_timeline(self, timeline, target_filename):
        pass

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
 
    parser.add_argument("path")
    parser.add_argument("filename")
 
    # parse the arguments
    args = parser.parse_args()
 
    # get the arguments value
    if args.path == None or not os.path.isdir(args.path):
        print("Invalid target path: {}".format(args.path))

    if args.filename == None or os.path.isfile(args.filename):
        print("Invalid filename: {} or it exists.".format(args.filename))
    
    print("Target path: {}".format(args.path))
    print("Writing to: {}".format(args.filename))
    crawler = Crawler()
    timeline = crawler.collect_timeline(args.path)
    print("10 first elements")
    for i in range(10):
        print("{}".format(timeline[i]))

    crawler.write_timeline(args.filename, timeline)
        
