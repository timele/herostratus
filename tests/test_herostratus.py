#!/usr/bin/python3

import unittest 
import tempfile
import os
import warnings
from distutils.dir_util import copy_tree
import xlrd

from herostratus import herostratus

warnings.filterwarnings("ignore") 

class Test_files_discovery(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.TemporaryDirectory()
        test_data_dir = os.path.join(os.getcwd(), "tests/data")
        copy_tree(test_data_dir, self.test_dir.name)
        self.file_count = len(os.listdir(self.test_dir.name))

    def tearDown(self):
        self.test_dir.cleanup()

    def assert_document_has_data(self, document_info):
        self.assertNotEqual(document_info.name, None)
        self.assertNotEqual(document_info.size, None)

    def test_document_info_can_convert_to_xml(self):
        app = herostratus.Crawler()
        filenames = [
            'file_example_XLSX_1.xlsx', 'file_example_XLSX_50.xlsx',
            'file_example_XLSX_100.xlsx', 'file_example_XLSX_1000.xlsx',
            'file_example_XLSX_5000.xlsx'
        ]
        for file in filenames:
            document_info = app.create_document_info_from_file(os.path.join(self.test_dir.name, file))
            xml = document_info.to_xml()
            self.assertIsNotNone(xml)
            self.assertNotEqual(len(xml), 0)

    def test_crawler_can_discover_files_in_target_path(self):
        app = herostratus.Crawler()
        files = app.discover(self.test_dir.name)
        self.assertTrue(files != None)
        self.assertEqual(len(files), self.file_count)

    def test_crawler_can_collect_file_information(self):
        app = herostratus.Crawler()
        timeline = app.collect_timeline(self.test_dir.name)
        self.assertEqual(timeline.total(), self.file_count)

    # DOC
    def test_crawler_can_get_information_from_DOC_file(self):
        app = herostratus.Crawler()
        filenames = [
            'file_example_DOC_1.doc', 'file_example_DOC_100kB.doc', 
            'file_example_DOC_500kB.doc'
        ]
        for file in filenames:
            file_path = os.path.join(self.test_dir.name, file)
            file_docu_info = app.create_document_info_from_file(file_path)
            self.assert_document_has_data(file_docu_info)

    # DOCX
    def test_crawler_can_get_information_from_DOCX_file(self):
        app = herostratus.Crawler()
        filenames = [
         'file_example_DOCX_100kB.docx', 'file_example_DOCX_1.docx', 
         'file_example_DOCX_3.docx', 'file_example_DOCX_500kB.docx'
        ]
        for file in filenames:
            file_path = os.path.join(self.test_dir.name, file)
            file_docu_info = app.create_document_info_from_file(file_path)
            self.assert_document_has_data(file_docu_info)

    # XLS
    def test_crawler_can_get_information_from_XLS_file(self):
        app = herostratus.Crawler()
        filenames = [
            'file_example_XLS_1.xls', 'file_example_XLS_100.xls', 
            'file_example_XLS_1000.xls', 'file_example_XLS_5000.xls'
        ]
        for file in filenames:
            file_path = os.path.join(self.test_dir.name, file)
            file_docu_info = app.create_document_info_from_file(file_path)
            self.assert_document_has_data(file_docu_info)

    # XLSX
    def test_crawler_can_get_information_from_XLX_file(self):
        app = herostratus.Crawler()
        filenames = [
            'file_example_XLSX_1.xlsx', 'file_example_XLSX_50.xlsx',
            'file_example_XLSX_100.xlsx', 'file_example_XLSX_1000.xlsx',
            'file_example_XLSX_5000.xlsx'
        ]
        for file in filenames:
            file_path = os.path.join(self.test_dir.name, file)
            file_docu_info = app.create_document_info_from_file(file_path)
            self.assert_document_has_data(file_docu_info)

    # PPT
    def test_crawler_can_get_information_from_PPT_file(self):
        app = herostratus.Crawler()
        filenames = [
            'file_example_PPT_1.ppt', 'file_example_PPT_1MB.ppt',
            'file_example_PPT_250kB.ppt', 'file_example_PPT_500kB.ppt'
        ]
        for file in filenames:
            file_path = os.path.join(self.test_dir.name, file)
            file_docu_info = app.create_document_info_from_file(file_path)
            self.assert_document_has_data(file_docu_info)

    # PPTX
    def test_crawler_can_get_information_from_PPTX_file(self):
        app = herostratus.Crawler()
        filenames = [
            'file_example_PPTX_1.pptx', 'file_example_PPTX_2.pptx'
        ]
        for file in filenames:
            file_path = os.path.join(self.test_dir.name, file)
            file_docu_info = app.create_document_info_from_file(file_path)
            self.assert_document_has_data(file_docu_info)

    # PDF
    def test_crawler_can_get_information_from_PDF_file(self):
        app = herostratus.Crawler()
        filenames = [
            'file_example_PDF_1.pdf', 'file_example_PDF_2.pdf', 
            'file_example_PDF_3.pdf', 'file_example_PDF_500_kB.pdf'
        ]
        for file in filenames:
            file_path = os.path.join(self.test_dir.name, file)
            file_docu_info = app.create_document_info_from_file(file_path)
            self.assert_document_has_data(file_docu_info)

    # HTML Timeline
    def test_crawler_can_create_HTML_timeline(self):
        filename = os.path.join(os.getcwd(), '/tmp/output.html')
        print("Filename: {}".format(filename))
        app = herostratus.Crawler()
        timeline = app.collect_timeline(self.test_dir.name)
        self.assertEqual(timeline.total(), self.file_count)
        app.write_timeline_html(self.test_dir.name, filename, timeline)
        self.assertTrue(os.path.isfile(filename))
    
    # XML Timeline
    def test_crawler_can_create_XML_timeline(self):
        filename = os.path.join(os.getcwd(), '/tmp/output.xml')
        print("Filename: {}".format(filename))
        app = herostratus.Crawler()
        timeline = app.collect_timeline(self.test_dir.name)
        self.assertEqual(timeline.total(), self.file_count)
        app.write_timeline_xml(self.test_dir.name, filename, timeline)
        self.assertTrue(os.path.isfile(filename))


    def assert_xls_processed_header(self, sheet):
        self.assertEqual(sheet.name, 'processed')
        self.assertEqual(sheet.cell_value(0,0), self.test_dir.name)
        self.assertEqual(sheet.cell_value(1,0), '#')
        self.assertEqual(sheet.cell_value(1,1), 'name')
        self.assertEqual(sheet.cell_value(1,2), 'path')
        self.assertEqual(sheet.cell_value(1,3), 'date_create')
        self.assertEqual(sheet.cell_value(1,4), 'author')
        self.assertEqual(sheet.cell_value(1,5), 'date_modified')
        self.assertEqual(sheet.cell_value(1,6), 'author_last')
        self.assertEqual(sheet.cell_value(1,7), 'pages')
        self.assertEqual(sheet.cell_value(1,8), 'size')

    def assert_xls_unprocessed_header(self, sheet):
        self.assertEqual(sheet.name, 'unprocessed')
        self.assertEqual(sheet.cell_value(0,0), self.test_dir.name)
        self.assertEqual(sheet.cell_value(1,0), '#')
        self.assertEqual(sheet.cell_value(1,1), 'name')
        self.assertEqual(sheet.cell_value(1,2), 'path')
        self.assertEqual(sheet.cell_value(1,3), 'date_create')
        self.assertEqual(sheet.cell_value(1,4), 'size')

    # XLS Timeline
    def test_crawler_can_create_XLS_timeline(self):
        filename = os.path.join(os.getcwd(), '/tmp/output.xls')
        print("Filename: {}".format(filename))
        app = herostratus.Crawler()
        timeline = app.collect_timeline(self.test_dir.name)
        self.assertEqual(timeline.total(), self.file_count)
        app.write_timeline_xls(self.test_dir.name, filename, timeline)
        self.assertTrue(os.path.isfile(filename))
        book = xlrd.open_workbook(filename)
        self.assertEqual(book.nsheets, 2)
        sheet = book.sheet_by_index(0)
        self.assertIsNotNone(sheet)
        self.assert_xls_processed_header(sheet)
        self.assertIsNotNone(sheet.cell_value(2,0))
        self.assertIsNotNone(sheet.cell_value(2,1))

        sheet = book.sheet_by_index(1)
        self.assertIsNotNone(sheet)
        self.assert_xls_unprocessed_header(sheet)
        self.assertIsNotNone(sheet.cell_value(2,0))
        self.assertIsNotNone(sheet.cell_value(2,1))

if __name__ == '__main__':
    unittest.main()
