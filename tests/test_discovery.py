#!/usr/bin/python3

import unittest 
import tempfile
import os
import warnings
from distutils.dir_util import copy_tree

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
        self.assertNotEqual(document_info.pages, None)
        self.assertNotEqual(document_info.size, None)

    def test_crawler_can_discover_files_in_target_path(self):
        app = herostratus.Crawler()
        files = app.discover(self.test_dir.name)
        self.assertTrue(files != None)
        self.assertEqual(len(files), self.file_count)

    def test_crawler_can_collect_file_information(self):
        app = herostratus.Crawler()
        timeline = app.collect_timeline(self.test_dir.name)
        self.assertEqual(len(timeline), self.file_count)

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

    # XLX
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
    # def test_crawler_can_create_HTML_timeline(self):
    #     filename = os.path.join(os.getcwd(), 'output.html')
    #     print("Filename: {}".format(filename))
    #     app = herostratus.Crawler()
    #     timeline = app.collect_timeline(self.test_dir.name)
    #     self.assertEqual(len(timeline), self.file_count)
    #     herostratus.writeHTML(timeline, filename)
    #     self.assertTrue(os.path.isfile(filename))

if __name__ == '__main__':
    unittest.main()
