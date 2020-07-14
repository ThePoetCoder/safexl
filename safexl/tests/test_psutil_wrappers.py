# Copyright (c) 2020 safexl
import unittest
import tempfile
import pythoncom
import pywintypes
import win32com.client
import safexl


class test_is_excel_open(unittest.TestCase):
    def test_when_excel_is_open(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")
        self.assertTrue(safexl.is_excel_open())
        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_when_excel_is_not_open(self):
        safexl.kill_all_instances_of_excel()
        self.assertFalse(safexl.is_excel_open())


class test_excel_open_files(unittest.TestCase):
    @staticmethod
    def count_tempfiles():
        return sum(1 for file in safexl.toolkit.excel_open_files() if file.endswith(".tmp"))

    def test_that_results_update_immediately(self):
        safexl.kill_all_instances_of_excel()
        self.assertFalse(safexl.is_excel_open())

        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        # no workbooks have been added yet
        open_files_1 = self.count_tempfiles()
        self.assertEqual(0, open_files_1)

        # 1 workbook added, 1 more .tmp file than before
        wb1 = application.Workbooks.Add()
        open_files_2 = self.count_tempfiles()
        self.assertEqual(1, open_files_2)
        self.assertEqual(open_files_1 + 1, open_files_2)

        # 2 workbooks added, 1 more .tmp file than before
        wb2 = application.Workbooks.Add()
        open_files_3 = self.count_tempfiles()
        self.assertEqual(2, open_files_3)
        self.assertEqual(open_files_2 + 1, open_files_3)

        # 1 workbook removed, 1 less .tmp file than before
        application.DisplayAlerts = False
        wb2.Close(SaveChanges=False)
        application.DisplayAlerts = True
        open_files_4 = self.count_tempfiles()
        self.assertEqual(1, open_files_4)
        self.assertEqual(open_files_3 - 1, open_files_4)

        # 2 workbooks removed, 1 less .tmp file than before, back to beginning
        application.DisplayAlerts = False
        wb1.Close(SaveChanges=False)
        application.DisplayAlerts = True
        open_files_5 = self.count_tempfiles()
        self.assertEqual(0, open_files_5)
        self.assertEqual(open_files_4 - 1, open_files_5)
        self.assertEqual(open_files_5, open_files_1)

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_that_newly_saved_file_is_picked_up_and_lost_when_closed(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb1 = application.Workbooks.Add()
        with tempfile.TemporaryDirectory() as temp_dir:
            save_filepath = f"{temp_dir}\\temporary.xlsx"
            application.DisplayAlerts = False
            wb1.SaveAs(save_filepath)

            self.assertIn(save_filepath, safexl.toolkit.excel_open_files())
            wb1.Close()
            self.assertNotIn(save_filepath, safexl.toolkit.excel_open_files())

            application.DisplayAlerts = True

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()


class test_kill_all_instances_of_excel(unittest.TestCase):
    def test_can_kill_excel_instance(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        self.assertEqual("Microsoft Excel", application.Name)
        safexl.kill_all_instances_of_excel()
        with self.assertRaises(pywintypes.com_error):
            app_name = application.Name

        del application
        pythoncom.CoUninitialize()

    def test_can_kill_excel_instance_when_passed_app(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        self.assertEqual("Microsoft Excel", application.Name)
        safexl.kill_all_instances_of_excel(application)
        with self.assertRaises(pywintypes.com_error):
            app_name = application.Name

        del application
        pythoncom.CoUninitialize()

    def test_can_kill_multiple_excel_instances(self):
        pythoncom.CoInitialize()

        # Using DispatchEx specifically here to create multiple instances of Excel open at once
        application1 = win32com.client.DispatchEx("Excel.Application")
        application2 = win32com.client.DispatchEx("Excel.Application")
        application3 = win32com.client.DispatchEx("Excel.Application")

        self.assertEqual("Microsoft Excel", application1.Name)
        self.assertEqual("Microsoft Excel", application2.Name)
        self.assertEqual("Microsoft Excel", application3.Name)

        safexl.kill_all_instances_of_excel()

        with self.assertRaises(pywintypes.com_error):
            app_name = application1.Name
        with self.assertRaises(pywintypes.com_error):
            app_name = application2.Name
        with self.assertRaises(pywintypes.com_error):
            app_name = application3.Name

        del application1
        del application2
        del application3
        pythoncom.CoUninitialize()
