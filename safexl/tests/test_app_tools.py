# Copyright (c) 2020 safexl
import unittest
import tempfile
import pythoncom
import pywintypes
import win32com.client
import safexl


class test_close_workbooks(unittest.TestCase):
    def test_closing_one_new_workbook(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")
        wb = application.Workbooks.Add()

        self.assertEqual("Book1", wb.Name)
        safexl.close_workbooks(application, [wb])
        with self.assertRaises(pywintypes.com_error):
            wb_name = wb.Name

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_with_no_new_workbooks(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")
        wb = application.Workbooks.Add()

        self.assertEqual("Book1", wb.Name)
        safexl.close_workbooks(application, set())
        self.assertEqual("Book1", wb.Name)

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_with_many_new_workbooks(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        # Open 50 workbooks at one time
        wbs = []
        for i in range(50):
            wb = application.Workbooks.Add()
            wbs.append(wb)
        self.assertEqual(50, application.Workbooks.Count)
        safexl.close_workbooks(application, wbs)
        self.assertEqual(0, application.Workbooks.Count)

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()


class test_new_workbooks(unittest.TestCase):
    def test_new_workbook_is_returned_and_old_workbook_is_not(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb1 = application.Workbooks.Add()
        original_workbook_list = safexl.workbooks_currently_open(application)
        self.assertIn(wb1, original_workbook_list)

        wb2 = application.Workbooks.Add()
        new_workbook_list = safexl.toolkit.new_workbooks(application, original_workbook_list)
        self.assertNotIn(wb1, new_workbook_list)
        self.assertIn(wb2, new_workbook_list)
        self.assertEqual([wb2], new_workbook_list)

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_only_currently_open_workbooks_can_be_found(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb1 = application.Workbooks.Add()
        original_workbook_list = safexl.workbooks_currently_open(application)
        self.assertIn(wb1, original_workbook_list)
        wb2 = application.Workbooks.Add()
        wb2.Close()
        new_workbook_list = safexl.toolkit.new_workbooks(application, original_workbook_list)
        self.assertNotIn(wb1, new_workbook_list)
        self.assertNotIn(wb2, new_workbook_list)
        self.assertEqual([], new_workbook_list)

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_ref_to_workbook_remains_even_if_it_is_closed_after_being_added_to_list(self):
        """
        pywin32 oddity, workbook COM object exists after closing,
        but cannot be used for anything without causing an error
        """
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb1 = application.Workbooks.Add()
        original_workbook_list = safexl.workbooks_currently_open(application)
        self.assertIn(wb1, original_workbook_list)
        wb2 = application.Workbooks.Add()
        new_workbook_list = safexl.toolkit.new_workbooks(application, original_workbook_list)
        wb2.Close()
        with self.assertRaises(pywintypes.com_error):
            wb_name = wb2.Name
        self.assertNotIn(wb1, new_workbook_list)
        self.assertIn(wb2, new_workbook_list)
        self.assertEqual([wb2], new_workbook_list)

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()


class test_see_excel(unittest.TestCase):
    def test_ability_to_make_multiple_windows_visible(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb = application.Workbooks.Add()
        # Open 5 windows at one time
        original_window = wb.Windows(1)
        original_window.Visible = True
        self.assertEqual(wb.Name, original_window.Caption)
        for i in range(4):
            new_window = original_window.NewWindow()
            new_window.WindowState = safexl.xl_constants.xlMinimized
            new_window.Visible = False

        self.assertEqual(5, wb.Windows.Count)
        self.assertEqual(1, sum(1 for window in wb.Windows if window.Visible))
        self.assertEqual(4, sum(1 for window in wb.Windows if not window.Visible))
        safexl.see_excel([wb], safexl.xl_constants.xlMinimized)
        self.assertEqual(5, wb.Windows.Count)
        self.assertEqual(5, sum(1 for window in wb.Windows if window.Visible))
        self.assertEqual(0, sum(1 for window in wb.Windows if not window.Visible))

        # cleanup
        application.DisplayAlerts = False
        for window in wb.Windows:
            window.Close()
        application.DisplayAlerts = True
        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_error_occurs_when_a_workbook_has_no_windows_at_all(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb = application.Workbooks.Add()
        wb.Windows(1).Close()

        with self.assertRaises(pywintypes.com_error):
            window_caption = wb.Windows(1).Caption

        with self.assertRaises(pywintypes.com_error):
            # Once the final window of a workbook has closed, the workbook itself closes...
            wb_name = wb.Name

        with self.assertRaises(pywintypes.com_error):
            # So there is nothing to 'see' here anymore
            safexl.see_excel([wb], safexl.xl_constants.xlMinimized)

        application.DisplayAlerts = True
        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_PERSONAL_workbook_is_not_affected(self):
        with safexl.application(kill_after=True) as app:
            personal_wb_paths = [wb for wb in app.Workbooks if "PERSONAL.XLSB" in wb.FullName]
            if personal_wb_paths:
                personal_wb = personal_wb_paths[0]
                self.assertFalse(personal_wb.Windows(1).Visible)
                safexl.see_excel([personal_wb], safexl.xl_constants.xlMinimized)
                self.assertFalse(personal_wb.Windows(1).Visible)


class test_workbooks_currently_open(unittest.TestCase):
    @staticmethod
    def count_tempfiles(app):
        return sum(1 for file in safexl.workbooks_currently_open(app) if file.FullName.startswith("Book"))

    def test_that_results_update_immediately(self):
        safexl.kill_all_instances_of_excel()
        self.assertFalse(safexl.is_excel_open())

        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        # no workbooks have been added yet
        openfiles_1 = self.count_tempfiles(application)
        self.assertEqual(0, openfiles_1)

        # 1 workbook added, 1 more .tmp file than before
        wb1 = application.Workbooks.Add()
        openfiles_2 = self.count_tempfiles(application)
        self.assertEqual(1, openfiles_2)
        self.assertEqual(openfiles_1 + 1, openfiles_2)

        # 2 workbooks added, 1 more .tmp file than before
        wb2 = application.Workbooks.Add()
        openfiles_3 = self.count_tempfiles(application)
        self.assertEqual(2, openfiles_3)
        self.assertEqual(openfiles_2 + 1, openfiles_3)

        # 1 workbook removed, 1 less .tmp file than before
        application.DisplayAlerts = False
        wb2.Close(SaveChanges=False)
        application.DisplayAlerts = True
        openfiles_4 = self.count_tempfiles(application)
        self.assertEqual(1, openfiles_4)
        self.assertEqual(openfiles_3 - 1, openfiles_4)

        # 2 workbooks removed, 1 less .tmp file than before, back to beginning
        application.DisplayAlerts = False
        wb1.Close(SaveChanges=False)
        application.DisplayAlerts = True
        openfiles_5 = self.count_tempfiles(application)
        self.assertEqual(0, openfiles_5)
        self.assertEqual(openfiles_4 - 1, openfiles_5)
        self.assertEqual(openfiles_5, openfiles_1)

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_that_newly_saved_file_is_noticed_and_then_lost_when_it_is_closed(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb1 = application.Workbooks.Add()
        with tempfile.TemporaryDirectory() as temp_dir:
            save_filepath = f"{temp_dir}\\temporary.xlsx"
            application.DisplayAlerts = False
            wb1.SaveAs(save_filepath)

            self.assertIn(save_filepath, [wb.FullName for wb in safexl.workbooks_currently_open(application)])
            wb1.Close()
            self.assertNotIn(save_filepath, [wb.FullName for wb in safexl.workbooks_currently_open(application)])

            application.DisplayAlerts = True

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()


class test_last_row(unittest.TestCase):
    def test_row_count_increases_and_decreases(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb = application.Workbooks.Add()
        ws = wb.ActiveSheet
        ws.Range("A1").Value = 1
        self.assertEqual(1, safexl.last_row(ws))
        ws.Range("A2").Value = 1
        self.assertEqual(2, safexl.last_row(ws))
        ws.Range("A2").Value = ""
        self.assertEqual(1, safexl.last_row(ws))

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_if_highest_row_number_is_not_in_column_a(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb = application.Workbooks.Add()
        ws = wb.ActiveSheet
        ws.Range("A1").Value = 1
        self.assertEqual(1, safexl.last_row(ws))
        ws.Range("B2").Value = 1
        self.assertEqual(2, safexl.last_row(ws))

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_stays_focused_on_currentregion(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb = application.Workbooks.Add()
        ws = wb.ActiveSheet
        ws.Range("A1").Value = 1
        self.assertEqual(1, safexl.last_row(ws))
        ws.Range("A2").Value = ""
        ws.Range("A3").Value = 1
        self.assertEqual(1, safexl.last_row(ws))

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()


class test_last_column(unittest.TestCase):
    def test_column_count_increases_and_decreases(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb = application.Workbooks.Add()
        ws = wb.ActiveSheet
        ws.Range("A1").Value = 1
        self.assertEqual(1, safexl.last_column(ws))
        ws.Range("B1").Value = 1
        self.assertEqual(2, safexl.last_column(ws))
        ws.Range("B1").Value = ""
        self.assertEqual(1, safexl.last_column(ws))

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_if_highest_column_number_is_not_in_row_1(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb = application.Workbooks.Add()
        ws = wb.ActiveSheet
        ws.Range("A1").Value = 1
        self.assertEqual(1, safexl.last_column(ws))
        ws.Range("B2").Value = 1
        self.assertEqual(2, safexl.last_column(ws))

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()

    def test_stays_focused_on_currentregion(self):
        pythoncom.CoInitialize()
        application = win32com.client.Dispatch("Excel.Application")

        wb = application.Workbooks.Add()
        ws = wb.ActiveSheet
        ws.Range("A1").Value = 1
        self.assertEqual(1, safexl.last_column(ws))
        ws.Range("B1").Value = ""
        ws.Range("C1").Value = 1
        self.assertEqual(1, safexl.last_column(ws))

        safexl.kill_all_instances_of_excel(application)
        del application
        pythoncom.CoUninitialize()


class test_worksheet_name_sanitization(unittest.TestCase):
    def test_len_limit_is_31(self):
        input_name  = "0123456789012345678901234567890123456789"
        self.assertEqual(40, len(input_name))
        expectation = "0123456789012345678901234567890"
        self.assertEqual(31, len(expectation))
        result = safexl.worksheet_name_sanitization(input_name)
        self.assertEqual(expectation, result)

    def test_empty_string_causes_error(self):
        input_name  = ""
        with self.assertRaises(safexl.toolkit.ExcelError):
            result = safexl.worksheet_name_sanitization(input_name)

    def test_invalid_characters(self):
        input_name  = "\\/*[]:?a"
        expectation = "a"
        result = safexl.worksheet_name_sanitization(input_name)
        self.assertEqual(expectation, result)
