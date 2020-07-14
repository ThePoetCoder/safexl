# Copyright (c) 2020 safexl
import unittest
import safexl


class test_app_when_excel_is_not_running_at_onset(unittest.TestCase):
    def setUp(self):
        safexl.kill_all_instances_of_excel()
        self.assertFalse(safexl.is_excel_open())

    def tearDown(self):
        # want to be sure that we start and end each app test with a clean slate
        safexl.kill_all_instances_of_excel()
        self.assertFalse(safexl.is_excel_open())

    def test_no_error_kill_after(self):
        # No Error | Kill After - Excel not in proc @ end
        with safexl.application(kill_after=True, maximize=False, include_addins=False) as app:
            wb = app.Workbooks.Add()
            ws = wb.ActiveSheet
            ws.Range("A1").Value = 555
            self.assertTrue(safexl.is_excel_open())
            self.assertEqual(wb.Name, "Book1")
            self.assertEqual(ws.Name, "Sheet1")
            self.assertEqual(ws.Range("A1").Value, 555)
        self.assertFalse(safexl.is_excel_open())

    def test_error_kill_after(self):
        # Error | Kill After - Excel not in proc @ end
        with self.assertRaises(safexl.toolkit.ExcelError):
            with safexl.application(kill_after=True, maximize=False, include_addins=False) as app:
                wb = app.Workbooks.Add()
                ws = wb.ActiveSheet
                self.assertTrue(safexl.is_excel_open())
                # Error on this line
                ws.Name = "a*b*c"
                # None of the following characters are allowed in sheet names
                # ["\\", "/", "*", "[", "]", ":", "?"]
        self.assertFalse(safexl.is_excel_open())

    def test_no_error_alive_after(self):
        # No Error | Alive After - Excel in proc @ end
        with safexl.application(kill_after=False, maximize=False, include_addins=False) as app:
            wb = app.Workbooks.Add()
            ws = wb.ActiveSheet
            ws.Range("A1").Value = 555
            self.assertTrue(safexl.is_excel_open())
            self.assertEqual(wb.Name, "Book1")
            self.assertEqual(ws.Name, "Sheet1")
            self.assertEqual(ws.Range("A1").Value, 555)
        self.assertTrue(safexl.is_excel_open())

    def test_error_alive_after(self):
        # Error | Alive After - Excel not in proc @ end
        with self.assertRaises(safexl.toolkit.ExcelError):
            with safexl.application(kill_after=False, maximize=False, include_addins=False) as app:
                wb = app.Workbooks.Add()
                ws = wb.ActiveSheet
                self.assertTrue(safexl.is_excel_open())
                # Error on this line
                ws.Name = "a*b*c"
                # None of the following characters are allowed in sheet names
                # ["\\", "/", "*", "[", "]", ":", "?"]
        self.assertFalse(safexl.is_excel_open())


class test_app_when_excel_is_running_at_onset(unittest.TestCase):
    def setUp(self):
        with safexl.application(kill_after=False, maximize=False, include_addins=False) as prev_app:
            wb = prev_app.Workbooks.Add()
            self.assertEqual("Book1", wb.Name)
        self.assertTrue(safexl.is_excel_open())

    def tearDown(self):
        # want to be sure that we end each app test with a clean slate
        safexl.kill_all_instances_of_excel()
        self.assertFalse(safexl.is_excel_open())

    def test_no_error_kill_after(self):
        # No Error | Kill After - prev Excel open after, new Excel gone
        current_openfile_count = len(safexl.toolkit.excel_open_files())
        with safexl.application(kill_after=True, maximize=False, include_addins=False) as app:
            wb = app.Workbooks.Add()
            ws = wb.ActiveSheet
            ws.Range("A1").Value = 555
            self.assertTrue(safexl.is_excel_open())
            self.assertEqual(len(safexl.toolkit.excel_open_files()), current_openfile_count + 1)
            self.assertEqual(wb.Name, "Book2")
            self.assertEqual(ws.Name, "Sheet1")
            self.assertEqual(ws.Range("A1").Value, 555)
        self.assertEqual(len(safexl.toolkit.excel_open_files()), current_openfile_count)
        self.assertTrue(safexl.is_excel_open())

    def test_error_kill_after(self):
        # Error | Kill After - prev Excel open after, new Excel gone
        with self.assertRaises(safexl.toolkit.ExcelError):
            current_openfile_count = len(safexl.toolkit.excel_open_files())
            with safexl.application(kill_after=True, maximize=False, include_addins=False) as app:
                wb = app.Workbooks.Add()
                ws = wb.ActiveSheet
                self.assertTrue(safexl.is_excel_open())
                self.assertEqual(len(safexl.toolkit.excel_open_files()), current_openfile_count + 1)
                # Error on this line
                ws.Name = "a*b*c"
                # None of the following characters are allowed in sheet names
                # ["\\", "/", "*", "[", "]", ":", "?"]
        self.assertEqual(len(safexl.toolkit.excel_open_files()), current_openfile_count)
        self.assertTrue(safexl.is_excel_open())

    def test_no_error_alive_after(self):
        # No Error | Alive After - prev Excel open after, new Excel open after
        current_openfile_count = len(safexl.toolkit.excel_open_files())
        with safexl.application(kill_after=False, maximize=False, include_addins=False) as app:
            wb = app.Workbooks.Add()
            ws = wb.ActiveSheet
            ws.Range("A1").Value = 555
            self.assertTrue(safexl.is_excel_open())
            self.assertEqual(len(safexl.toolkit.excel_open_files()), current_openfile_count + 1)
            self.assertEqual(wb.Name, "Book2")
            self.assertEqual(ws.Name, "Sheet1")
            self.assertEqual(ws.Range("A1").Value, 555)
        self.assertEqual(len(safexl.toolkit.excel_open_files()), current_openfile_count + 1)
        self.assertTrue(safexl.is_excel_open())

    def test_error_alive_after(self):
        # Error | Alive After - prev Excel open after, new Excel gone
        with self.assertRaises(safexl.toolkit.ExcelError):
            current_openfile_count = len(safexl.toolkit.excel_open_files())
            with safexl.application(kill_after=False, maximize=False, include_addins=False) as app:
                wb = app.Workbooks.Add()
                ws = wb.ActiveSheet
                self.assertTrue(safexl.is_excel_open())
                self.assertEqual(len(safexl.toolkit.excel_open_files()), current_openfile_count + 1)
                # Error on this line
                ws.Name = "a*b*c"
                # None of the following characters are allowed in sheet names
                # ["\\", "/", "*", "[", "]", ":", "?"]
        self.assertEqual(len(safexl.toolkit.excel_open_files()), current_openfile_count)
        self.assertTrue(safexl.is_excel_open())
