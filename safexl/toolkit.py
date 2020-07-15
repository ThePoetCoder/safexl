# Copyright (c) 2020 safexl
from contextlib import contextmanager
import psutil
import pythoncom
import win32com.client
EXCEL_PROCESS_NAME = "EXCEL.EXE"

__all__ = [
    'is_excel_open',
    'kill_all_instances_of_excel',
    'close_workbooks',
    'see_excel',
    'workbooks_currently_open',
    'last_row',
    'last_column',
    'worksheet_name_sanitization',
    'application',
]


def is_excel_open() -> bool:
    """
    Simple wrapper around `psutil.process_iter()` searching for individual processes of EXCEL.EXE
    :return: bool - Indicating whether or not Excel is open
    """
    for proc in psutil.process_iter():
        try:
            if proc.name() == EXCEL_PROCESS_NAME:
                return True
        except psutil.AccessDenied:
            pass
    return False


def excel_open_files() -> list:
    """
    Simple wrapper around `psutil.process_iter()` searching for individual processes of EXCEL.EXE and returning
    all the filepaths of the open files. Used here only for testing purposes, when an `app` object cannot
    necessarily be passed as well, as is the case with `workbooks_currently_open`.
    :return: list - Full of filepaths, including all open files, addin files, etc. Note that prior to saving a file it is
                    given a .tmp filepath.
    """
    result = []
    for proc in psutil.process_iter():
        try:
            if proc.name() == EXCEL_PROCESS_NAME:
                result.extend([popenfile.path for popenfile in proc.open_files()])
        except psutil.AccessDenied:
            pass
    return result


def kill_all_instances_of_excel(app: 'win32com.client.Dispatch("Excel.Application")' = None) -> None:
    """
    Simple wrapper around `psutil.process_iter()` searching for individual processes of EXCEL.EXE, and killing each one it finds
    :param app: Optional win32com.client.Dispatch("Excel.Application") - Programmatic access to Excel application object
    :return: None
    """
    if app:
        # If application is passed, try to shut it down peacefully first
        close_workbooks(app, app.Workbooks)
        app.Quit()
        del app

    for proc in psutil.process_iter():
        try:
            if proc.name() == EXCEL_PROCESS_NAME:
                proc.kill()
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            # passing on psutil.NoSuchProcess avoids erroring out if race conditions
            # close Excel *between* finding it and killing it with psutil
            pass


def new_workbooks(app: 'win32com.client.Dispatch("Excel.Application")', workbooks_open_at_onset: iter) -> list:
    """
    Determines which workbooks are open currently in comparison to list of `workbooks_open_at_onset`, returns the delta
    :param app: win32com.client.Dispatch("Excel.Application") - Programmatic access to Excel application object
    :param workbooks_open_at_onset: iterable - Full of workbook COM objects that you want to close without saving
    :return: iterable - Full of workbook COM objects that are both open currently and not present in your `workbooks_open_at_onset`
    """
    paths_for_workbooks_open_at_onset = set(wb.FullName for wb in workbooks_open_at_onset)

    currently_open_workbooks = workbooks_currently_open(app)
    paths_for_currently_open_workbooks = set(wb.FullName for wb in currently_open_workbooks)

    paths_for_new_workbooks = paths_for_currently_open_workbooks - paths_for_workbooks_open_at_onset
    return [wb for wb in currently_open_workbooks if wb.FullName in paths_for_new_workbooks]


def close_workbooks(app: 'win32com.client.Dispatch("Excel.Application")', workbooks: iter) -> None:
    """
    Best practice pywin32 for close workbooks without saving
    :param app: win32com.client.Dispatch("Excel.Application") - Programmatic access to Excel application object
    :param workbooks: iterable - Full of workbook COM objects that you want to close without saving
    :return: None
    """
    for wb in workbooks:
        app.DisplayAlerts = 0
        wb.Close(SaveChanges=False)
        app.DisplayAlerts = 1


def see_excel(workbooks: iter, window_state: int) -> None:
    """
    Makes every window of every workbook passed visible, will ignore the PERSONAL workbook and anything else in your StartupPath
    :param workbooks: iterable - Full of workbook COM objects whose windows you wish to maximize, minimize, or normalize
    :param window_state: int - xl_constant for Window.WindowState, available options include:
                                 * safexl.xl_constants.xlMaximized = -4137
                                 * safexl.xl_constants.xlMinimized = -4140
                                 * safexl.xl_constants.xlNormal = -4143
    :return: None
    """
    for wb in workbooks:
        wb.Application.Visible = True

        # Ignore changing the visibility of any workbooks you have set to open in your StartupPath
        # such as the PERSONAL.XLSB
        if wb.Application.StartupPath in wb.FullName:
            continue

        for window in wb.Windows:
            window.Visible = True
            window.WindowState = window_state


def workbooks_currently_open(app: 'win32com.client.Dispatch("Excel.Application")') -> list:
    """
    Turns 'win32com.client.CDispatch' returned by `app.Workbooks` into Python list
    :param app: win32com.client.Dispatch("Excel.Application") - Programmatic access to Excel application object
    :return: list - Full of workbook COM objects currently open in Excel. Note that prior to saving a file it is given a generic
                    non-path such as 'Book1', 'Book2', etc.
    """
    return [wb for wb in app.Workbooks]


def last_row(worksheet) -> int:
    """
    Quick way to determine the number of rows in a worksheet. Assumes that data is within the `CurrentRegion` of cell A1.
    :param worksheet: Excel Worksheet COM object, such as the one created by code like:
        app = win32com.client.Dispatch("Excel.Application")
        wb = app.Workbooks.Add()
        ws = wb.ActiveSheet
    :return: int - indicating the number of rows a worksheet is using up
    """
    return worksheet.Range("A1").CurrentRegion.Rows.Count


def last_column(worksheet) -> int:
    """
    Quick way to determine the number of columns in a worksheet. Assumes that data is within the `CurrentRegion` of cell A1.
    :param worksheet: Excel Worksheet COM object, such as the one created by code like:
        app = win32com.client.Dispatch("Excel.Application")
        wb = app.Workbooks.Add()
        ws = wb.ActiveSheet
    :return: int - indicating the number of columns a worksheet is using up
    """
    return worksheet.Range("A1").CurrentRegion.Columns.Count


def worksheet_name_sanitization(worksheet_name: str) -> str:
    """
    Tool to cleanse worksheet names of common problems
    :param worksheet_name: str - String of name you're about to assign to a worksheet
    :return: str - String that won't cause an error when assigned to a worksheet. Note this function will throw an error
                   itself if the result of removing the invalid worksheet name characters leaves you with an empty string only
    """
    for char in ("\\", "/", "*", "[", "]", ":", "?"):
        worksheet_name = worksheet_name.replace(char, "")
    if not worksheet_name:
        raise ExcelError("Worksheet name cannot be empty string")
    return worksheet_name[:31]


@contextmanager
def application(
        kill_after: bool,
        maximize: bool = True,
        include_addins: bool = False,
) -> 'win32com.client.Dispatch("Excel.Application")':
    """
    Wrapper for the pywin32 interface for handling programmatic access to the Excel Application from Python on Windows.
    This context-managed generator function will yield an Excel application COM object that is safer to use
    than a bare call to `win32com.client.Dispatch("Excel.Application")`.
    :param kill_after: bool - Programmatic access to Excel will be removed outside the `with` block, but this argument
                              designates whether you wish to close the actual application as well, or to leave it running
    :param maximize: Optional bool - Defaults to `True`. Specifies what you would like to happen with your application windows,
                                     whether you want to maximize them or minimize them. This bool is contingent upon both
                                     your selection for `kill_after` and whether you run into an error during your `with` block.
                                     If `kill_after=True` then the code that maximizes or minimizes your windows will never
                                     be run, and the same goes for if Python encounters an error prior to reaching the end
                                     of your `with` block. In that way, the following 4 code snippets will have the same effect:
                                         1.) with safexl.application(kill_after=True) as app:
                                                 pass
                                         2.) with safexl.application(kill_after=True, maximize=True) as app:
                                                 pass
                                         3.) with safexl.application(kill_after=True, maximize=False) as app:
                                                 pass
                                         3.) with safexl.application(kill_after=False) as app:
                                                 raise
    :param include_addins: Optional bool - Defaults to `False`. As a feature (read "bug") Excel will not automatically
                                           load/open your *installed & active* Excel addins when an instance is called from code,
                                           neither Python nor VBA, as discussed in the following links:
                                             * https://stackoverflow.com/questions/213375/loading-addins-when-excel-is-instantiated-programmatically
                                             * https://www.mrexcel.com/board/threads/add-in-doesnt-load.849923/
                                           The `include_addins` parameter indicates whether you would like to include the addins
                                           you have previously set to 'Installed' to show up in the Excel instance created.
                                           Similar to the `maximize` parameter, if an error occurs in your `with` block, or if you
                                           set `kill_after=True` it doesn't matter what value `include_addins` is, as that part of
                                           the code will not be executed. Please note there is a performance hit taken by setting
                                           this parameter to `True`, especially if you or your user has many addins installed.
    :return: win32com.client.Dispatch("Excel.Application") - Wrapped to follow best practices and clean up after itself
             Note, I specifically chose `Dispatch` over both `DispatchEx` and `EnsureDispatch` to avoid some odd bugs
             that can crop up with those methods, as discussed further on SO:
               * https://stackoverflow.com/questions/18648933/using-pywin32-what-is-the-difference-between-dispatch-and-dispatchex
               * https://stackoverflow.com/questions/22930751/autofilter-method-of-range-class-failed-dispatch-vs-ensuredispatch

    """
    open_at_onset = is_excel_open()
    pythoncom.CoInitialize()
    _app = win32com.client.Dispatch("Excel.Application")
    if open_at_onset:
        workbooks_open_at_onset = workbooks_currently_open(_app)
    else:
        workbooks_open_at_onset = []

    try:
        # For use inside a `with` block, with exceptions caught and cleaned up for you
        yield _app

    except Exception as e:
        err_msg = e

    else:
        err_msg = ""

    finally:
        workbooks_opened_during_with_block = new_workbooks(_app, workbooks_open_at_onset)
        if kill_after or err_msg:
            # If user wants to kill the app after the with block OR if an error occurs
            # close everything that was opened during this `with` block alone
            if workbooks_open_at_onset:
                # close newly created workbooks instead of killing the entire app
                close_workbooks(_app, workbooks_opened_during_with_block)
            else:
                # kill the app's entire presence on computer
                kill_all_instances_of_excel()
        else:
            # Excel Application oddity where addins are not visible on the ribbon even when installed
            # when app instance is created via code. Thankfully the `.Installed` attribute remains intact,
            # and to make your addins show up on the ribbon, you must turn the installed addins off and then on again...
            # See docstring for links describing the problem, this solution, and more details.
            if include_addins:
                for add_in in _app.AddIns:
                    if add_in.Installed:
                        add_in.Installed = False
                        add_in.Installed = True

            # Running `see_excel` at the end here
            # makes sure that no Excel instances are left running in the background
            # as it is still easy to forget to make everything visible before leaving
            # a successful `with` block
            if maximize:
                see_excel(workbooks_opened_during_with_block, -4137)  # xlMaximized
            else:
                see_excel(workbooks_opened_during_with_block, -4140)  # xlMinimized

        del _app
        pythoncom.CoUninitialize()
        if err_msg:
            raise ExcelError(err_msg)


class ExcelError(Exception):
    pass
