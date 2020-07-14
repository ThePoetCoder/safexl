# safexl - The Safe Way to Excel
##### A wrapper around the pywin32 module for easier use and automated cleanup of Excel Application COM objects
The pywin32 library grants extraordinary capabilities to interact with Windows applications, 
but includes many oddities usually learned through trial and error, as seen in Stack Overflow posts such as these:
* [COM: excelApplication.Application.Quit() preserves the process](https://stackoverflow.com/questions/18421457/com-excelapplication-application-quit-preserves-the-process)
* [Can't close Excel completely using win32com on Python](https://stackoverflow.com/questions/10221150/cant-close-excel-completely-using-win32com-on-python)
* [Loading addins when Excel is instantiated programmatically](https://stackoverflow.com/questions/213375/loading-addins-when-excel-is-instantiated-programmatically)
* [AutoFilter method of Range class failed (Dispatch vs EnsureDispatch)](https://stackoverflow.com/questions/22930751/autofilter-method-of-range-class-failed-dispatch-vs-ensuredispatch)

My experience with automating Excel using pywin32 lead me to create `safexl`, a pywin32 wrapper centered around easier use and 
automated cleanup of Excel Application COM objects in Python. The main functionality of this package is a context-managed 
`application` generator that you can use inside a `with` block, built with some pywin32 best practices in place and a few psutil 
tools focused on working with Excel.

----------------------------------------------------------------------------------------------------------------------------------

## Install
You can import the library with pip:
```cmd
pip install safexl
```
Or you can download it [here](github.com/ThePoetCoder/safexl) at GitHub.

This module requires:
* [pywin32](https://pypi.org/project/pywin32/)
* [psutil](https://pypi.org/project/psutil/)

----------------------------------------------------------------------------------------------------------------------------------

## Usage
This package makes writing pywin32 code in Python as simple as:
```python
import safexl

with safexl.application(kill_after=False, maximize=True, include_addins=True) as app:
    wb = app.Workbooks.Add()
    ws = wb.ActiveSheet
    rng = ws.Range("B1")
    rng.Value = "Hello, World!"
    rng.Interior.Color = safexl.colors.rgbRed  # colors are included
    rng.EntireColumn.AutoFit()
    ws.Columns("A").Delete(Shift=safexl.xl_constants.xlToLeft)  # constants are included

# This results in Excel being opened to a Sheet where cell "A1" has 'Hello, World!' in it with a red background
```

If you've programmatically worked with Excel in a Win32 environment before, this code should look very familiar, 
as I am not altering the COM object itself before yielding it to you inside a `with` block; I am instead providing 
a means to create and delete it more easily. 

_If you would like to alter the COM object (for things like turning off ScreenUpdating
 while your code runs), then please see the **Performance** section near the bottom._

In this way, the following two code snippets will have the same effect:
#### 1.) without safexl
```python
import pythoncom
import win32com.client

pythoncom.CoInitialize()
app = win32com.client.Dispatch("Excel.Application")
try:
    #######################
    # Your code goes here #
    #######################
finally:
    app.Quit()
    del app
    pythoncom.CoUninitialize()
```

#### 2.) with safexl
```python
import safexl

with safexl.application(kill_after=True) as app:
    #######################
    # Your code goes here #
    #######################
```
As you can see, using safexl results in a lot less boilerplate code, from 9 lines to 2.

The `application` wrapper comes with 3 boolean parameters to indicate what you would like to do with the application once your 
`with` block is complete:
1. `kill_after` - kill the Excel process upon leaving the `with` block
2. `maximize` - Optional / Defaults to `True` - Will not be used if you set `kill_after=True`. Maximizes each Excel Window for
each Workbook added during the `with` block.
3. `include_addins` - Optional / Defaults to `False` - Will not be used if you set `kill_after=True`. Loads your installed Excel 
Add-ins to the newly created instance (with a performance hit to do so).

In the event of an error occuring inside your `with` block, the `safexl.application` cleanup process will carefully remove any new
workbooks you've opened in Excel, leaving any workbooks you already had open prior to the `with` block untouched. The same goes 
for if you chose to set `kill_after=True`; only the Workbooks you create inside the `with` block will be closed.
In addition to the `application` wrapper, I have included an handful of other tools to make working with Excel even easier, including:

* is_excel_open()
* kill_all_instances_of_excel()
* close_workbooks(app, workbooks)
* see_excel(app, window_state)
* workbooks_currently_open(app)
* last_row(worksheet)
* last_column(worksheet)
* worksheet_name_sanitization(worksheet_name)

----------------------------------------------------------------------------------------------------------------------------------

## Performance
A number of performance enhancing options can be set on Excel Application objects, and will come in handy most whenever you are 
working with large workbooks and amounts of data. In my balance between allowing you the most freedom to do what you wish with the 
application object and wrapping your object for safer error handling, I am yielding a bare pywin32 application object to you 
inside the `with` block. If you wish to take advantage of the various performance enhancing settings available natively in the 
Excel Application, I suggest using your own error handling inside the `with` block, to verify that the settings get switched back 
to normal when you're finished, even if you encounter an error during your work. Using safexl in this way would look something 
like this:
```python
import safexl

with safexl.application(kill_after=False) as app:
    try:
        app.ScreenUpdating = False
        app.DisplayStatusBar = False
        app.EnableEvents = False

        wb = app.Workbooks.Add()
        # can only set calculation once at least 1 workbook is open
        app.Calculation = safexl.xl_constants.xlCalculationManual

        #######################
        # Your code goes here #
        #######################
        
    except Exception as e:
        # if you don't re-raise the error here, you will not be warned that an error occured 
        # or get the benefit of reading the error message
        raise e
    
    else:
        pass
    
    finally:
        app.ScreenUpdating = True
        app.DisplayStatusBar = True
        app.EnableEvents = True
        app.Calculation = safexl.xl_constants.xlCalculationAutomatic

```

##### A note on setting the Calculation
Unfortunately, due to an oddity in the Excel Application OOP design, even though the Calculation mode is set on the Application object 
(instead of the Workbook object) if no workbooks are open or visible in your instance of the Application, then the constant for 
an "#N/A" error is returned, as seen by code like this:
```
>>> import win32com.client
>>> import pythoncom
>>> pythoncom.CoInitialize()
>>> app = win32com.client.Dispatch("Excel.Application")
>>> app.Calculation  # expect constant for #N/A
-2146826246
>>> wb = app.Workbooks.Add()
>>> app.Calculation  # expect XlCalculation constant
-4105
>>> wb.Close()
>>> app.Calculation  # expect constant for #N/A
-2146826246
```
More can be read about how Excel handles Calculation modes at these links:
* [How Excel determines the current mode of calculation](https://docs.microsoft.com/en-us/office/troubleshoot/excel/current-mode-of-calculation)
* [Excel: 'Unable to set the Calculation property of the Application class'](https://stackoverflow.com/questions/275630/excel-unable-to-set-the-calculation-property-of-the-application-class)

Suffice it to say, even though we think about the calculation mode being an attribute of each individual workbook, it is actually 
__set__ at the application level. I'm assuming this was for performance and/or sanity reasons, but the end result is that you are unable to 
get or set a proper Calculation mode for the application until you open a workbook first.

## Cookbook

##### Create & Save Workbook without viewing Application
```python
import safexl

with safexl.application(kill_after=True) as app:
    wb = app.Workbooks.Add()
    
    #######################
    # Your code goes here #
    #######################
    
    wb.SaveAs("Cookbook.xlsx")
    wb.Close()
```

##### Create a Workbook & View it Without Saving
```python
import safexl

with safexl.application(kill_after=False, maximize=True, include_addins=True) as app:
    wb = app.Workbooks.Add()

    #######################
    # Your code goes here #
    #######################
```

##### Minimize All Excel Windows that are Currently Open
```python
import safexl

with safexl.application(kill_after=True) as app:
    safexl.see_excel(app.Workbooks, safexl.xl_constants.xlMinimized)
```

##### Send Pandas Dataframe to Excel Worksheet
```python
import safexl
import pandas as pd
data = {
    'A': [1, 2, 3],
    'B': [4, 5, 6],
    'C': [7, 8, 9]
    }
df = pd.DataFrame(data)

with safexl.application(kill_after=False) as app:
    wb = app.Workbooks.Add()
    ws = wb.ActiveSheet

    df.to_clipboard(excel=True)
    ws.Paste()
    ws.Range("A1").Select()  # Otherwise entire dataframe range will be selected upon viewing
```

----------------------------------------------------------------------------------------------------------------------------------
## Similar Packages to Consider
* [xlwings](https://docs.xlwings.org/en/stable/)
* [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)
* [XlsxWriter](https://xlsxwriter.readthedocs.io/)

## Contact Me
* [Email](mailto:ThePoetCoder@gmail.com)
