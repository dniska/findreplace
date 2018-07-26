# findreplace
Using Python 2.7
Python script to find &amp; replace specified strings inside a given directory of files.

Unstable. Tkinter GUI is rudimentary.

Minimal testing has been performed, though the current operations have been successfully performed in a production environment. The format has been altered since and placed into individual functions. The individual functions were once each separate python scripts; this script combines each script into one and allows for user input.

For compilation into standalone executable, PyInstaller is sufficient.

Use 'PyInstaller --onefile frames.py'

If you run into ImportError in relation to np_datetime from the pandas library, add a file named hook-pandas.py to the PyInstaller/hooks directory. This file only contains one line and can be found in this repository.

If your error is specific to np_timedeltas, change 'pandas._libs.tslibs.np_datetime' to 'pandas._libs.tslibs.np_timedeltas' inside the hook-pandas.py file
