PyInstaller -F --hidden-import webdriver_manager.firefox --hidden-import tkcalendar --hidden-import "babel.numbers"--hidden-import xlwings --hidden-import pandas --hidden-import tabula-py --add-data "tabula-1.0.5-jar-with-dependencies.jar;./tabula/" --icon=biourjaLogo.ico westplains_gui.py --onefile -w