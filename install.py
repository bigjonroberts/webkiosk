import sys
from os.path import join, isfile
import os
import urllib.request 
from zipfile import ZipFile
from win32com.shell import shell, shellcon



# download chrome driver
if not isfile('chromedriver.exe'):

    chrome_driver_zipfilename = 'chromedriver_win32.zip'
    chrome_driver_url='http://chromedriver.storage.googleapis.com/2.11/' + chrome_driver_zipfilename
    urllib.request.urlretrieve(chrome_driver_url, chrome_driver_zipfilename)

    # extract chrome driver
    zip_ref = ZipFile(chrome_driver_zipfilename, 'r')
    zip_ref.extractall(os.path.dirname(os.path.realpath(__file__)))
    zip_ref.close()

    # cleanup chrome driver zip
    os.remove(chrome_driver_zipfilename)

# get url and refresh interval


# delete existing shortcuts

# install new shortcut
script_path = os.path.realpath(__file__)
working_dir = os.path.dirname(script_path)
pyw_executable = join(working_dir, "python3", "pythonw.exe")
shortcut_filename = "webkiosk - %s.lnk" % url
shortcut_paths = 

if sys.argv[1] == '-install':

    # Get paths to the desktop and start menu
    desktop_path = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOPDIRECTORY, None, 0)
    startmenu_path = shell.SHGetFolderPath(0, shellcon.CSIDL_STARTMENU, None, 0)
	startup_path = shell.SHGetFolderPath(0, shellcon.CSIDL_STARTUP, None, 0)

    # Create shortcuts.
    for path in [desktop_path, startmenu_path, startup_path]:
        create_shortcut(pyw_executable,
                    "Web Kiosk for %s, refreshed every %d seconds" % (url, interval),
                    join(path, shortcut_filename),
                    script_path,
                    working_dir)
