import time
import argparse
import sys
from os.path import join, isfile
import os
import urllib.request 
from zipfile import ZipFile
from win32com.shell import shell, shellcon
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

def install(site, seconds):
    # get url and refresh interval
    website = input("Please input desired website default(%s):" % site)
    if website == '':
        website = site
    interval = input("Please input desired interval in seconds default(%d):" % seconds)
    if interval == '':
        interval = seconds

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

    # install new shortcuts
    script_path = os.path.realpath(__file__)
    working_dir = os.path.dirname(script_path)
    pyw_executable = join(working_dir, "python3", "pythonw.exe")
    shortcut_filename = "webkiosk - %s.lnk" % url
    shortcut_paths = {'desktop_path': shellcon.CSIDL_DESKTOPDIRECTORY,
            'startmenu_path': shellcon.CSIDL_STARTMENU,
            'startup_path': shellcon.CSIDL_STARTUP}

    # TODO: find and prompt to delete existing shortcuts in target directories starting with 'kiosk - '

    # Create shortcuts.
    for location_name,path_id in shortcut_paths:
        create_shortcut(pyw_executable,
                    "Web Kiosk for %s, refreshed every %d seconds" % (website, interval),
                    join(shell.SHGetFolderPath(0, path_id, None, 0), shortcut_filename),
                    script_path,
                    working_dir)

parser = argparse.ArgumentParser(description='Simple Web Display Kiosk')

parser.add_argument('-website', '-w', default='http://members.geekdom.com/', help='website to launch and refresh, defaults to http://members.geekdom.com/')
parser.add_argument('-interval', '-n', type=int, default=300, help='refresh interval in seconds, defaults to 300 seconds for 5 minutes')
parser.add_argument('-install', '-i', action='store_true', help='install to the directory where you run this script and then install shortcuts to start menu, desktop, and startup folder')

args = parser.parse_args()

if args.install:
    install(args.website, args.interval)


options = Options()
options.add_argument("--kiosk")
driver = webdriver.Chrome(chrome_options=options)

driver.get(args.website)

while True:
    time.sleep(args.interval)
    driver.refresh()
driver.close()
