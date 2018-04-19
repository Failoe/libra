"""

Tableau Libra is a script written by Ryan Lynch to help handle Tableau
Workbooks that were saved in previous versions of Tableau Desktop. It
aims to provide information on workbooks and save time for users by
preventing them from opening workbook files in incompatible versions.

Additional functionalities to be added:
	Automatic corrupted workbook diagnosis and repair
	Script fails to execute on some machines
	Loading indicator for element_parse()
	Fix issue with it not appearing for "TWBX Files"
		Currently it works for "Tableau Workbook" and "Tableau Packaged Workbook" files
	Add options menu
	Add handling for TPublic and TReader
	Give option for install path
	Refactor Calculated Field code


	User Feature Requests:
		Add context menu for open in native
		Show "Created in version" and "Most recently saved in version" as different fields

Finished features:
	Replacing the default "Open" option for workbooks - Created Shepherd for this
	Remove dependance on C:/Program Files/Tableau
		This is jank and it is unneeded due to tableau_paths()
	-Actually use Python functions...
	Sort the version selector dropdown from newest to oldest versions
	Make sure we only address Tableau Desktop entries in the registry
	Add debug logging
	.twb files have issues opening
	Newest Version button works intermittently
	Detailed information on the structure of the workbook
	Info -> Filter Count
	Highlight Red Flags in Workbooks:
		CustomSQL
		ODBC

Change Log:
	v0.1.4
		Filter count added to info panel
		ODBC and CustomSQL alerts (in red)

	v0.1.5
	Added logging

Share with Wess Woo.
"""

from tkinter import *
from tkinter import ttk
import re
import sys
import os
import subprocess
from bs4 import BeautifulSoup as bs
import winreg
import zipfile
import random
import traceback
import logging

__author__ = "Ryan Lynch"
__version__ = "0.1.5"
__maintainer__ = "Ryan Lynch"
__email__ = "rlynch@tableau.com"
__status__ = "Development"
__credits__ = ["Ryan Lynch", "Carmen Hernando"]

logging.basicConfig(
	format="%(asctime)s [%(levelname)s]\t%(message)s",
	level=logging.INFO)
logger = logging.getLogger("libra.logger")

logger.setLevel(logging.INFO)


def initialize():
	global filearg
	logger.info("Initializing Libra")
	try:
		filearg = sys.argv[1]
	except IndexError:
		""" TESTING WORKBOOKS """
		# filearg = "C:/Users/rlynch/Downloads/Support Case Queue.twb"
		# filearg = r"C:\Users\rlynch\Downloads\highlight_parameter_controls.twbx"
		# filearg = "C:/Users/rlynch/Downloads/FailMetricHeatmap.twb"
		# filearg = "C:/Users/rlynch/Downloads/Section 3 to share with Tableau support.twbx"
		# filearg = "C:/Users/rlynch/Downloads/University Division TAT.twbx"
		# filearg = "C:/Users/rlynch/Downloads/pytest.twb"
		# filearg = "C:/Users/rlynch/Downloads/DTSS Facilities Review.twbx"
		# filearg = r"C:\pytest.twb"
		# filearg = r"C:\Users\rlynch\Downloads\pytest.twb"
		# filearg = r"C:\Users\rlynch\Downloads\Table calculation not working_Vivek.twbx"
		filearg = r"C:\Users\rlynch\Downloads\Global_Workbook.twbx"
		# filearg = r"C:\Users\rlynch\Downloads\ODBCConnectionTest-SQLite.twb"
		# filearg = r"C:\Users\rlynch\Downloads\CustomSQL_ODBCConnectionTest-SQLite.twb"
		pass
	logger.info("filearg: {}".format(filearg))


titles = [
	"Libra - Workbook Magic",
	"TabLibra",
	"Something",
	"Dr Horrible's Sing-Along-Script",
	"Twilight's Friendship Report",
	"Workbooks are Magic",
	"Llama",
	"Linguine",
	"arbiL",
	"Hail Python",
	"Totally Legit Software",
	"You're awesome!",
	"Have a great day!",
	"Name Pending",
	"Drop and give me 20!",
	"Mo' data mo' problems",
	"Sunset is best pony",
	"Let me take a selfie",
	"I'll never let go, Jack",
	"Will you bend the knee?",
	"You know nothing",
	"Ryuu ga waga teki wo kurau",
	"Mama always said: dont lose",
	"It's HIIIIGH noon",
	"NYEH HEH HEH",
	"Algebraic!",
	"TWBX or GTFO",
	"MTT Brand Libra Product",
	"This product contains chemicals known to the state of California to cause cancer.",
	"This product may cause spontaneous break dancing in laboratory animals"]


def wkbk_alerts(filexml, frame):

	gui_style = ttk.Style()
	gui_style.configure('alert.TLabel', foreground='red')
	# gui_style.configure('My.TFrame', background='#334353')
	if "class='genericodbc'" in filexml:
		ttk.Label(frame, text="ALERTS:", style='alert.TLabel').grid(column=1, row=4)
		ttk.Label(frame, text="Generic ODBC\nConnection", style='alert.TLabel').grid(column=2, row=4, sticky=W)
	if re.search(r"<relation.+?type='text'", filexml):
		ttk.Label(frame, text="ALERTS:", style='alert.TLabel').grid(column=1, row=4)
		ttk.Label(frame, text="Custom SQL", style='alert.TLabel').grid(column=3, row=4, sticky=W)


def tableau_paths():
	""" Fetches the installed versions and their paths
		from the registry. Returns a tuple ex:
		(Tableau 10.0, 'C:/Tableau/Tableau 10.0')"""

	logger.debug("Getting Tableau paths from registry.")

	paths = []
	key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "Software\\Tableau", access=winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	values = winreg.QueryInfoKey(key)
	for subkeys in range(values[0]):
		try:
			keys = winreg.EnumKey(key, subkeys)
			sub_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "Software\\Tableau\\{0}\\{1}".format(keys, "Directories"), access=winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
			if re.match(r'(Tableau \d+\.\d+)', keys):
				paths.append((keys, winreg.QueryValueEx(sub_key, "Application")[0]))
			sub_key.Close()
		except WindowsError:
			pass

	key.Close()

	logger.debug("Tableau Paths: ")
	[logger.debug("{} {}".format(x, y)) for x, y in paths]

	return(paths)


def open_tableau(file_path, tkwindow, close=True):
	""" Opens the Tableau Desktop version provided.
		Does not check if it is installed """
	logger.debug("Opening Tableau version: {}".format(file_path))
	logger.debug("File Path to open: {}".format(filearg))

	subprocess.Popen([file_path, filearg])
	if close:
		tkwindow.destroy()


def add_to_clipboard(text):
	""" Adds the provided text to the clipboard """
	clip = Tk()
	clip.withdraw()
	clip.clipboard_clear()
	clip.clipboard_append(text)
	clip.destroy()

	logger.debug("Added text to clipboard: {}".format(text))


def download_version(versionarg, tkwindow):
	""" Opens a browser page to download the native
		version of the workbook. """
	modifiedversion = re.sub(r'(\.0$)', '', versionarg)
	url = "https://www.tableau.com/support/releases/desktop/{0}".format(modifiedversion)
	logger.debug("Opening URL: {}".format(url))
	os.startfile(url)

	tkwindow.destroy()


"""
Workbook scanner
	Workbook Filesize
	Type Twb/Twbx
	Workbook Name

	Datasources
	Calculated fields
	Worksheets
	Dashboards
	Parameters
	Custom SQL
"""


def element_parse(filexml, tkwindow):
	soup = bs(filexml, "lxml")
	""" Finds calculated fields in workbook """
	total_calcs = 0

	newcalcs = soup.find_all(formula=True)

	for calc in newcalcs:
		if any([calc.parent.name == _ for _ in ['column']]):
			if (calc.parent.parent.name == 'datasource-dependencies' and calc.parent.parent['datasource'] == 'Parameters') or (calc.parent.parent.name == 'datasource' and calc.parent.parent['name'] == 'Parameters'):
				pass
			else:
				if calc.parent.parent.name == "connection":
					total_calcs += 1
					logger.debug("Calc Field: {}".format(calc['column']))
				elif calc.parent.parent.name == "datasource":
					total_calcs += 1
					if calc.parent.has_attr('caption'):
						logger.debug("Calc Field: {}".format(calc.parent['caption']))
					else:
						logger.debug("Calc Field: {}".format(calc.parent['name']))

	logger.info('Calcs: {}'.format(total_calcs))

	""" datasources """
	datasources_list = soup.datasources.find_all('datasource')
	logger.info("Datasources: {}".format(len(datasources_list)))

	for ds in datasources_list:
		""" Finds datasource's name and alias """
		if ds.get('caption'):
			ds_name = ds.get('caption')
		else:
			ds_name = ds.get('name')

		logger.debug('Datasource: {}'.format(ds_name))

	""" parameters """
	param_tag = soup.find("datasource", {"name": "Parameters", "hasconnection": "false"})
	params = []
	if param_tag:
		params = param_tag.find_all('column')
	logger.info("Parameters: {}".format(len(params)))

	""" worksheets """
	worksheet_list = soup.worksheets.find_all('worksheet')

	logger.info("Worksheets: {}".format(len(worksheet_list)))
	[logger.debug("Worksheet: {}".format(i['name'])) for i in worksheet_list]

	""" dashboards """
	dashboard_list = []
	if soup.dashboards:
		dashboard_list = soup.dashboards.find_all('dashboard')

	logger.info("Dashboards: {}".format(len(dashboard_list)))
	[logger.debug("Dashboard: {}".format(i['name'])) for i in dashboard_list]

	""" filters """
	filter_list = soup.find_all('filter')

	logger.info("Filters: {}".format(len(filter_list)))
	[logger.debug("Filter: {}".format(i['column'])) for i in filter_list]

	""" Workbook Info Popup Code """
	element_root = Toplevel(tkwindow)
	try:
		element_root.iconbitmap('C:/Tableau Scripts/favicon.ico')
	except TclError:
		pass

	element_main = ttk.Frame(element_root, padding="3 3 12 12")
	element_main.grid(column=0, row=0, sticky=(N, W, E, S))
	element_main.columnconfigure(0, weight=1)
	element_main.rowconfigure(0, weight=1)

	ttk.Label(element_main, text="Workbook Info").grid(column=1, row=1, sticky=E)

	ttk.Label(element_main, text="Worksheets:").grid(column=1, row=2, sticky=E)
	ttk.Label(element_main, text="{0}".format(len(worksheet_list), ("" if len(worksheet_list) == 1 else "s"))).grid(column=2, row=2, sticky=W)

	ttk.Label(element_main, text="Dashboards:").grid(column=1, row=3, sticky=E)
	ttk.Label(element_main, text="{0}".format(len(dashboard_list))).grid(column=2, row=3, sticky=W)

	ttk.Label(element_main, text="Parameters:").grid(column=1, row=4, sticky=E)
	ttk.Label(element_main, text="{0}".format(len(params))).grid(column=2, row=4, sticky=W)

	ttk.Label(element_main, text="Calc Fields:").grid(column=1, row=5, sticky=E)
	ttk.Label(element_main, text="{0}".format(total_calcs)).grid(column=2, row=5, sticky=W)

	ttk.Label(element_main, text="Filters:").grid(column=1, row=6, sticky=E)
	ttk.Label(element_main, text="{0}".format(len(filter_list))).grid(column=2, row=6, sticky=W)

	for child in element_main.winfo_children():
		child.grid_configure(padx=3, pady=3)

	element_root.bind('<Return>', lambda x: element_root.destroy())
	element_root.bind('<Escape>', lambda x: element_root.destroy())
	element_root.focus()

	""" Runs the tk window """
	element_root.mainloop()


def open_workbook():
	# this block checks if it is a .twb or .twbx and sets the fileread variable to the .twb file
	if filearg[-4:] == ("twbx" or "TWBX"):
		archive = zipfile.ZipFile(filearg, 'r')
		twbname = [s for s in archive.namelist() if ".twb" in s][0]
		fileread = archive.read(twbname).decode("utf-8", "replace")
	else:
		file = open(filearg, 'r', errors="replace")
		fileread = file.read()

	# this checks if the workbook uses source-build to store the build info
	sourceRE = re.compile(r"<workbook source-build='([\d\.]+?) \(([\d\.]+?)\)")
	sourcecheck = sourceRE.search(fileread)

	if sourcecheck is not None:
		# if source-build is present in the XML it uses that to find the build/version
		wkbk_build = sourcecheck.group(2)
		wkbk_version = sourcecheck.group(1)
	else:
		# if sourcebuild isn't present it finds the build in the
		# commented section and the version after "version="
		buildRE = re.compile(r"<!-- build (.+?) ")
		versionRE = re.compile(r"<workbook.+?version='([\d\.]*)")

		buildsearch = buildRE.search(fileread)
		versionsearch = versionRE.search(fileread)

		wkbk_build = buildsearch.group(1)
		wkbk_version = versionsearch.group(1)

	logger.info("Workbook's version: {}".format(wkbk_version))
	logger.info("Workbook's build: {}".format(wkbk_build))

	return wkbk_build, wkbk_version, fileread


def create_popup(tableau_paths):
	""" All the popup window code """
	wkbk_build, wkbk_version, fileread = open_workbook()

	""" Orders the list from newest version to oldest """
	tableau_paths.sort(key=lambda x: float(x[0].replace('Tableau ', '')), reverse=True)

	root = Tk()
	root.title("Libra {}".format(__version__) if random.random() > .3 else random.choice(titles))

	try:
		root.iconbitmap('C:/Tableau Scripts/favicon.ico')
	except TclError:
		pass

	mainframe = ttk.Frame(root, padding="3 3 12 12")
	mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
	mainframe.columnconfigure(0, weight=1)
	mainframe.rowconfigure(0, weight=1)

	ttk.Label(mainframe, text="Version:").grid(column=2, row=1, sticky=E)
	ttk.Label(mainframe, text="{0}".format(wkbk_version)).grid(column=3, row=1, sticky=W)

	ttk.Label(mainframe, text="Build:").grid(column=2, row=2, sticky=E)
	ttk.Label(mainframe, text="{0}".format(wkbk_build)).grid(column=3, row=2, sticky=W)

	shortversion = re.match(r"(\d+?\.\d+?)", wkbk_version).group(1)

	logger.debug("Shortversion: {}".format(shortversion))

	for fullname, path in tableau_paths:
		if shortversion in fullname:
			# if the Tableau version is installed do this
			ttk.Button(mainframe, text="Open Native", command=lambda x=path + "bin\\tableau.exe": open_tableau(x, root), width=16).grid(column=1, row=1, sticky=W)
			root.bind('<Return>', lambda x: open_tableau(path + "bin\\tableau.exe", root))
			break
		else:
			# If the Tableau version isn't installed, do this
			ttk.Button(mainframe, text="Download Version", command=lambda x=wkbk_version: download_version(x, root), width=16).grid(column=1, row=1, sticky=W)
			# Set Return key to open in newest version installed
			root.bind('<Return>', lambda x=path + "bin\\tableau.exe": open_tableau(x, root))

	ttk.Button(mainframe, text="Open Newest", command=lambda x=tableau_paths[0][1] + r"bin\tableau.exe": open_tableau(x, root), width=16).grid(column=1, row=2, sticky=W)
	ttk.Button(mainframe, text="Copy Build #", command=lambda x=wkbk_build: add_to_clipboard(str(x))).grid(column=2, row=3, sticky=W)
	ttk.Button(mainframe, text="(I)nfo", command=lambda x=wkbk_build: element_parse(fileread, root), width=16).grid(column=1, row=3, sticky=W)

	tkvar = StringVar(root)
	tkvar.set('Open in Version...')
	popupMenu = OptionMenu(mainframe, tkvar, *[x[0] for x in tableau_paths])
	popupMenu.grid(column=3, row=3)

	def change_dropdown(*args):
		open_tableau([path for build, path in tableau_paths if build == tkvar.get()][0] + r'bin\tableau.exe', root)

	tkvar.trace('w', change_dropdown)

	for child in mainframe.winfo_children():
		child.grid_configure(padx=5, pady=5)

	root.bind('<Escape>', lambda x: root.destroy())
	root.bind('<i>', lambda x: element_parse(fileread, root))

	wkbk_alerts(fileread, mainframe)

	""" Runs the tk window """
	root.mainloop()


def main():
	initialize()
	create_popup(tableau_paths())


try:
	main()
except Exception as err:
	with open('./LibraLog.txt', 'a') as logfile:
		print("Error occurred: {}".format(err))
		logfile.write(traceback.format_exc())

"""
# Code to pull build numbers and page links from
# https://www.tableau.com/support/releases

page = requests.get('https://www.tableau.com/support/releases')
soup = bs(page.content, 'html5lib')
versions_list = soup.find_all('div', attrs={'class': 'view-content'})
versions_list.pop(0)

for version in versions_list:
    builds = version.table.tbody.find_all('tr')
    for build in builds:
        print(build.td.a.text.split(' ')[0], build.td.a['href'])
"""
