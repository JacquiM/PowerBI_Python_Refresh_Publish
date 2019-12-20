import time
import os
import sys
import argparse
import psutil
from pywinauto.application import Application
from pywinauto import timings
from datetime import datetime


def type_keys(string, element):
    """Type a string char by char to Element window"""
    for char in string:
        element.type_keys(char)

# Parse arguments from cmd
parser = argparse.ArgumentParser()
parser.add_argument("workbook", help = "Path to .pbix file")
parser.add_argument("--workspace", help = "name of online Power BI service work space to publish in", default = "My workspace")
parser.add_argument("--refresh-timeout", help = "refresh timeout", default = 30000, type = int)
parser.add_argument("--no-publish", dest='publish', help="don't publish, just save", default = True, action = 'store_false' )
parser.add_argument("--init-wait", help = "initial wait time on startup", default = 15, type = int)
args = parser.parse_args()

timings.after_clickinput_wait = 1
WORKBOOK = args.workbook
WORKSPACE = args.workspace
INIT_WAIT = args.init_wait
REFRESH_TIMEOUT = args.refresh_timeout
PROCNAME = "PBIDesktop.exe"


# Kill running PBI
def kill():

	try:
		for proc in psutil.process_iter():
			# check whether the process name matches
			if proc.name() == PROCNAME:
				proc.kill()
		time.sleep(3)

		return True

	except Exception as e:

		print(e)
		return False

# Start PBI and open the workbook
def start():

	try:
		print("Starting Power BI")
		os.system('start "" "' + WORKBOOK + '"')
		print("Waiting ",INIT_WAIT,"sec")
		time.sleep(INIT_WAIT)
	except Exception as e:

		print(e)
		return False

# Create the application and the window
def create():

	try:
		print('Creating application and window')
		app = Application(backend='uia').connect(path=PROCNAME)
		win = app.window(title_re='.*Power BI Desktop')

		return win
	except Exception as e:

		print(e)
		return False

# Connect pywinauto
def connect(win):

	try:
		print("Identifying Power BI window")
		time.sleep(5)
		win.wait("enabled", timeout = 300)
		win.Save.wait("enabled", timeout = 300)
		win.set_focus()
		win.Home.click_input()
		win.Save.wait("enabled", timeout = 300)
		win.wait("enabled", timeout = 300)
	except Exception as e:

		print(e)
		return False

# Refresh
def refresh(win):

	try:
		print("Refreshing")
		win.Refresh.click_input()
		#wait_win_ready(win)
		time.sleep(5)
		print("Waiting for refresh end (timeout in ", REFRESH_TIMEOUT,"sec)")
		win.wait("enabled", timeout = REFRESH_TIMEOUT)
	except Exception as e:

		print(e)
		return False

# Save
def save(win):

	try:
		print("Saving")
		type_keys("%1", win)
		#wait_win_ready(win)
		time.sleep(5)
		win.wait("enabled", timeout = REFRESH_TIMEOUT)
	except Exception as e:

		print(e)
		return False

# Publish
def publish(win):

	try:
		if args.publish:
			print("Publish")
			win.Publish.click_input()
			publish_dialog = win.child_window(auto_id = "KoPublishToGroupDialog")
			publish_dialog.child_window(title = WORKSPACE).click_input()
			publish_dialog.Select.click()
			try:
				win.Replace.wait('visible', timeout = 10)
			except Exception:
				pass
			if win.Replace.exists():
				win.Replace.click_input()
			win["Got it"].wait('visible', timeout = REFRESH_TIMEOUT)
			win["Got it"].click_input()
	except Exception as e:

		print(e)
		return False

# Close
def close(win):

	try:
		print("Exiting")
		win.close()
	except Exception as e:

		print(e)
		return False

def main():

	success = False

	kill()

	if success:
		success = start()
		win = create()

	if success:
		success = connect(win)

	if success:
		success = refresh(win)

	if success:
		success = save(win)

	if success:
		success = publish(win)

	if success:
		success = close(win)

	if success:
		kill()
		return True
	else:
		return False

if __name__ == '__main__':

	success = main()

	print(success)