#! python3
import pyautogui as p
import time
import datetime
import subprocess
import sys
import os

# Globals

p.PAUSE = 1
# This is the global pause interval between pyautogui commands; 1 is equal to 1 second, 2 equals 2 seconds, and so on. Documentation says this can't go below 1; no reason to have it go below 1, anyway. 

l = 'left' 
# We're setting the letter 'l' equal to the string 'left' for the clk function. This way, when the clk function is used, the user doesn't have to type in 'left' with appropriate quotes every time.

insDate = datetime.date.today() 
# Set up today's date so we can save and open documents with today's date in their file name.

scriptHome = sys.path[0]
# Get where the script is so everything isn't so hardcoded

#-----------------------------------------------------------------------------------------------------------------------------------

# Greg came up with this function as a workaround to deal with pyautogui's current issue with clicking not always working. The bug itself is due to a deprecated click event in Windows (noted by Mr. Sweigart himself and others on the library's Github: https://github.com/asweigart/pyautogui/issues/23 )
def clk(x,y,b):
	try:
		p.click(x,y,button=b)
	except FileNotFoundError:
		pass
		
# Since we're going to be searching for files a lot in the file explorer, we need a function that jumps to the folder url bar, deletes the current file path, and appends the path supplied by the script into it		
def fe(path):
	time.sleep(2)
	p.press('f4')
	p.hotkey('ctrlleft', 'a', 'del')
	p.typewrite(path)
	p.press('enter')

# Clear up hardcoding	
def imgPath(img):
	path = os.path.join(scriptHome, 'elements\\' + img)
	return path

# Bundle together the image recognition commands: locate the image of the element, get the coordinates, click it. This definitely calls for a function as it's a script-critical repetiton.	
def locate(img):
	btn = p.locateOnScreen(img)
	btnX, btnY = p.center(btn)
	clk(btnX,btnY,l)	

#-----------------------------------------------------------------------------------------------------------------------------------
	
sqlOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft SQL Server\\120\\Tools\\Binn\\ManagementStudio\\Ssms.exe')
time.sleep(15)
'''
MS SQL 2k14 takes a bit to load up the first time on the office computer if the computer is coming up from a cold boot, so just in case of slow starts, delay any further actions for a bit. Not really necessary for a machine that's just going to sit in a corner but WKFC is pretty upside-down when it comes to this kind of stuff, so this could be shoved on someone's computer that shuts down at the end of the day/during the weekend.

I'm also not too worried about the program filepath being hardcoded since these are pretty traditional installation directories that aren't going to change much in the office environment unless Microsoft suddenly decides to change their entire filesystem overnight.
'''

# The chunk below is interacting with the file explorer to open the SQL query
locate(imgPath('connect.png')) # Surprisingly, sending "enter" is less reliable than actually clicking the "connect" button
p.hotkey('ctrlleft', 'o')
time.sleep(2)
p.typewrite('quote follow-up notes 5') 
p.press('enter') 
p.press('f5') # Run the query
# End interaction with file explorer

time.sleep(60)
# Right now the New Business QFU query is huge. It can take up to and sometimes above a minute to return all of the fields. Until we shrink that query to the data we explicitly NEED for QFUs, this lengthy delay will stay
locate(imgPath('selectAll.png')) # Want to get all of the column names as well for the spreadsheet
time.sleep(2)
p.hotkey('ctrlleft', 'shift', 'c')
p.hotkey('altleft', 'f4')
# Close MS SQL Management Studio

#-----------------------------------------------------------------------------------------------------------------------------------
# Open Excel and paste/save the results of the query
# Wait two seconds for Excel to fully load, then paste the query results and move the arrow key up to get rid of the highlight
xlOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft Office\\Office14\\EXCEL.exe')
time.sleep(2)
p.hotkey('ctrlleft', 'v')
p.press('up')

# 1. Switch to the data tab
# 2. Remove duplicate values based on Control Number
locate(imgPath('dataTab.png'))
locate(imgPath('removeDupes.png'))
locate(imgPath('unselect.png'))
locate(imgPath('ctrlNum.png'))
locate(imgPath('ok.png'))
locate(imgPath('okSmall.png'))

time.sleep(2)

# Save the spreadsheet and close Excel
p.hotkey('ctrlleft', 's')
fe('[filepath]') # However, I don't like how I've got this folder path hardcoded. It doesn't feel like as big of an issue if it's going to be primarily on one computer, which means I can just py2exe it on that same computer, but just from a proper programming standpoint I don't like it.
locate(imgPath('excelFileName.png'))
p.typewrite('Quote Follow Up for ' + insDate.strftime('%m-%d-%Y'))
p.press('enter')
p.hotkey('altleft', 'f4')

#------------------------------------------------------------------------------------------------------------------------------------
# Open Word and perform the mail merge
# Wait two seconds for Word to fully load and open the Mail Merge document
wordOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft Office\\Office14\\WINWORD.exe')
time.sleep(2)
p.hotkey('ctrlleft', 'o')
fe('[filepath]')# Again here, don't like the hardcoding.
locate(imgPath('wordOpenDoc.png'))
p.typewrite('Mail Merge') 

# This is all interacting with importing the Excel file we just made into Mail Merge
p.press('enter')
locate(imgPath('wordYes.png'))
locate(imgPath('mailings.png'))
locate(imgPath('selectTo.png'))
locate(imgPath('useList.png'))
fe('[filepath]')
locate(imgPath('listFileName.png'))	
p.typewrite('Quote Follow Up for ' + insDate.strftime('%m-%d-%Y')) 
p.press('enter')
p.press('enter')

time.sleep(4) # Ran into issue where large volume of QFUs would take a few seconds to load

# SEND THEM OUT (Navigate to the Finish & Merge button and send out all of the follow ups)
locate(imgPath('merge.png'))
locate(imgPath('send.png'))
locate(imgPath('mergeOk.png'))

time.sleep(25)
# Wait a bit for all the emails to actually be sent out; comes into play when the QFUs for the day break into the high double-digits
p.hotkey('ctrlleft', 's')
p.hotkey('altleft', 'f4')	