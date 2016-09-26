#! python3
import pyautogui as p
import time
import datetime
import subprocess
import sys
import os

p.PAUSE = 1
l = 'left'
insDate = datetime.date.today()
scriptHome = sys.path[0]

def clk(x,y,b):
	try:
		p.click(x,y,button=b)
	except FileNotFoundError:
		pass
		
def fe(path):
	time.sleep(2)
	p.press('f4')
	p.hotkey('ctrlleft', 'a', 'del')
	p.typewrite(path)
	p.press('enter')
	
def imgPath(img):
	path = os.path.join(scriptHome, 'elements\\' + img)
	return path

def locate(img):
	btn = p.locateOnScreen(img)
	btnX, btnY = p.center(btn)
	clk(btnX,btnY,l)	
	
sqlOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft SQL Server\\120\\Tools\\Binn\\ManagementStudio\\Ssms.exe')
time.sleep(25)
locate(imgPath('connect.png'))

time.sleep(2)
p.hotkey('ctrlleft', 'o')

time.sleep(2)
p.typewrite('quote follow-up notes 5') 
p.press('enter')
p.press('f5')

time.sleep(60)
locate(imgPath('selectAll.png'))

time.sleep(2)
p.hotkey('ctrlleft', 'shift', 'c')
p.hotkey('altleft', 'f4')

xlOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft Office\\Office14\\EXCEL.exe')
time.sleep(2)
p.hotkey('ctrlleft', 'v')
p.press('up')

locate(imgPath('dataTab.png'))
locate(imgPath('removeDupes.png'))
locate(imgPath('unselect.png'))
locate(imgPath('ctrlNum.png'))
locate(imgPath('ok.png'))
locate(imgPath('okSmall.png'))

time.sleep(2)

p.hotkey('ctrlleft', 's')
fe('[filepath]')
locate(imgPath('excelFileName.png'))
p.typewrite('Quote Follow Up for ' + insDate.strftime('%m-%d-%Y'))
p.press('enter')
p.hotkey('altleft', 'f4')


wordOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft Office\\Office14\\WINWORD.exe')
time.sleep(2)
p.hotkey('ctrlleft', 'o')
fe('[filepath]')
locate(imgPath('wordOpenDoc.png'))
p.typewrite('Mail Merge') 

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

time.sleep(4)

locate(imgPath('merge.png'))
locate(imgPath('send.png'))
locate(imgPath('mergeOk.png'))

time.sleep(25)
p.hotkey('ctrlleft', 's')
p.hotkey('altleft', 'f4')	
