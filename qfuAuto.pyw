#! python3
import pyautogui as p
import time, datetime, subprocess, sys, os

p.PAUSE = 1
l = 'left'
insDate = datetime.date.today()
scriptHome = sys.path[0]
relPath = os.path.expanduser('~')

def clk(x,y,b):
	try:
		p.click(x,y,button=b)
	except FileNotFoundError:
		pass
		
def explorer(path):
	time.sleep(2)
	p.press('f4')
	p.hotkey('ctrlleft', 'a', 'del')
	p.typewrite(path)
	p.press('enter')
	
def imgPath(img):
	path = os.path.join(scriptHome, 'elements\\' + img)
	return path

def get(img):
	btn = p.locateOnScreen(img)
	btnX, btnY = p.center(btn)
	clk(btnX,btnY,l)	
	
sqlOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft SQL Server\\120\\Tools\\Binn\\ManagementStudio\\Ssms.exe')
time.sleep(25)
get(imgPath('connect.png'))

time.sleep(2)
p.hotkey('ctrlleft', 'o')

time.sleep(2)
p.typewrite('new and renewal qfu.sql') 
p.press('enter')
p.press('f5')

time.sleep(80)
get(imgPath('selectAll.png'))

time.sleep(2)
p.hotkey('ctrlleft', 'shift', 'c')
p.hotkey('altleft', 'f4')

xlOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft Office\\Office14\\EXCEL.exe')
time.sleep(8)
p.hotkey('ctrlleft', 'v')
p.press('up')

get(imgPath('dataTab.png'))
get(imgPath('removeDupes.png'))
#get(imgPath('dataheaders.png'))
get(imgPath('unselect.png'))
get(imgPath('ctrlNum.png'))
get(imgPath('ok.png'))
get(imgPath('okSmall.png'))

time.sleep(2)

p.hotkey('ctrlleft', 's')
explorer(relPath + '\\Documents\\Quote Follow Ups Archive')
get(imgPath('excelFileName.png'))
p.typewrite('Quote Follow Up for ' + insDate.strftime('%m-%d-%Y'))
p.press('enter')
p.hotkey('altleft', 'f4')

wordOpen = subprocess.Popen('C:\\Program Files (x86)\\Microsoft Office\\Office14\\WINWORD.exe')
time.sleep(8)
p.hotkey('ctrlleft', 'o')
explorer(relPath + '\\Documents')
get(imgPath('wordOpenDoc.png'))
p.typewrite('Mail Merge.docm') 

p.press('enter')
get(imgPath('wordYes.png'))
get(imgPath('mailings.png'))
get(imgPath('selectTo.png'))
get(imgPath('useList.png'))
explorer(relPath + '\\Documents\\Quote Follow Ups Archive')
get(imgPath('listFileName.png'))	
p.typewrite('Quote Follow Up for ' + insDate.strftime('%m-%d-%Y') + ".xlsx") 
p.press('enter')
p.press('enter')

time.sleep(4)

get(imgPath('merge.png'))
get(imgPath('send.png'))
get(imgPath('mergeOk.png'))

time.sleep(35)
p.hotkey('ctrlleft', 's')
p.hotkey('altleft', 'f4')
