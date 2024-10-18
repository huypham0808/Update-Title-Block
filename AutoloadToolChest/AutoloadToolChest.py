import pyautogui as pag
import time


time.sleep(2)

pag.keyDown('ctrl')
pag.press('a')
pag.keyUp('ctrl')
pag.mouseDown()

time.sleep(1)
posFile = pag.position(1382, 423)
posBluebeam = pag.position(560, 549)

time.sleep(1)
# pag.moveTo(posFile)
pag.dragTo(posBluebeam, duration=1)


for _ in range(42):
    pos1 = pag.position(560, 549)
    pag.click(pos1)
    pos2 = pag.position(495, 548)
    pag.click(pos2)

# **time.sleep(10)
# image1 = 'test1.png'
# loc1 = pag.locateOnScreen(r'D:\4.PERSONAL\5.CODING\PYTHON BASIC\AutoloadToolChest\YesButton.png')
# pag.click(loc1) 1382 423
