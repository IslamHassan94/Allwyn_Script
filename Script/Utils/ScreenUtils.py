import lackey
import win32gui
from lackey import *
import yaml
import pyperclip3 as pc
import pyautogui
import pygetwindow as gw
import os as oos
import platform
import pywinauto

config = yaml.safe_load(open("../../config.yml"))
img_Path = config['imgs_Path']['clarify']
ScreenObject = lackey.Screen()


def ScreenObjectsettings():
    ScreenObject.setAutoWaitTimeout(60)
    # ScreenObject.wait(2)
    # if SendMessage(GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, 0, 0x4090409) != 0:
    #     print("olaaa")
    #     win32api.LoadKeyboardLayout('00000409', 1)


def openNotPad():
    ScreenObject.click(img_Path + "windosImg.png")
    ScreenObject.wait(2)
    ScreenObject.type('notepad')


def getClipboardText():
    clipBoardText = ScreenObject.getClipboard()
    return clipBoardText


def setClipboardText(Text):
    ScreenObject.wait(1)
    # pc.copy(copiedText)
    App.setClipboard(Text)


def captureEntireScreen():
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(config['Screenshots'])


def copy():
    ScreenObject.wait(1)
    ScreenObject.type("c", Key.CTRL)
    ScreenObject.wait(1)
    ScreenObject.type("c", Key.CTRL)


def pasteWithoutValue():
    ScreenObject.wait(1)
    ScreenObject.type("v", Key.CTRL)


def pasteText(Text):
    ScreenObject.wait(1)
    ScreenObject.paste(Text)


def selectAll():
    ScreenObject.wait(1)
    ScreenObject.type("a", Key.CTRL)


def existsAndWaitBeforeClick(pattern, wait, waitBeforeClick):
    if ScreenObject.exists(pattern, wait) is not None:
        ScreenObject.wait(waitBeforeClick)
        ScreenObject.click(pattern)

# windowsImg = img_Path + "windosImg.png"
# getWindowByTitle("outlook", windowNum=0)
# existsAndWaitBeforeClick(windowsImg , 30 , 4)

# ScreenObject.find(windowsImg).click()
# ScreenObjectsettings()
# openNotPad()
# # print(ScreenObject.getClipboard())
# setClipboardText("abcbbbbbbb")
# print(getClipboardText())

# ScreenObject.doubleClick()
# copy()
#
# openNotPad()
# ScreenObject.wait(2)
# ScreenObject.type(Key.ENTER)
# ScreenObject.wait(2)
# pasteWithoutValue()
# pasteText("PPPPaste")
# ScreenObject.wait(2)
# ScreenObject.type("hhhhhhh")
# ScreenObject.wait(2)
# selectAll()
