import win32con
import win32gui
import lackey as lk
from lackey import *
import pyautogui
import win32api

ScreenObject = lk.Screen()

def bringwindowToFront(windowName):
    App.focus(windowName)
    ScreenObject.wait(1)

def bringwindowToFront_2(my_window):

    def get_window_handle(partial_window_name):

        # https://www.blog.pythonlibrary.org/2014/10/20/pywin32-how-to-bring-a-window-to-front/

        def window_enumeration_handler(hwnd, windows):
            windows.append((hwnd, win32gui.GetWindowText(hwnd)))

        windows = []
        win32gui.EnumWindows(window_enumeration_handler, windows)

        for i in windows:
            if partial_window_name.lower() in i[1].lower():
                return i
                break

        print('window not found!')
        return None

    # https://stackoverflow.com/questions/6312627/windows-7-how-to-bring-a-window-to-the-front-no-matter-what-other-window-has-fo

    def bring_window_to_foreground(HWND):
        win32gui.ShowWindow(HWND, win32con.SW_RESTORE)
        win32gui.SetWindowPos(HWND, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(HWND, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(HWND, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_SHOWWINDOW + win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
        win32gui.ShowWindow(HWND, win32con.SW_MAXIMIZE)
        # win32gui.ShowWindow(HWND, win32con.FOCUS)
        ScreenObject.wait(2)
        ScreenObject.click()


    hwnd = get_window_handle(my_window)

    if hwnd is not None:
        bring_window_to_foreground(hwnd[0])

# bringwindowToFront('Inbox - ola.abbas@vodafone.com')
# bringwindowToFront('google translate')
# bringwindowToFront('Downloads')
# App.focus('Inbox - ola.abbas@vodafone.com')

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

def bringwindowToFront_3(windowName):
    if __name__ == "__main__":
        results = []
        top_windows = []
        win32gui.EnumWindows(windowEnumerationHandler, top_windows)
        for i in top_windows:
            if windowName in i[1]:
                print(i)
                win32gui.ShowWindow(i[0], 5)
                win32gui.ShowWindow(i[0], 5)
                win32gui.SetForegroundWindow(i[0])
                break

# def setLanguageToEN():
#     # pyautogui.hotkey('ctrl' , 'shift' , '1')
#     win32api.LoadKeyboardLayout('00000809' , 1)

# setLanguageToEN()