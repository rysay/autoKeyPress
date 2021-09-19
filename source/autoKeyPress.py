# coding: utf-8
import pyautogui
import time
import subprocess
import win32gui
import win32con
import array


def main():
    """メイン処理2
    """
    pyautogui.FAILSAFE = False
    targetWindowName = 'Bluestacks'

    i = 0
    time.sleep(1)
    while i < 1000:
        for _ in range(5):
            #print(str(i) + ':1times')
            executeTargetWindow(targetWindowName)
            time.sleep(10)
        time.sleep(100)
        i += 1


def executeTargetWindow(targetWindowName):
    """対象のアプリケーションウィンドウを一時アクティブにしてメソッドを実行。
    Parameters
    ----------
    targetWindowName : String
        ウィンドウハンドルのタイトル
    """
    activeWindowTitle = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    activeWindowClass = win32gui.GetClassName(win32gui.GetForegroundWindow())

    targetWindow = win32gui.FindWindow(None, targetWindowName)
    if targetWindow == 0:
        cmd = 'start C:\\Program Files\\BlueStacks_bgp64_hyperv\\Bluestacks.exe'
        subprocess.Popen(cmd, shell=True)
        return

    windowActive(targetWindow)
    win32gui.SetForegroundWindow(win32gui.FindWindow(activeWindowClass, activeWindowTitle))


def windowActive(handle):
    """指定したウィンドウをアクティブにする(ウィンドウ左上クリック)
    """
    # 例外が出るwin32gui.SetForegroundWindow(habdle)の代わり
    # win32con.HWND_TOPMOST:-1(最前列)
    # win32con.HWND_NOTOPMOST:-2(最前列の解除)
    # SWP_SHOWWINDOW:0x40
    # win32con.SWP_NOMOVE:0x2
    # win32con.SWP_NOSIZE:0x1
    win32gui.SetWindowPos(handle, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOMOVE | win32con.SWP_SHOWWINDOW)



    # マウスクリックでウィンドウアクティブ
    win_x1, win_y1, win_x2, win_y2 = win32gui.GetWindowRect(handle)
    mousePos = pyautogui.position()
    pyautogui.leftClick(win_x1, win_y1)
    pyautogui.moveTo(mousePos)

    pyautogui.press('enter')


if __name__ == "__main__":
    main()
