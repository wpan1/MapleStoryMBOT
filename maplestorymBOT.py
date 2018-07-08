# TODO: PIXEL LOCATIONS FOR 1920x1040 RES
# TODO: Refactor GetWindowRect to complete only once
# TODO: Allow for lowered complete button
# TODO: Allow for complete and avaialable quest dialog buttons
# TODO: Change magic numbers for getcolors(), should work with exact
# TODO: Raw input for PID

import win32con
import win32gui
import win32ui
import win32api
import win32process
import threading
import time
import array

from PIL import Image, ImageOps

class bot(threading.Thread):

    def __init__ (self, hwnd):
        threading.Thread.__init__(self)
        self.hwnd = hwnd

        # TODO: PIXEL LOCATIONS FOR 1920x1040 RES
        # NEED TO CHANGE TO FRACTIONS FOR ANY RES

        # Pixel locations for start of quest dialog
        self.questWidthStart = 119.0
        self.questHeightStart = 342.0
        self.questWidthEnd = 443.0
        self.questHeightEnd = 470.0

        # Pixel locations for complete quest dialog
        self.compWidthStart = 1450.0
        self.compHeightStart = 868.0
        self.compWidthEnd = 1761.0
        self.compHeightEnd = 923.0

        # Pixel locations for claim quest dialog
        self.claimWidthStart = 782.0
        self.claimHeightStart = 856.0
        self.claimWidthEnd = 1132.0
        self.claimHeightEnd = 922.0

        # Pixel locations for skip prompt
        self.skipWidthStart = 222.0
        self.skipHeightStart = 105.0
        self.skipWidthEnd = 224.0
        self.skipHeightEnd = 128.0

        self.prevAction = None
        self.sleepTime = 5

    def run(self):
        while(1):
            self.tick()
            time.sleep(self.sleepTime)

    def tick(self):
        # TODO: Refactor GetWindowRect to complete only once

        # Take screenshot
        windowcor = win32gui.GetWindowRect(self.hwnd)
        w=windowcor[2]-windowcor[0]
        h=windowcor[3]-windowcor[1]
        wDC = win32gui.GetWindowDC(self.hwnd)
        dcObj=win32ui.CreateDCFromHandle(wDC)
        cDC=dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, w, h)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0,0),(w, h) , dcObj, (0,0), win32con.SRCCOPY)

        bmpinfo = dataBitMap.GetInfo()
        bmpstr = dataBitMap.GetBitmapBits(True)

        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)

        width, height = im.size

        # Capture accept quest cropped image
        acceptQuest = im.crop((
            int(width*self.compWidthStart/w),
            int(height*self.compHeightStart/h),
            int(width*self.compWidthEnd/w),
            int(height*self.compHeightEnd/h)
            ))
        acceptQuest = ImageOps.grayscale(acceptQuest)

        # Capture claim quest cropped image
        claimQuest = im.crop((
            int(width*self.claimWidthStart/w),
            int(height*self.claimHeightStart/h),
            int(width*self.claimWidthEnd/w),
            int(height*self.claimHeightEnd/h)
            ))
        claimQuest = ImageOps.grayscale(claimQuest)

        # Capture skip quest cropped image
        # No gray scale since image is < 256 pixels
        skipQuest = im.crop((
            int(width*self.skipWidthStart/w),
            int(height*self.skipHeightStart/h),
            int(width*self.skipWidthEnd/w),
            int(height*self.skipHeightEnd/h)
            ))

        # TODO: Allow for lowered complete button
        # TODO: Allow for complete and avaialable quest dialog buttons
        # TODO: Change magic numbers for getcolors(), should work with exact
        # Skip Quest
        if sorted(skipQuest.getcolors())[-1][0] < 18:
            print "Skip"
            skipClickX = int((self.skipWidthStart+windowcor[0]+self.skipWidthEnd+windowcor[0])/2)
            skipClickY = int((self.skipHeightStart+windowcor[1]+self.skipHeightEnd+windowcor[1])/2)
            print skipClickX
            print skipClickY
            self.click((skipClickX,skipClickY))
            self.prevAction = "Skip"
            self.sleepTime = 1
        # Accept Quest
        elif sorted(acceptQuest.getcolors())[-1][0] >= 15000:
            print "Accept"
            completeClickX = int((self.compWidthStart+windowcor[0]+self.compWidthEnd+windowcor[0])/2)
            completeClickY = int((self.compHeightStart+windowcor[1]+self.compHeightEnd+windowcor[1])/2)
            print completeClickX
            print completeClickY
            self.click((completeClickX,completeClickY))
            self.prevAction = "Accept"
            self.sleepTime = 1
        # Claim Quest
        elif sorted(claimQuest.getcolors())[-1][0] >= 15000:
            print "Claim"
            claimClickX = int((self.claimWidthStart+windowcor[0]+self.claimWidthEnd+windowcor[0])/2)
            claimClickY = int((self.claimHeightStart+windowcor[1]+self.claimHeightEnd+windowcor[1])/2)
            print claimClickX
            print claimClickY
            self.click((claimClickX,claimClickY))
            self.prevAction = "Claim"
            self.sleepTime = 3
        # Start Quest (Quest automatically started after accept, quest canceled if double tapped)
        elif self.prevAction != "Accept" and self.prevAction != "Quest":
            print "Quest"
            questClickX = int((self.questWidthStart+windowcor[0]+self.questWidthEnd+windowcor[0])/2)
            questClickY = int((self.questHeightStart+windowcor[1]+self.questHeightEnd+windowcor[1])/2)
            print questClickX
            print questClickY
            self.click((questClickX,questClickY))
            self.prevAction = "Quest"
            self.sleepTime = 1
        return True

    # Move mouse and click on position
    def click(self, pos):
        startPos = win32api.GetCursorPos()
        win32api.SetCursorPos(pos)
        win32api.mouse_event(
            win32con.MOUSEEVENTF_LEFTDOWN,0,0)
        win32api.mouse_event(
            win32con.MOUSEEVENTF_LEFTUP,0,0)
        win32api.SetCursorPos(startPos)


# Grab HWNDs for all windows and compares with PID
def get_hwnds_for_pid (pid):
    def callback (hwnd, hwnds):
        if win32gui.IsWindowVisible (hwnd) and win32gui.IsWindowEnabled (hwnd):
            _, found_pid = win32process.GetWindowThreadProcessId (hwnd)
            if found_pid == pid:
                hwnds.append (hwnd)
        return True
    hwnds = []
    win32gui.EnumWindows (callback, hwnds)
    return hwnds

# TODO: Raw input for PID
if __name__ == '__main__':
    for hwnd in get_hwnds_for_pid (5564):
        thread = bot(hwnd)
        thread.start()
        thread.join()

