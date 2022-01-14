import cv2
import numpy as np
import HandTrackingModule as Htm
import time
import math
import pyautogui
import win32gui, win32process, psutil
#for volume
#start
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

volRange=(volume.GetVolumeRange())

minVol = volRange[0]
maxVol = volRange[1]
vol = 0
volBar = 400
volPer = 0

#for volume END

wCam, hCam = 640, 480
'''Frame reduction :- Due to which we are easily able to detect our hand in certain frame
and the pointer moves up to corners easily.'''
frameR = 150
'''smoothening variable is used to make mouse move freely.'''
smoothening = 10

pTime = 0
plocX, plocY = 0, 0
clocX, clocY = 0, 0
prevloc, currloc = 0, 0


cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)
'''This handDetector function is imported from the HandTrackingModule.'''
'''And detector is the general variable which shows the handDetector function.'''
detector = Htm.handDetector(maxHands=1)
wScr, hScr = pyautogui.size()
print(wScr, hScr)

count = 1

while True:

    success, img = cap.read()
    img = detector.findHands(img)
    lmList, bbox = detector.findPosition(img)   # returns locations of hand landmarks
    # print(fingers)
    cv2.rectangle(img, (frameR, frameR), (wCam - frameR, hCam - frameR),
                      (255, 255, 0), 2)

    if len(lmList) != 0:    #  lmlist is list of co-ordinnates of hand landmarks
        fingers = detector.fingersUp()  #  Checks if fingers are open
        if fingers[1] == 1:
            
            if fingers[0] == 0 and fingers[2] == 0:
                # Getting co - ordinates of finger 0 and 1
                x1, y1 = lmList[8][1:]
                x2, y2 = lmList[4][1:]

                x3 = np.interp(x1, (frameR, wCam-frameR), (0, wScr))
                y3 = np.interp(y1, (frameR, hCam-frameR), (0, hScr))

                clocX = plocX + (x3 - plocX) / smoothening
                clocY = plocY + (y3 - plocY) / smoothening

                pyautogui.moveTo(wScr - clocX, clocY,_pause=False)
                cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
                plocX, plocY = clocX, clocY

            elif fingers[0] == 1 and fingers[2] == 0:
                length, img, lineInfo = detector.findDistance(8, 4, img)
                print(length)
                if length < 35:
                    cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                    pyautogui.click(button='right')  # right-click the mouse


            elif fingers[2] == 1 and fingers[3] == 0:
                length, img, lineInfo = detector.findDistance(8, 12, img)
                print(length)
                if length < 20:
                    cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                    pyautogui.click(interval=0.1)

            elif fingers[2] == 1 and fingers[3] == 1 and fingers[4]==1:
                x1, y1 = lmList[12][1:]

                if count == 1:
                    prevloc = currloc = y1
                    count = 0
                    diff = 0
                else:
                    diff = y1 - prevloc
                    prevloc = y1

                print(y1)
                pyautogui.scroll(diff * smoothening)
        else:
            count = 1
    try:
        #for volume
        if (psutil.Process(win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[-1])).name() == 'vlc.exe':
            if len(lmList) and fingers[2] == fingers[3] == fingers[4] == 1:  # not equals to 0 then it will execute

                x1, y1 = lmList[4][1], lmList[4][2]     # delacaring variable
                x2, y2 = lmList[8][1], lmList[8][2]
                cx, cy = (x1 + x2) // 2, (y1 + y2) // 2 # for center circle

                cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
                cv2.circle(img, (x2, y2), 15, (255, 0, 255), cv2.FILLED)
                cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
                cv2.circle(img, (cx, cy), 15, (255, 0, 255), cv2.FILLED)


                length = math.hypot(x2 - x1, y2 - y1)#length of line
                
                vol = np.interp(length, [50, 300], [minVol, maxVol])
                volBar = np.interp(length, [50, 300], [400, 150])
                volPer = np.interp(length, [50, 300], [0, 100])
                print(int(length), vol)
                volume.SetMasterVolumeLevel(vol, None)

                if length < 50:
                    cv2.circle(img, (cx, cy), 15, (0, 255, 0), cv2.FILLED)
                                        #width       #color
                cv2.rectangle(img, (50, 150), (85, 400), (255, 0, 0), 3)
                cv2.rectangle(img, (50, int(volBar)), (85, 400), (255, 0, 0), cv2.FILLED)
                cv2.putText(img, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX,
                        1, (255, 0, 0), 3)

        elif (psutil.Process(win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[-1])).name() == 'ApplicationFrameHost.exe' and len(lmList) and fingers[2] == fingers[3] == fingers[4] == 1:
            count = 1
            print("In upper elif")
            pyautogui.keyDown('ctrl')
            var = fingers[2] == fingers[3] == fingers[4] == 1
            while ((psutil.Process(win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[-1])).name() == 'ApplicationFrameHost.exe' and var):
                print("In while")
                success, img = cap.read()
                img = detector.findHands(img)
                lmList, bbox = detector.findPosition(img)
                if len(lmList): # not equals to 0 then it will execute
                    fingers = detector.fingersUp()
                  
                    print("In if")
                    x1, y1 = lmList[4][1], lmList[4][2] #delacaring variable
                    x2, y2 = lmList[8][1], lmList[8][2]
                    cx, cy = (x1 + x2) // 2, (y1 + y2) // 2#for center circle

                    cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
                    cv2.circle(img, (x2, y2), 15, (255, 0, 255), cv2.FILLED)
                    cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
                    cv2.circle(img, (cx, cy), 15, (255, 0, 255), cv2.FILLED)


                    length = math.hypot(x2 - x1, y2 - y1)#length of line
                    if count == 1:
                        prev = length
                        diff = 0
                        count = 0
                    else:
                        diff = length-prev
                        prev = length
                    print(diff)
                    pyautogui.scroll(int(diff)*5)
                    cv2.imshow("Image", img)
                    cv2.waitKey(1)
                    if length < 50:
                        cv2.circle(img, (cx, cy), 15, (0, 255, 0), cv2.FILLED)
                                            #width       #color
                    var = fingers[2] == fingers[3] == fingers[4] == 1
                else:
                    count = 1
            pyautogui.keyUp('ctrl')
    except Exception as e:
        print(e)
    
    cTime = time.time()
    fps = 1 / (cTime-pTime)
    pTime = cTime
    cv2.putText(img, str(int(fps)), (20, 50), cv2.FONT_HERSHEY_PLAIN, 3,
                (255, 0, 0), 3)

    cv2.imshow("Image", img)
    cv2.waitKey(1)