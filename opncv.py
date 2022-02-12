
#Importing Libraries

import math
import cv2
import numpy as np
import time
import HandTrackingModule as htm
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


#Setting up webCam using OpenCV
wCam, hCam= 720 , 480
cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)
pTime = 0

detector = htm.handDetector(detectionCon=0.7)

#Volume Control Library Usage
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

#Getting Volume Range using volume.GetVolumeRange() Method
volume.GetMute()
volume.GetMasterVolumeLevel()
volRange=volume.GetVolumeRange()
minVol=volRange[0]
maxVol=volRange[1]
vol=0
volBar=400
volPer=0



while(cap.isOpened()):
    success, img = cap.read()
    img = detector.findHands(img)
    lmlist= detector.findPosition(img, draw=False)

    #Assigning variables for Thumb and Index finger position
    if len(lmlist)!=0:
        print(lmlist[4],lmlist[8])
        x1, y1 = lmlist[4][1], lmlist[4][2]
        x2, y2 = lmlist[8][1], lmlist[8][2]

        cx,cy =(x1+x2)//2,(y1+y2)//2

        #Marking Thumb and Index finger using cv2.circle() and Drawing a line between them using cv2.line()
        cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), 15, (255, 0, 255), cv2.FILLED)
        cv2.line(img,(x1, y1),(x2, y2), (255, 0, 255),3)
        cv2.circle(img, (cx, cy), 15, (255, 0, 255), cv2.FILLED)
        length=math.hypot(x2-x1,y2-y1)
        print(length)


        #
        # # Hand Range 50-300
        # #volume Range -95  -  0
        #

        #Converting Length range into Volume range using numpy.interp()
        vol=np.interp(length,[30,200],[minVol,maxVol])

        # Changing System Volume using volume.SetMasterVolumeLevel() method
        volBar = np.interp(length, [30, 280], [400, 150])
        volPer = np.interp(length, [30, 280], [0, 100])
        print(int(length),vol)
        volume.SetMasterVolumeLevel(vol, None)

        if length<50:
            cv2.circle(img, (cx, cy), 15, (255, 0, 0), cv2.FILLED)

    #Drawing Volume Bar using cv2.rectangle() method
    cv2.rectangle(img,(50,150),(85,400),(255, 0, 0),3)
    cv2.rectangle(img, (50, int(volBar)), (85, 400), (255, 0, 0),cv2.FILLED)
    cv2.putText(img, f'{int(volPer)}%', (40, 450), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 2)


    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS: {int(fps)}', (40, 50), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 3)
    cv2.imshow("Img", img)

    #Displaying Output using cv2.imshow method
    if cv2.waitKey(1) & 0xFF == ord('q'):
         break
