import cv2
import time
import pyautogui
import threading
import HandTrackingModule as htm

# cam resolution
wCam, hCam = 640, 480

# how far hand needs to move (in px) before we count it as a swipe
swipeThreshold = 40

# min time between two slide changes so it doesnt double trigger
cooldown = 0.8

cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)

# using lower detection confidence (0.7) because 0.8 was dropping detection too often
detector = htm.handDetector(detectionCon=0.7, maxHands=1)

pTime = 0
lastActionTime = 0
startWristX = None       # where the hand was when swipe mode started
showAction = ""          # text to flash on screen after a swipe
showActionTime = 0

# these get updated by the background thread
currentSlide = 0
totalSlides = 0

# states: IDLE, SWIPE, POINTER, LOCKED, RESET
# needed a state machine because simple if/else was flickering too much
currentMode = "IDLE"

# without this pyautogui throws error if mouse hits screen corner
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0


def update_slide_info():
    """
    runs in a separate thread because calling PowerPoint COM API
    on every frame was freezing the video feed (50-100ms per call).
    now it just polls every 2 sec in the background - camera stays smooth.
    """
    global currentSlide, totalSlides
    import pythoncom
    # COM needs to be initialized separately for each thread
    pythoncom.CoInitialize()
    import win32com.client
    pptApp = win32com.client.Dispatch('PowerPoint.Application')
    while True:
        try:
            if pptApp.SlideShowWindows.Count > 0:
                ssw = pptApp.SlideShowWindows(1)
                currentSlide = ssw.View.Slide.SlideIndex
                totalSlides = pptApp.ActivePresentation.Slides.Count
        except:
            pass
        time.sleep(2)


# daemon=True so it dies when main program exits
slideThread = threading.Thread(target=update_slide_info, daemon=True)
slideThread.start()

while True:
    success, img = cap.read()
    if not success:
        break

    # flip so it feels like a mirror - swipe left = left on screen
    img = cv2.flip(img, 1)

    img = detector.findHands(img)
    lmList, bbox = detector.findPosition(img, draw=True)

    if len(lmList) != 0:
        fingers = detector.fingersUp()

        # FIST = reset. this is how you unlock after a swipe
        # or switch between pointer and swipe mode
        if fingers == [0, 0, 0, 0, 0]:
            currentMode = "RESET"
            startWristX = None

        # can only enter a new mode from RESET or IDLE
        # this prevents flickering - once in a mode, you're locked in
        elif currentMode in ["RESET", "IDLE"]:

            # one finger up = pointer mode, activates ppt laser
            if fingers == [0, 1, 0, 0, 0]:
                currentMode = "POINTER"
                # ctrl+l turns on powerpoint's built-in laser
                pyautogui.hotkey('ctrl', 'l')

            # two fingers (index+middle) = swipe mode
            # ignoring thumb because its detection is unreliable at angles
            elif fingers[1] == 1 and fingers[2] == 1 and fingers[3] == 0 and fingers[4] == 0:
                currentMode = "SWIPE"
                # save where wrist is right now - swipe = movement from this point
                startWristX = lmList[0][1]

        # --- POINTER: move cursor on screen so ppt laser follows ---
        if currentMode == "POINTER":
            # index fingertip = landmark 8
            ix, iy = lmList[8][1], lmList[8][2]

            # map camera pixel coords to actual screen coords
            screenW, screenH = pyautogui.size()
            mapX = int(ix * screenW / wCam)
            mapY = int(iy * screenH / hCam)
            pyautogui.moveTo(mapX, mapY)

            # red dot on camera feed so you can see where you're pointing
            cv2.circle(img, (ix, iy), 12, (0, 0, 255), cv2.FILLED)
            cv2.circle(img, (ix, iy), 18, (0, 0, 255), 2)

        # --- SWIPE: track wrist movement from start position ---
        elif currentMode == "SWIPE":
            # using wrist (landmark 0) not fingertip - much more stable
            wristX = lmList[0][1]

            if startWristX is not None:
                diff = wristX - startWristX
                currentTime = time.time()

                # moved left enough + cooldown passed = next slide
                if diff < -swipeThreshold and (currentTime - lastActionTime) > cooldown:
                    pyautogui.press('right')
                    lastActionTime = currentTime
                    showAction = ">> NEXT"
                    showActionTime = time.time()
                    # lock immediately so bringing hand back doesnt trigger prev
                    currentMode = "LOCKED"

                # moved right enough = prev slide
                elif diff > swipeThreshold and (currentTime - lastActionTime) > cooldown:
                    pyautogui.press('left')
                    lastActionTime = currentTime
                    showAction = "<< PREV"
                    showActionTime = time.time()
                    currentMode = "LOCKED"

        # --- LOCKED: do nothing. need fist to get out of this ---
        elif currentMode == "LOCKED":
            pass

    else:
        # no hand visible at all
        currentMode = "IDLE"
        startWristX = None

    # --- drawing the status bar at bottom ---
    cv2.rectangle(img, (0, hCam - 100), (wCam, hCam), (40, 40, 40), cv2.FILLED)

    # color code the current mode
    if currentMode == "POINTER":
        modeColor = (0, 0, 255)
    elif currentMode == "SWIPE":
        modeColor = (0, 255, 0)
    elif currentMode == "LOCKED":
        modeColor = (0, 255, 255)
    elif currentMode == "RESET":
        modeColor = (255, 255, 0)
    else:
        modeColor = (200, 200, 200)

    cv2.putText(img, f'Mode: {currentMode}', (10, hCam - 70),
                cv2.FONT_HERSHEY_PLAIN, 1.5, modeColor, 2)

    # slide number comes from background thread - never stalls the feed
    if totalSlides > 0:
        cv2.putText(img, f'Slide: {currentSlide}/{totalSlides}', (10, hCam - 35),
                    cv2.FONT_HERSHEY_PLAIN, 1.5, (255, 255, 255), 2)
    else:
        cv2.putText(img, 'No slideshow running', (10, hCam - 35),
                    cv2.FONT_HERSHEY_PLAIN, 1.5, (100, 100, 255), 2)

    if currentMode == "LOCKED":
        cv2.putText(img, "Make fist to reset", (300, hCam - 35),
                    cv2.FONT_HERSHEY_PLAIN, 1.2, (0, 255, 255), 2)

    # flash the action text for 1 sec after a swipe
    if showAction and (time.time() - showActionTime) < 1.0:
        actionColor = (0, 255, 0) if "NEXT" in showAction else (0, 100, 255)
        cv2.putText(img, showAction, (200, 80), cv2.FONT_HERSHEY_PLAIN, 3, actionColor, 3)
    else:
        showAction = ""

    # fps
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS: {int(fps)}', (500, hCam - 70), cv2.FONT_HERSHEY_PLAIN, 1.5, (255, 0, 0), 2)

    cv2.imshow("Presentation Controller", img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
