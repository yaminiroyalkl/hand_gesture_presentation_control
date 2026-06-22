import cv2
import time
import pyautogui
import threading
import HandTrackingModule as htm

# Settings
wCam, hCam = 640, 480
swipeThreshold = 40
cooldown = 0.8

cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)

detector = htm.handDetector(detectionCon=0.7, maxHands=1)
pTime = 0
lastActionTime = 0
startWristX = None
showAction = ""
showActionTime = 0

# Slide info (updated in background thread)
currentSlide = 0
totalSlides = 0

# State machine
currentMode = "IDLE"

# Disable pyautogui failsafe
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0


def update_slide_info():
    """Background thread - polls PowerPoint for slide number every 2 seconds"""
    global currentSlide, totalSlides
    import pythoncom
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


# Start background thread for slide info
slideThread = threading.Thread(target=update_slide_info, daemon=True)
slideThread.start()

while True:
    success, img = cap.read()
    if not success:
        break

    img = cv2.flip(img, 1)
    img = detector.findHands(img)
    lmList, bbox = detector.findPosition(img, draw=True)

    if len(lmList) != 0:
        fingers = detector.fingersUp()

        # === RESET: fist (all fingers down) ===
        if fingers == [0, 0, 0, 0, 0]:
            currentMode = "RESET"
            startWristX = None

        # === From RESET or IDLE, you can enter a mode ===
        elif currentMode in ["RESET", "IDLE"]:

            # Enter POINTER mode - index only [0,1,0,0,0]
            if fingers == [0, 1, 0, 0, 0]:
                currentMode = "POINTER"
                pyautogui.hotkey('ctrl', 'l')

            # Enter SWIPE mode - index + middle [X,1,1,0,0]
            elif fingers[1] == 1 and fingers[2] == 1 and fingers[3] == 0 and fingers[4] == 0:
                currentMode = "SWIPE"
                startWristX = lmList[0][1]

        # === POINTER MODE ===
        if currentMode == "POINTER":
            ix, iy = lmList[8][1], lmList[8][2]
            screenW, screenH = pyautogui.size()
            mapX = int(ix * screenW / wCam)
            mapY = int(iy * screenH / hCam)
            pyautogui.moveTo(mapX, mapY)
            cv2.circle(img, (ix, iy), 12, (0, 0, 255), cv2.FILLED)
            cv2.circle(img, (ix, iy), 18, (0, 0, 255), 2)

        # === SWIPE MODE ===
        elif currentMode == "SWIPE":
            wristX = lmList[0][1]

            if startWristX is not None:
                diff = wristX - startWristX
                currentTime = time.time()

                # Swipe LEFT → Next slide
                if diff < -swipeThreshold and (currentTime - lastActionTime) > cooldown:
                    pyautogui.press('right')
                    lastActionTime = currentTime
                    showAction = ">> NEXT"
                    showActionTime = time.time()
                    currentMode = "LOCKED"

                # Swipe RIGHT → Previous slide
                elif diff > swipeThreshold and (currentTime - lastActionTime) > cooldown:
                    pyautogui.press('left')
                    lastActionTime = currentTime
                    showAction = "<< PREV"
                    showActionTime = time.time()
                    currentMode = "LOCKED"

        # === LOCKED ===
        elif currentMode == "LOCKED":
            pass

    else:
        currentMode = "IDLE"
        startWristX = None

    # --- UI OVERLAY ---
    cv2.rectangle(img, (0, hCam - 100), (wCam, hCam), (40, 40, 40), cv2.FILLED)

    # Mode
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

    # Slide number (from background thread - never blocks)
    if totalSlides > 0:
        cv2.putText(img, f'Slide: {currentSlide}/{totalSlides}', (10, hCam - 35),
                    cv2.FONT_HERSHEY_PLAIN, 1.5, (255, 255, 255), 2)
    else:
        cv2.putText(img, 'No slideshow running', (10, hCam - 35),
                    cv2.FONT_HERSHEY_PLAIN, 1.5, (100, 100, 255), 2)

    if currentMode == "LOCKED":
        cv2.putText(img, "Make fist to reset", (300, hCam - 35),
                    cv2.FONT_HERSHEY_PLAIN, 1.2, (0, 255, 255), 2)

    if showAction and (time.time() - showActionTime) < 1.0:
        actionColor = (0, 255, 0) if "NEXT" in showAction else (0, 100, 255)
        cv2.putText(img, showAction, (200, 80), cv2.FONT_HERSHEY_PLAIN, 3, actionColor, 3)
    else:
        showAction = ""

    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS: {int(fps)}', (500, hCam - 70), cv2.FONT_HERSHEY_PLAIN, 1.5, (255, 0, 0), 2)

    cv2.imshow("Presentation Controller", img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
