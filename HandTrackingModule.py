import cv2
import mediapipe as mp
import time
import math

class handDetector():
    def __init__(self, mode=False, maxHands=2, detectionCon=0.5, trackCon=0.5):
        self.mode = mode
        self.maxHands = maxHands
        self.detectionCon = detectionCon
        self.trackCon = trackCon

        # New mediapipe API setup
        BaseOptions = mp.tasks.BaseOptions
        self.HandLandmarker = mp.tasks.vision.HandLandmarker
        HandLandmarkerOptions = mp.tasks.vision.HandLandmarkerOptions
        VisionRunningMode = mp.tasks.vision.RunningMode

        # Choose running mode based on 'mode' parameter
        # mode=False means VIDEO (continuous frames), mode=True means IMAGE (single photo)
        running_mode = VisionRunningMode.IMAGE if self.mode else VisionRunningMode.VIDEO

        options = HandLandmarkerOptions(
            base_options=BaseOptions(model_asset_path='hand_landmarker.task'),
            running_mode=running_mode,
            num_hands=self.maxHands,
            min_hand_detection_confidence=self.detectionCon,
            min_tracking_confidence=self.trackCon
        )

        self.handLandmarker = self.HandLandmarker.create_from_options(options)
        self.results = None
        self.tipIds = [4, 8, 12, 16, 20]  # Fingertip landmark IDs
        self.lmList = []

        # Hand connections for drawing skeleton
        self.HAND_CONNECTIONS = [
            (0, 1), (1, 2), (2, 3), (3, 4),        # Thumb
            (0, 5), (5, 6), (6, 7), (7, 8),        # Index
            (0, 9), (9, 10), (10, 11), (11, 12),   # Middle
            (0, 13), (13, 14), (14, 15), (15, 16), # Ring
            (0, 17), (17, 18), (18, 19), (19, 20), # Pinky
            (5, 9), (9, 13), (13, 17)              # Palm
        ]

    def findHands(self, img, draw=True):
        imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        mp_image = mp.Image(image_format=mp.ImageFormat.SRGB, data=imgRGB)

        # Detect hands
        timestamp_ms = int(time.time() * 1000)
        self.results = self.handLandmarker.detect_for_video(mp_image, timestamp_ms)

        if self.results.hand_landmarks:
            for hand_landmarks in self.results.hand_landmarks:
                if draw:
                    h, w, c = img.shape
                    points = []
                    for lm in hand_landmarks:
                        cx, cy = int(lm.x * w), int(lm.y * h)
                        points.append((cx, cy))
                    # cv2.circle(img, points[4], 15, (255, 0, 0), cv2.FILLED)  # Blue - Thumb tip
                    # Draw lines (replaces old mpDraw.draw_landmarks)
                    for connection in self.HAND_CONNECTIONS:
                        start = points[connection[0]]
                        end = points[connection[1]]
                        cv2.line(img, start, end, (0, 255, 0), 2)

                    # Draw dots
                    for cx, cy in points:
                        cv2.circle(img, (cx, cy), 4, (0, 0, 255), cv2.FILLED)

        return img

    def findPosition(self, img, handNo=0, draw=True):
        xList = []
        yList = []
        bbox = []
        lmList = []
        if self.results and self.results.hand_landmarks:
            if handNo < len(self.results.hand_landmarks):
                myHand = self.results.hand_landmarks[handNo]
                h, w, c = img.shape
                for id, lm in enumerate(myHand):
                    cx, cy = int(lm.x * w), int(lm.y * h)
                    xList.append(cx)
                    yList.append(cy)
                    lmList.append([id, cx, cy])
                    if draw:
                        cv2.circle(img, (cx, cy), 5, (255, 0, 255), cv2.FILLED)

                xmin, xmax = min(xList), max(xList)
                ymin , ymax = min(yList), max(yList)

                bbox = xmin,ymin,xmax,ymax

                if draw:
                    cv2.rectangle(img, (bbox[0]-20, bbox[1]-20), (bbox[2]+20 ,bbox[3]+20), (0, 255, 0), 2)

        self.lmList = lmList
        return lmList, bbox

    def fingersUp(self):
        fingers = []
        # Thumb - checks x-axis (horizontal movement)
        if self.lmList[self.tipIds[0]][1] < self.lmList[self.tipIds[0] - 1][1]:
            fingers.append(1)
        else:
            fingers.append(0)
        # 4 Fingers - checks y-axis (vertical movement)
        for id in range(1, 5):
            if self.lmList[self.tipIds[id]][2] < self.lmList[self.tipIds[id] - 2][2]:
                fingers.append(1)
            else:
                fingers.append(0)
        return fingers

    def findDistance(self,p1,p2,img,draw = True):
        x1, y1 = self.lmList[p1][1], self.lmList[p1][2]
        x2, y2 = self.lmList[p2][1], self.lmList[p2][2]

        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
        if draw:
            cv2.circle(img, (x1, y1), 7, (255, 0, 255), cv2.FILLED)
            cv2.circle(img, (x2, y2), 7, (255, 0, 255), cv2.FILLED)

            cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 2)
            cv2.circle(img, (cx, cy), 7, (255, 0, 255), cv2.FILLED)

        length = math.hypot(x2 - x1, y2 - y1)
        return length,img,[x1,y1,x2,y2,cx,cy]


def main():
    pTime = 0
    cap = cv2.VideoCapture(0)
    detector = handDetector()

    while True:
        success, img = cap.read()
        if not success:
            break

        img = detector.findHands(img)
        lmList = detector.findPosition(img, draw=False)
        if len(lmList) != 0:
            print(lmList[4])  # Thumb tip position

        cTime = time.time()
        fps = 1 / (cTime - pTime)
        pTime = cTime

        cv2.putText(img, str(int(fps)), (10, 70), cv2.FONT_HERSHEY_PLAIN, 3,
                    (255, 0, 255), 3)

        cv2.imshow("Image", img)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
