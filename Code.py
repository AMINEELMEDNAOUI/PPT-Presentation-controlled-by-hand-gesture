import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import numpy as np

Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open("C:\\Users\\user dell\\Desktop\\Amine\\vision par ordinateur\\PPT-Presentation-controlled-by-hand-gesture\\Amine.pptx")

print(Presentation.Name)
Presentation.SlideShowSettings.Run()

# Parameters
width, height = 900, 720
gestureThreshold = 300  # Threshold to detect hand near face
gestureZone = height // 2  # Middle of the screen to differentiate gestures

# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Hand Detector
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# Variables
buttonPressed = False
counter = 0
imgNumber = 20
delay = 30
annotations = [[]]
annotationNumber = -1
annotationStart = False

while True:
    # Get image frame
    success, img = cap.read()
    hands, img = detectorHand.findHands(img)  # Detect hands and draw landmarks

    if hands and not buttonPressed:  # If hand is detected
        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]  # List of landmark points
        fingers = detectorHand.fingersUp(hand)  # Which fingers are up?

        # Check if the hand is in the gesture zone
        if cy <= gestureThreshold:
            # Gesture 1 (Previous Slide): Thumb up only
            if fingers == [1, 0, 0, 0, 0]:  # Thumb up only
                print("Previous")
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.Previous()
                    imgNumber += 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False
            
            # Gesture 2 (Next Slide): Whole hand pointing up
            elif fingers == [1, 1, 1, 1, 1]:  # All fingers up (pointing up)
                print("Next")
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.Next()
                    imgNumber -= 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False

    else:
        annotationStart = False

    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    # Display annotations or any other visual feedback
    for annotation in annotations:
        for i in range(1, len(annotation)):
            cv2.line(img, annotation[i - 1], annotation[i], (0, 0, 200), 12)

    # Show the image frame
    cv2.imshow("Image", img)

    # Key press to exit
    key = cv2.waitKey(1)
    if key == ord('q'):
        break

cv2.destroyAllWindows()
