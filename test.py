from sklearn.neighbors import KNeighborsClassifier
import cv2 as cv
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

# -----------------------------
# Text-to-Speech Function
# -----------------------------
def speak(strl):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(strl)

# -----------------------------
# Initialize Webcam & Face Detector
# -----------------------------
video = cv.VideoCapture(0)
# if not video.isOpened():
#     print("❌ Could not access webcam. Close other apps using it and retry.")
#     exit()

facedetect = cv.CascadeClassifier('data/haarcascade_frontalface_default.xml')
os.makedirs("Attendance", exist_ok=True)  # Ensure attendance folder exists

# -----------------------------
# Load Training Data
# -----------------------------
with open('data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

min_len = min(len(FACES), len(LABELS))
FACES = FACES[:min_len]
LABELS = LABELS[:min_len]

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# imgBackground = cv.imread("background.png")
COL_NAMES = ['NAME', 'TIME']

# -----------------------------
# Main Loop
# -----------------------------
while True:
    ret, frame = video.read()
    # if not ret or frame is None:
    #     print("⚠️ Failed to read frame from webcam. Exiting...")
    #     break

    gray = cv.cvtColor(frame, cv.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    ts = time.time()
    date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
    timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
    filename = f"Attendance/Attendance_{date}.csv"
    exist = os.path.isfile(filename)

    attendance_logged = False  # Track if we logged attendance this frame

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv.resize(crop_img, (10, 5)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        cv.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)
        cv.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv.putText(frame, str(output[0]), (x, y-15),
                   cv.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 1)

        attendance = [str(output[0]), str(timestamp)]

        # Write to CSV only once per frame
        if not attendance_logged:
            if exist:
                with open(filename, "a", newline="") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(attendance)
            else:
                with open(filename, "w", newline="") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    writer.writerow(attendance)
            attendance_logged = True
            speak("Attendance Taken...")

        # imgBackground[162:162 + 480, 55:55 + 640] = frame

    # Display Frame
    # cv.imshow('Attendance System', imgBackground)
    cv.imshow('Attendance System',frame)
    k = cv.waitKey(1)

    # Quit if 'q' is pressed
    if k == ord('q'):
        break

# -----------------------------
# Cleanup
# -----------------------------
video.release()
cv.destroyAllWindows()
