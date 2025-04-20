from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(str1)

# Load face detection model
video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

# Load training data
with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print("Original FACES shape:", np.shape(FACES))
print("Original LABELS length:", len(LABELS))

# Ensure lengths match
min_len = min(len(FACES), len(LABELS))
FACES = np.array(FACES[:min_len])
LABELS = LABELS[:min_len]

print("Trimmed FACES shape:", FACES.shape)
print("Trimmed LABELS length:", len(LABELS))

# Train classifier
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# Load background image
imgBackground = cv2.imread("background.png")

# CSV column headers
COL_NAMES = ['NAME', 'TIME']

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    attendance = []

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)

        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist = os.path.isfile(f"Attendance/Attendance_{date}.csv")

        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)

        attendance = [str(output[0]), str(timestamp)]

    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame", imgBackground)

    k = cv2.waitKey(1)

    if k == ord('o') and attendance:
        speak("Attendance Taken..")
        time.sleep(1)

        attendance_path = f"Attendance/Attendance_{date}.csv"
        with open(attendance_path, "a", newline='') as csvfile:
            writer = csv.writer(csvfile)
            if not exist:
                writer.writerow(COL_NAMES)
            writer.writerow(attendance)

    if k == ord('q'):
        break

video.release()
cv2.destroyAllWindows()
