from sklearn.neighbors import KNeighborsClassifier

import cv2
import pickle
import csv
import time
from datetime import datetime
import numpy as np
import os

from win32com.client import Dispatch

def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)
  
video = cv2.VideoCapture(0)
faces_detect = cv2.CascadeClassifier("haarcascade_frontalface_default.xml")
with open('data/names.pkl',"rb") as w:
    LABELS=pickle.load(w)
with open('data/faces_data.pkl',"rb") as f:
    FACES=pickle.load(f)   

print('Shape of Faces matrix --> ', FACES.shape)
# assert len(FACES) == len(LABELS), "Number of samples and labels must be the same"
num_labels = len(LABELS)  # Get the number of labels
filtered_faces = FACES[:num_labels, :]  # Filter FACES to match the number of labels
 

knn = KNeighborsClassifier(n_neighbors=5)
    
knn.fit(filtered_faces, LABELS)

imgBackground=cv2.imread("dg1.png")

COL_NAME=['NAME','TIME']

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)  
    faces = faces_detect.detectMultiScale(gray, scaleFactor=1.3, minNeighbors=5)  

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w , :]
        resized_img=cv2.resize(crop_img,(50,50)).flatten().reshape(1,-1)
        output=knn.predict(resized_img)
        ts=time.time()
        date=datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp=datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist=os.path.isfile("Attendance/Attendance_"+date+".csv")
        cv2.rectangle(frame,(x,y),(x+w,y+h),(0,0,255),1)
        cv2.rectangle(frame,(x,y),(x+w,y+h),(50,50,255),2)
        cv2.rectangle(frame,(x,y-40),(x+w,y),(50,50,255),-1)
        cv2.putText(frame,str(output[0]),(x,y-15),cv2.FONT_HERSHEY_COMPLEX,1,(255,255,255),1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 1)
        attendance=[str(output[0]),str(timestamp)]    


    imgBackground[130:130 + 480,65:65 + 640] = frame
    cv2.imshow("Frame", imgBackground)
    k = cv2.waitKey(1)
    if k==ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("Attendance/Attendance_" + date + ".csv","+a" ) as csvfile:
                writer=csv.writer(csvfile)
                writer.writerow(attendance)
            csvfile.close()    
        else:
            with open("Attendance/Attendance_" + date + ".csv","+a" ) as csvfile:
                writer=csv.writer(csvfile)
                writer.writerow(COL_NAME)
                writer.writerow(attendance)
            csvfile.close()    
    if k == ord('q') :
        break

video.release()
cv2.destroyAllWindows()  

