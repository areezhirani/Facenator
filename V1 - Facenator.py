# -*- coding: utf-8 -*-
"""
Created on Thu Jan  6 20:06:10 2022

@author: areez
"""

import face_recognition
import cv2
import numpy as np
import pandas as pd
import os



# Get a reference to webcam #0 (the default one)
video_capture = cv2.VideoCapture(0)

# Load a sample picture and learn how to recognize it.
person1_image = face_recognition.load_image_file("C:\\Users\\areez\\Downloads\\faceid\\87326.jpg")
person1_face_encoding = face_recognition.face_encodings(person1_image)[0]

# Load a second sample picture and learn how to recognize it.
person2_image = face_recognition.load_image_file("C:\\Users\\areez\\Downloads\\faceid\\sheetalimg.jfif")
person2_face_encoding = face_recognition.face_encodings(person2_image)[0]

# Load a third sample picture and learn how to recognize it.
person3_image = face_recognition.load_image_file("C:\\Users\\areez\\Downloads\\faceid\\sabrina.jfif")
person3_face_encoding = face_recognition.face_encodings(person3_image)[0]

# Create arrays of known face encodings and their names
known_face_encodings = [
    person1_face_encoding,
    person2_face_encoding,
    person3_face_encoding
]
known_face_names = [
    "Areez Hirani",
    "Sheetal Hirani",
    "Sabrina Hirani"
]

# Initialize some variables
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True

namelimit = ("")
#SPREADSHEET
os.chdir("C:\\Science Fair Temp")
PathAndFileString = "C:\\Users\\areez\\OneDrive\\Documents\\Python\\Science Fair\\spreadsheet stuff\\attendance.xlsx"
df = pd.read_excel(PathAndFileString, 'Sheet1')

areezcounter = 0
sheetalcounter = 0
sabrinacounter = 0


while True:
    # Grab a single frame of video
    ret, frame = video_capture.read()

    # Resize frame of video to 1/4 size for faster face recognition processing
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

    # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
    rgb_small_frame = small_frame[:, :, ::-1]

    # Only process every other frame of video to save time
    if process_this_frame:
        # Find all the faces and face encodings in the current frame of video
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

        face_names = []
        for face_encoding in face_encodings:
            # See if the face is a match for the known face(s)
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "Unknown"


            # Or instead, use the known face with the smallest distance to the new face
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)
            if matches[best_match_index]:
                name = known_face_names[best_match_index]
                

                if(name=="Areez Hirani"):
                    areezcounter += 1
                    if areezcounter > 2:
                        FullName = df['Name'][0]
                        print("Full Name: ",FullName)
                        df['Present'][0] = "Y"
                        print (areezcounter)
                        
                if(name=="Sheetal Hirani"):
                    sheetalcounter += 1
                    if sheetalcounter > 2:
                        FullName = df['Name'][1]
                        print("Full Name: ",FullName)
                        df['Present'][1] = "Y"
                        print(sheetalcounter)
                        
                if(name=="Sabrina Hirani"):
                    sabrinacounter += 1
                    if sabrinacounter > 2:
                        FullName = df['Name'][2]
                        print("Full Name: ",FullName)
                        df['Present'][2] = "Y"


                print(df)               
                NewFileName = 'Attendance.xlsx'
                df.to_excel(NewFileName)

            face_names.append(name)
            name = name

    process_this_frame = not process_this_frame


    # Display the results
    for (top, right, bottom, left), name in zip(face_locations, face_names):
        # Scale back up face locations since the frame we detected in was scaled to 1/4 size
        top *= 4
        right *= 4
        bottom *= 4
        left *= 4

        # Draw a box around the face
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

        # Draw a label with a name below the face
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

    # Display the resulting image
    cv2.imshow('Video', frame)
    
    #adding value to spreadsheet
    #namelimit = ("")
    #sheetname = str(face_names)
    #if namelimit != sheetname:
        #df = pd.read_excel("C:\\Users\\areez\\OneDrive\\Documents\\Python\\Science Fair\\spreadsheet stuff\\attendance.xlsx");
        #YourDataInAList = name
        #df = df.append({"People Present": sheetname}, ignore_index=True)
        #df.to_excel("C:\\Users\\areez\\OneDrive\\Documents\\Python\\Science Fair\\spreadsheet stuff\\attendance.xlsx",index=False);
        #print(sheetname)
        #namelimit = sheetname
        #print(namelimit + "f") # smthg wrong with namelimit 
        #ask pops how to read spreadsheet to see if theres already the same name

    # Hit 'q' on the keyboard to quit!
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Release handle to the webcam
video_capture.release()
cv2.destroyAllWindows()

