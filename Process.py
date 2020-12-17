import cv2
import numpy as np
from PIL import Image
import os
import time
import pickle

def generate_dataset(img, id, img_id):
    path = "data/"+ id

    if not os.path.exists(path):
        os.makedirs(path)

    cv2.imwrite(path + "/" + str(img_id)+".jpg", img)

def train_classifier():
    faces = []
    ids = []
    label_ids = {}
    current_id = 0

    for id in os.listdir("data"):
        path = [os.path.join("data\\" + id, f) for f in os.listdir("data\\" + id)]

        for image in path:
            img = Image.open(image).convert('L')
            image_array = np.array(img, 'uint8')
        
            faces.append(image_array)

            if not id in label_ids:
                label_ids[id] = current_id
                current_id += 1

            ids.append(label_ids[id])

    ids = np.array(ids)

    with open("labels.pickle", "wb") as f:
        pickle.dump(label_ids, f)
    
    clf = cv2.face.LBPHFaceRecognizer_create()
    clf.train(faces, ids)
    clf.save("Classifier.yml")

def draw_rect(img, classifier, scaleFactor, minNeighbors, color, label):
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    features = classifier.detectMultiScale(gray_img, scaleFactor, minNeighbors)
    coords = []

    for (x, y, w, h) in features:
        cv2.rectangle(img, (x,y), (x+w, y+h), color, 2)
        cv2.putText(img, label, (x, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 1, cv2.LINE_AA)
        coords = [x, y, w, h]

    return coords

def train_data(img, coords, id, img_id, count):
    if len(coords) == 4:
        face_img = img[coords[1]:coords[1]+coords[3], coords[0]:coords[0]+coords[2]]
        generate_dataset(face_img, id, img_id)
        count = count + 1

    return count

def train_init():
    faceCascade = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
    return faceCascade

def detect_face(img, faceCascade):
    coords = draw_rect(img, faceCascade, 1.5, 5, (255,255,255), "Yuz")
    return coords, img

def process_level(img, coords, id, count):
    count = train_data(img, coords, id, time.time(), count)
    return count

def recognize_init():
    faceCascade = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')

    clf = cv2.face.LBPHFaceRecognizer_create()
    clf.read('Classifier.yml')

    labels = {}
    with open("labels.pickle", "rb") as f:
        labels_ = pickle.load(f)
        labels = {v:k for k,v in labels_.items()}

    return faceCascade, clf, labels

def recognize_face(img, faceCascade, clf, labels):    
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    features = faceCascade.detectMultiScale(gray_img, 1.5, 5)

    label = ''

    for (x, y, w, h) in features:
        cv2.rectangle(img, (x,y), (x+w, y+h), (255, 255, 255), 2)
        id, conf = clf.predict(gray_img[y:y+h, x:x+w])
        if conf >= 45 and conf <= 85:
            label = labels[id]
        else:
            label = ''
        cv2.putText(img, label, (x, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 1, cv2.LINE_AA)

    return label, img
