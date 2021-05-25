# Deteksi Covid-19 dengan gambar X-Ray - CNN
# https://www.kaggle.com/kingdilan001/deteksi-covid-19-dengan-gambar-x-ray-cnn

import numpy as np # linear algebra
import pandas as pd # data processing, CSV file I/O (e.g. pd.read_csv)
import tensorflow as tf
import cv2
from keras.preprocessing.image import ImageDataGenerator
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from keras.preprocessing import image
import os
from keras.models import load_model

mainpath = r"D:\ALFRED - Workspace"

def prep(mainpath):
    train_datagen = ImageDataGenerator(rescale = 1./255, shear_range = 0.2, zoom_range = 0.2, horizontal_flip = True)
    training_set = train_datagen.flow_from_directory(str(mainpath + "\\Xray Images\\train_dataset"),
                                                     target_size = (64, 64),
                                                     batch_size = 32,
                                                     class_mode = 'binary')

    test_datagen = ImageDataGenerator(rescale = 1./255)
    test_set = test_datagen.flow_from_directory(
        str(mainpath + "\\Xray Images\\train_dataset"), # same directory as training data
        target_size=(64, 64),
        batch_size=32,
        class_mode='binary')

    return training_set, test_set

def modelit(mainpath, training_set, test_set):
    cnn = tf.keras.models.Sequential()
    cnn.add(tf.keras.layers.Conv2D(filters=32, kernel_size=3, activation='relu', input_shape=[64, 64, 3]))
    cnn.add(tf.keras.layers.MaxPool2D(pool_size=2, strides=2))
    cnn.add(tf.keras.layers.Conv2D(filters=32, kernel_size=3, activation='relu'))
    cnn.add(tf.keras.layers.MaxPool2D(pool_size=2, strides=2))
    cnn.add(tf.keras.layers.Flatten())
    cnn.add(tf.keras.layers.Dense(units=128, activation='relu'))
    cnn.add(tf.keras.layers.Dense(units=1, activation='sigmoid'))
    cnn.compile(optimizer = 'adam', loss = 'binary_crossentropy', metrics = ['accuracy'])

    # Modelling
    cnn.fit(x = training_set, validation_data = test_set, epochs = 100)
    cnn.save(str(mainpath + "\\Analytics\\model.h5"))

# Single prediction
def predicting(mainpath):
    os.chdir(str(mainpath + "\\Xray Images\\patients_covid"))
    cnn = load_model(str(mainpath + "\\Analytics\\model.h5"))
    for file in os.listdir():
        test_image = image.load_img(str(mainpath + "\\Xray Images\\patients_covid\\" + file), target_size = (64, 64))
        test_image = image.img_to_array(test_image)
        test_image = np.expand_dims(test_image, axis = 0)

        # CNN Model
        result = cnn.predict(test_image)
        print(str(file), 'Predictions: %', (float)(result*100), 'NORMAL' if result < 0.5 else 'Infected')

# Run IT
predicting(mainpath)

# Re-model?
training_set, test_set = prep(mainpath)
modelit(mainpath, training_set, test_set)