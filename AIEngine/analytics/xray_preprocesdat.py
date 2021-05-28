# Importing Lib & tools
import os
import cv2
import scipy
import warnings
import numpy as np
from tqdm import tqdm
from random import shuffle
import matplotlib.pyplot as plt
from keras.optimizers import Adam, RMSprop
from keras.models import Sequential, Model
from keras.layers import  Conv2D, MaxPooling2D, Activation, Flatten,Dense, Dropout, GlobalAveragePooling2D, BatchNormalization
from keras.preprocessing.image import ImageDataGenerator, img_to_array, load_img
from PIL import Image
from keras.preprocessing import image
from keras.models import load_model
from warnings import filterwarnings
from analytics.xray_dataload import *

# Filter Warnings
filterwarnings("ignore",category = DeprecationWarning)
filterwarnings("ignore", category = FutureWarning) 
filterwarnings("ignore", category = UserWarning)

def preprocess_data(input_path, img_dims, batch_size):
    
    # Data Augmentation for Infected & Normal Images
    train_datagen = ImageDataGenerator(
        rescale = 1./255,
        zoom_range = 0.5,
        shear_range = 0.6,      
        rotation_range = 30,
        width_shift_range = 0.3,
        height_shift_range = 0.5,
        horizontal_flip = True,
        fill_mode='nearest')

    test_datagen = ImageDataGenerator(
        rescale = 1./255)
    
    train_images = train_datagen.flow_from_directory(
        directory = input_path + '\\train_dataset', 
        target_size = (img_dims, img_dims), 
        batch_size = batch_size, 
        class_mode = 'binary')

    test_images = test_datagen.flow_from_directory(
        directory = input_path + '\\test_dataset', 
        target_size = (img_dims, img_dims), 
        batch_size = batch_size, 
        class_mode = 'binary')

    # Creating these lists for make prediction on test image and showing confusion matrix.
    train_labels = []
    test_labels = []
    label = ''

    for file_name in ['\\patients_normal\\', '\\patients_covid\\']:
        for img in (os.listdir(input_path + '\\train_dataset' + file_name + 'TEMP')):
            img = cv2.imread(input_path + '\\train_dataset' + file_name + 'TEMP\\' + img, cv2.IMREAD_GRAYSCALE)
            img = cv2.resize(img, (img_dims, img_dims))
            img = np.dstack([img, img, img])
            img = img.astype('float32') / 255
            if file_name == '/patients_normal/':
                label = 0
            elif file_name == '/patients_covid/':
                label = 1
            train_labels.append(label)            
            
            
    for file_name in ['\\patients_normal\\', '\\patients_covid\\']:
        for img in (os.listdir(input_path + '\\test_dataset' + file_name  + 'TEMP')):
            img = cv2.imread(input_path + '\\test_dataset' + file_name + 'TEMP\\' + img, cv2.IMREAD_GRAYSCALE)
            img = cv2.resize(img, (img_dims, img_dims))
            img = np.dstack([img, img, img])
            img = img.astype('float32') / 255
            if file_name == '/patients_normal/':
                label = 0
            elif file_name == '/patients_covid/':
                label = 1
            test_labels.append(label)
        
    train_labels = np.array(train_labels)
    test_labels = np.array(test_labels)
    
    return train_images, train_labels, test_images, test_labels
