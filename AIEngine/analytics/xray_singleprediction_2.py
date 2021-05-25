# From https://www.kaggle.com/hosseinmoradian/covid19-classification-x-ray-images

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

# Filter Warnings
warnings.filterwarnings("ignore")

# File Dir
input_path = "D:\ALFRED - Workspace\Xray Images"

# File Contents
for _set in ['train_dataset', 'test_dataset']:
    Normal = len(os.listdir(input_path + "\\" + _set + '\\patients_normal'))
    Infected = len(os.listdir(input_path + "\\" +  _set + '\\patients_covid'))
    print('The {} folder contains {} Normal and {} Infected images.'.format(_set, Normal, Infected))

# Preprocesing Data Function
img_dims = 180
epochs = 25
batch_size = 2
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
        for img in (os.listdir(input_path + '\\train_dataset' + file_name)):
            img = cv2.imread(input_path + '\\train_dataset' + file_name + img, cv2.IMREAD_GRAYSCALE)
            img = cv2.resize(img, (img_dims, img_dims))
            img = np.dstack([img, img, img])
            img = img.astype('float32') / 255
            if file_name == '/patients_normal/':
                label = 0
            elif file_name == '/patients_covid/':
                label = 1
            train_labels.append(label)            
            
            
    for file_name in ['\\patients_normal\\', '\\patients_covid\\']:
        for img in (os.listdir(input_path + '\\test_dataset' + file_name)):
            img = cv2.imread(input_path + '\\test_dataset' + file_name + img, cv2.IMREAD_GRAYSCALE)
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

# Set Images&Labels for Train,Test
train_images, train_labels, test_images, test_labels = preprocess_data(input_path, img_dims, batch_size)
#print(test_labels)

# Create Model with KERAS library
model = Sequential()
model.add(Conv2D(32, (3,3), activation="relu", input_shape=(img_dims,img_dims,3)))
model.add(MaxPooling2D(2,2))
model.add(Conv2D(64, (3,3), activation="relu"))
model.add(MaxPooling2D(2,2))
model.add(Conv2D(128, (3,3), activation="relu"))
model.add(MaxPooling2D(2,2))
model.add(Conv2D(128, (3,3), activation="relu"))
model.add(MaxPooling2D(2,2))
model.add(Conv2D(256, (3,3), activation="relu"))
model.add(MaxPooling2D(2,2))
model.add(Flatten())
model.add(Dense(512, activation="relu"))
model.add(Dropout(0.5))
model.add(Dense(1,activation="sigmoid"))
model.summary()

# Set Optimizer
optimizer = Adam(lr = 0.0001)

# Compile Model
model.compile( optimizer= optimizer,loss='binary_crossentropy', metrics=['acc'])

# Save Model
model.save("D:\ALFRED - Workspace\Analytics\Covid19-CNN-Model.h5")
model.save_weights("D:\ALFRED - Workspace\Analytics\Covid19-CNN-weights.h5")

# Fit the Model
history = model.fit_generator(
            train_images,
            steps_per_epoch = train_images.samples // batch_size, 
            epochs = epochs, 
            validation_data = test_images,
            validation_steps = test_images.samples // batch_size)

# Visualize Loss and Accuracy Rates
fig, ax = plt.subplots(2, 1, figsize=(10, 7))
ax = ax.ravel()
plt.style.use("ggplot")

for i, met in enumerate(['acc', 'loss']):
    ax[i].plot(history.history[met])
    ax[i].plot(history.history['val_' + met])
    ax[i].set_title('Model {}'.format(met))
    ax[i].set_xlabel('Epochs')
    ax[i].set_ylabel(met)
    ax[i].legend(['Train', 'Val'])

# Adding Predictions, Confusion Matrix & Performance Metrics

# Prediction on Model
Y_pred = model.predict(test_images)
Y_pred = [ 1 if y >= 0.5 else 0 for y in Y_pred]

print(len(Y_pred),len(test_labels))
# Confusion Matrix
#from sklearn.metrics import confusion_matrix
#cm = confusion_matrix(test_labels, Y_pred)

#from mlxtend.plotting import plot_confusion_matrix
#fig, ax = plot_confusion_matrix(conf_mat=cm)
#plt.show()

# Performance Metrics
#from sklearn.metrics import accuracy_score
#from sklearn.metrics import classification_report 
#print('Confusion Matrix :')
#print(cm) 
#print('Accuracy Score :',accuracy_score(test_labels, Y_pred))
#print('Report : ')
#print(classification_report(test_labels, Y_pred))

# Image Classifer Script
cnn = load_model("D:\\ALFRED - Workspace\\Analytics\\Covid19-CNN-Model.h5")
#cnn = load_weights("D:\\ALFRED - Workspace\\Analytics\\Covid19-CNN-weights.h5")

def predict_image(model, img_path, file, img_dims = 180):
    img = image.load_img(img_path, target_size = (img_dims, img_dims))
    #plt.imshow(img)
    #plt.show()
    img = image.img_to_array(img)
    x = np.expand_dims(img, axis=0) * 1./255
    score = model.predict(x)
    print(str(file), 'Predictions: %', (float)(score*100), 'NORMAL' if score < 0.5 else 'Infected')

# Test on Validation Images
os.chdir(str(mainpath + "\\Xray Images\\patients_covid"))
for file in os.listdir():
    predict_image(model,(str(mainpath + "\\Xray Images\\patients_covid\\" + file)), str(file))