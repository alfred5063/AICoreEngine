# Main Script

import numpy as np
import pandas as pd
import tensorflow as tf
from tensorflow.keras.preprocessing.image import ImageDataGenerator
from tensorflow.keras.metrics import categorical_accuracy
import matplotlib.pyplot as plt
from tensorflow.keras.preprocessing import image
import os, time
from tensorflow.keras.models import load_model
from tensorflow.keras.optimizers import Adam
from backend.loading.xray_dataload import *
from datetime import datetime
import os.path


mainpath = r"D:\ALFRED - Workspace\Xray Images"
data_type = ['train_dataset', 'test_dataset']
batch_size = 32

def getdate():
    try:
        print("Date and Timestamp")
        current_year = datetime.now().year
        current_timestamp = time.strftime("%Y%m%d_%H%M%S")
        run_date = time.strftime("%Y%m%d")
        return current_year, current_timestamp, run_date
    except Exception as error:
        print("Error in Date and Timestamp", error)
def loadata(mainpath, run_date, data_type):
    try:
        list = []
        print("Check if current folder exists. If not, create a new one.")
        currentpath = mainpath + "\\RUN_" + run_date
        if not os.path.exists(currentpath):
            os.makedirs(currentpath)
            os.makedirs(currentpath + "\\train_dataset")
            os.makedirs(currentpath + "\\train_dataset\\patients_covid")
            os.makedirs(currentpath + "\\train_dataset\\patients_normal")

            train_patients_covid = currentpath + "\\train_dataset\\patients_covid"
            list.append(train_patients_covid)
            train_patients_normal = currentpath + "\\train_dataset\\patients_normal"
            list.append(train_patients_normal)

            os.makedirs(currentpath + "\\test_dataset")
            os.makedirs(currentpath + "\\test_dataset\\patients_covid")
            os.makedirs(currentpath + "\\test_dataset\\patients_normal")

            test_patients_covid = currentpath + "\\test_dataset\\patients_covid"
            list.append(test_patients_covid)
            test_patients_normal = currentpath + "\\test_dataset\\patients_normal"
            list.append(test_patients_normal)

        time.sleep(5)
    except Exception as error:
        print("Error in creating new run folder", error)

    try:
        print("Data loading")
        labels = ['patients_covid', 'patients_normal']
        totalfile = sum_data(mainpath, labels)
        get_random_data(totalfile, currentpath, labels, data_type)
    except Exception as error:
        print("Error in data loading", error)

    return currentpath, list, labels, totalfile
def prep(currentpath, data_type, batch_size):
    train_datagen = ImageDataGenerator(rescale = 1./255, shear_range = 0.2, zoom_range = 0.2, horizontal_flip = True)
    training_set = train_datagen.flow_from_directory(str(currentpath + "\\" + data_type[0]),
                                                     target_size = (128, 128),
                                                     batch_size = batch_size,
                                                     class_mode = 'binary')

    test_datagen = ImageDataGenerator(rescale = 1./255, shear_range = 0.2, zoom_range = 0.2, horizontal_flip = True)
    test_set = test_datagen.flow_from_directory(str(currentpath + "\\" + data_type[1]),
                                                target_size = (128, 128),
                                                batch_size = batch_size,
                                                class_mode = 'binary')

    steps_per_epoch = training_set.samples // batch_size
    val_steps = test_set.samples // batch_size
    return training_set, test_set, steps_per_epoch, val_steps
def modelit(currentpath, mainpath, training_set, test_set, steps_per_epoch, val_steps, epochs_range):

    # Referenced from: https://www.learndatasci.com/tutorials/convolutional-neural-networks-image-classification/

    cnn = tf.keras.models.Sequential()

    #### Input Layer ####
    cnn.add(tf.keras.layers.Conv2D(filters=32, kernel_size=(3,3), padding='same', activation='relu', input_shape=(128, 128, 3)))

    #### Convolutional Layers ####
    cnn.add(tf.keras.layers.Conv2D(32, (3,3), activation='relu'))
    cnn.add(tf.keras.layers.MaxPooling2D((2,2)))  # Pooling
    cnn.add(tf.keras.layers.Dropout(0.2)) # Dropout

    cnn.add(tf.keras.layers.Conv2D(64, (3,3), padding='same', activation='relu'))
    cnn.add(tf.keras.layers.Conv2D(64, (3,3), activation='relu'))
    cnn.add(tf.keras.layers.MaxPooling2D((2,2)))
    cnn.add(tf.keras.layers.Dropout(0.2))

    cnn.add(tf.keras.layers.Conv2D(128, (3,3), padding='same', activation='relu'))
    cnn.add(tf.keras.layers.Conv2D(128, (3,3), activation='relu'))
    cnn.add(tf.keras.layers.Activation('relu'))
    cnn.add(tf.keras.layers.MaxPooling2D((2,2)))
    cnn.add(tf.keras.layers.Dropout(0.2))

    cnn.add(tf.keras.layers.Conv2D(512, (5,5), padding='same', activation='relu'))
    cnn.add(tf.keras.layers.Conv2D(512, (5,5), activation='relu'))
    cnn.add(tf.keras.layers.MaxPooling2D((4,4)))
    cnn.add(tf.keras.layers.Dropout(0.2))

    #### Fully-Connected Layer ####
    cnn.add(tf.keras.layers.Flatten())
    cnn.add(tf.keras.layers.Dense(1024, activation='relu'))
    cnn.add(tf.keras.layers.Dropout(0.2))
    #cnn.add(tf.keras.layers.Dense(128, activation='sigmoid')) # Or 'softmax'
    cnn.add(tf.keras.layers.Dense(units = 1 , activation = 'sigmoid')) #Sigmoid is used because we want to predict probability of Covid-19 infected category

    optimizer = Adam(lr = 0.00001)
    cnn.compile(optimizer = optimizer, loss = 'binary_crossentropy', metrics  =['accuracy'])
    #cnn.compile(optimizer = optimizer, loss = 'categorical_crossentropy', metrics = ['accuracy'])
    cnn.summary()

    # Modelling
    history = cnn.fit(training_set,
                      epochs = epochs_range,
                      steps_per_epoch = steps_per_epoch,
                      validation_data = test_set,
                      validation_steps = val_steps,
                      verbose = True)
    cnn.save("D:\\decodevid\\model\\model.h5")

    # Evaluating the result
    acc = history.history['accuracy']
    val_acc = history.history['val_accuracy']
    loss = history.history['loss']
    val_loss = history.history['val_loss']

    # Keras Accuracy
    keras_score = cnn.evaluate(training_set, verbose = 0)

    #plt.figure(figsize = (15, 15))
    #plt.subplot(2, 2, 1)
    #plt.plot(epochs_range, acc, label = 'Training Accuracy')
    #plt.plot(epochs_range, val_acc, label = 'Validation Accuracy')
    #plt.legend(loc = 'lower right')
    #plt.title('Training and Validation Accuracy')
    #plt.savefig(currentpath + "\\Epoch_" + epochs_range + "_training_validationaccuracy.png")

    #plt.subplot(2, 2, 2)
    #plt.plot(epochs_range, loss, label = 'Training Loss')
    #plt.plot(epochs_range, val_loss, label = 'Validation Loss')
    #plt.legend(loc = 'upper right')
    #plt.title('Training and Validation Loss')
    #plt.savefig(currentpath + "\\Epoch_" + epochs_range + "_training_validationloss.png")

    epochdf = pd.DataFrame()
    epochdf = epochdf.rename(columns = {'0': 'epoch',
                                        '1': 'acc',
                                        '2': 'val_acc',
                                        '3': 'loss', 
                                        '4': 'val_loss',
                                        '5': 'keras_score'}, inplace = False)

    epochdf.loc[1, 'epoch'] = str(epochs_range)
    epochdf.loc[1, 'acc'] = str(sum(acc)/len(acc))
    epochdf.loc[1, 'val_acc'] = str(sum(val_acc)/len(val_acc))
    epochdf.loc[1, 'loss'] = str(sum(loss)/len(loss))
    epochdf.loc[1, 'val_loss'] = str(sum(val_loss)/len(val_loss))
    epochdf.loc[1, 'keras_score'] = str(keras_score)

    return acc, val_acc, loss, val_loss, keras_score, epochdf
def predicting(mainpath, label):
    final_result = []
    if label != 'None':
        img_path = str(mainpath + "\\" + label)
        os.chdir(str(mainpath + "\\" + label))
    else:
        img_path = str(mainpath)
        os.chdir(str(mainpath))
    if os.path.isfile("D:\\ALFRED - Workspace\\Analytics\\model.h5") == True:
        try:
            cnn = load_model("D:\\ALFRED - Workspace\\Analytics\\model.h5")
        except Exception as error:
            print("Model cannot be loaded with error message as: %s", error)
    else:
        print("Path to model does not exist.")
    
    for file in os.listdir():
        if label != 'None':
            test_image = image.load_img(str(mainpath + "\\" + label + "\\" + file), target_size = (128, 128))
        else:
            test_image = image.load_img(str(mainpath + "\\" + file), target_size = (128, 128))
        test_image = image.img_to_array(test_image)
        test_image = np.expand_dims(test_image, axis = 0)

        # CNN Model
        try:
            result = cnn.predict(test_image)
            #resultline = img_path + "\\" + str(file), str(file), 'Predictions: %', (np.any(result)*100).astype(float), 'Normal' if np
            resultline = img_path + "\\" + str(file), str(file), 'Predictions: %', float(result*100), 'Normal' if float(result) < 0.5 else 'Infected'
            final_result.append(resultline)
            print(str(file), 'Predictions: %', float(result*100), 'Normal' if float(result) < float(0.5) else 'Infected')
        except:
            pass
        continue
    
    final_resultdf = pd.DataFrame(final_result)
    #if label != 'None':
    #    final_resultdf.to_csv(str(mainpath + "\\" + label + "\\final_resultdf.csv"))
    #else:
    #    final_resultdf.to_csv(str(mainpath + "\\final_resultdf.csv"))
    return final_resultdf
def assessment(mainpath, patients):

    final_resultdf = []
    infected = []
    infected_df = []
    normal = []
    normal_df = []

    final_resultdf = pd.read_csv(mainpath + "\\" + patients + "\\final_resultdf.csv")
    final_resultdf = final_resultdf.rename(columns = {'0': 'image_path',
                                                      '1': 'image',
                                                      '2': 'prediction',
                                                      '3': 'percent', 
                                                      '4': 'status'}, inplace = False).drop(columns = {'Unnamed: 0'})
    infected = sum(final_resultdf['status']=='Infected')
    infected_df = final_resultdf[final_resultdf['status']=='Infected']
    infected_df = infected_df.reset_index()
    infected_df = infected_df.drop(columns = {'index'})
    
    normal = sum(final_resultdf['status']=='Normal')
    normal_df = final_resultdf[final_resultdf['status']=='Normal']
    normal_df = normal_df.reset_index()
    normal_df = normal_df.drop(columns = {'index'})

    return infected, infected_df, normal, normal_df, final_resultdf
def rgb_analysis(dir, coviddata):

    mean_val = []
    std_dev_val = []
    max_val = []
    min_val = []

    coviddata['image_bin'] = coviddata['image_path'].map(lambda x: np.asarray(Image.open(x).resize((75,75))))

    for i in range(len(coviddata['image'])):
        mean_val.append(coviddata['image_bin'][i].mean())
        std_dev_val.append(np.std(coviddata['image_bin'][i]))
        max_val.append(coviddata['image_bin'][i].max())
        min_val.append(coviddata['image_bin'][i].min())

    coviddata['mean'] = mean_val
    coviddata['stedev'] = std_dev_val
    coviddata['max'] = max_val
    coviddata['min'] = min_val
    subt_mean_samples = coviddata['mean'].mean() - coviddata['mean']
    coviddata['subt_mean'] = subt_mean_samples
    ax = sns.displot(data = coviddata, x = 'mean', kind="kde")
    plt.title('Images Colour Mean Value Distribution', fontsize = 16, weight = 'bold')
    plt.savefig(dir + "\\Images_Colour_Mean_Value_Distribution.png")

    ax = sns.displot(data = coviddata, x = 'mean', kind="kde", hue = 'status')
    plt.title('Images Colour Mean Value Distribution by Class', fontsize = 16, weight = 'bold')
    plt.savefig(dir + "\\Images_Colour_Mean_Value_Distribution_by_Class.png")

    ax = sns.displot(data = coviddata, x = 'max', kind="kde", hue = 'status');
    plt.title('Images Colour Max Value Distribution by Class', fontsize = 16, weight = 'bold')
    plt.savefig(dir + "\\Images_Colour_Max_Value_Distribution_by_Class.png")

    ax = sns.displot(data = coviddata, x = 'min', kind="kde", hue = 'status');
    plt.title('Images Colour Min Value Distribution by Class', fontsize = 16, weight = 'bold')
    plt.savefig(dir + "\\Images_Colour_Min_Value_Distribution_by_Class.png")

    return coviddata

# Re-model?
current_year, current_timestamp, run_date = getdate()
currentpath, list, labels, totalfile = loadata(mainpath, run_date, data_type)
training_set, test_set, steps_per_epoch, val_steps = prep(currentpath, data_type, batch_size)
e = 5
epochdf_final = pd.DataFrame()
while e <= 5:
    acc, val_acc, loss, val_loss, keras_score, epochdf = modelit(currentpath, mainpath, training_set, test_set, steps_per_epoch, val_steps, e)
    epochdf_final = epochdf_final.append(epochdf, ignore_index = True)
    e += 5
epochdf_final = epochdf_final.reset_index()
epochdf_final = epochdf_final.drop(columns = {'index'}).drop_duplicates(subset = None, keep = 'first', inplace = False)
epochdf_final.to_csv(currentpath + "\\Epoch_" + str(e) + "_final_resultdf.csv")

# Run IT
final_result = []
types = ['Re-modelling'
         #'patients_bacterialpneumonia',
         #'patients_lungopacity',
         #'patients_normal',
         #'patients_pneumonia',
         #'patients_viralpneumonia'
         ]
for t in range(len(types)):
    try:
        final_result = predicting(mainpath, str(types[t]))
        if final_result.empty != True:
            final_resultdf = pd.DataFrame(final_result)
            final_resultdf.to_csv(mainpath + "\\" + str(types[t]) + "_prediction_result.csv")
        else:
            print(str(types[t]) + " failed 1")
            pass
    except:
        print(str(types[t]) + " failed 2")
        pass
    continue

# Re-assessments
current_year, current_timestamp, run_date = getdate()
infected, infected_df, normal, normal_df, final_resultdf = assessment(mainpath, 'patients_covid')
infected_dir, normal_dir, analyzed_dir = gettransf_images(mainpath, currentpath, infected_df, normal_df, 'patients_covid', run_date)
coviddata = rgb_analysis(analyzed_dir, final_resultdf)
final_result = predicting(infected_dir, 'None')# -*- coding: utf-8 -*-

