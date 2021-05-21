import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from warnings import filterwarnings
from sklearn.metrics import confusion_matrix, accuracy_score, classification_report, roc_auc_score, roc_curve
from tensorflow.keras.models import Sequential
from keras.layers import Dense, Dropout, Flatten, Conv2D, MaxPool2D, BatchNormalization,MaxPooling2D
from keras import models
from keras import layers
import tensorflow as tf
import os
import os.path
from pathlib import Path
import cv2 #pip install opencv-python
from tensorflow.keras.preprocessing.image import ImageDataGenerator
from keras.utils.np_utils import to_categorical
from sklearn.model_selection import train_test_split
from keras import regularizers
from keras.optimizers import RMSprop,Adam
import glob
from PIL import Image
from sklearn.preprocessing import StandardScaler
from keras.preprocessing import image

filterwarnings("ignore",category=DeprecationWarning)
filterwarnings("ignore", category=FutureWarning) 
filterwarnings("ignore", category=UserWarning)

Bact_Pneu_Data = Path("../input/covid19-detection-xray-dataset/TrainData/BacterialPneumonia")
Covid_Data = Path("../input/covid19-detection-xray-dataset/TrainData/COVID-19")
Normal_Data = Path("../input/covid19-detection-xray-dataset/TrainData/Normal")
Over_Samp_Aug_Data = Path("../input/covid19-detection-xray-dataset/TrainData/OversampledAugmentedCOVID-19")
Viral_Pneu = Path("../input/covid19-detection-xray-dataset/TrainData/ViralPneumonia")