# Preprocessing

# Data preprocessing
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from glob import glob
from PIL import Image
import shutil, os
import random
import cv2
from pathlib import Path
from warnings import filterwarnings
import random
import glob

# Warnings
filterwarnings("ignore",category = DeprecationWarning)
filterwarnings("ignore", category = FutureWarning) 
filterwarnings("ignore", category = UserWarning)

# Data Loading
jumfile = pd.DataFrame(columns = {'path','label','totalfile'})

def sum_data(mainpath, labels):
    for l in range(len(labels)):
        jumfile.loc[l, 'path'] = str(os.path.join(str(mainpath) + "\\" + str(labels[l])))
        jumfile.loc[l, 'label'] = str(labels[l])
        jumfile.loc[l, 'totalfile'] = len(glob.glob(os.path.join(str(mainpath) + "\\" + str(labels[l])) + "\\*"))
    return jumfile


def get_random_data(totalfile, currentpath, labels, data_type):
    data = []
    for rw in range(len(totalfile['totalfile'])):
        cur_path = totalfile['path'][rw]
        for label in labels:
            if label in cur_path:
                if totalfile['totalfile'][0] <= totalfile['totalfile'][1]:
                    totfile = int(totalfile['totalfile'][0]/2)
                else:
                    totfile = int(totalfile['totalfile'][1]/2)
                for dtaip in data_type:
                    for i in range(totfile):
                        random_index = random.randint(0,len(os.listdir(cur_path))-1)
                        shutil.copy(str(os.path.join(cur_path, str(os.listdir(cur_path)[random_index]))),
                                    str(currentpath + "\\" + dtaip + "\\" + label + "\\" + str(os.listdir(cur_path)[random_index])))


def gettransf_images(mainpath, currentpath, infected_df, normal_df, patient, run_date):
    
    print("Check if current folder exists. If not, create a new one.")
    currentpath = mainpath + "\\" + patient
    if not os.path.exists(currentpath + "\\RECLASS_RUN_" + run_date):
        os.makedirs(currentpath + "\\RECLASS_RUN_" + run_date)
        analyzed_dir = currentpath + "\\RECLASS_RUN_" + run_date
        os.makedirs(currentpath + "\\RECLASS_RUN_" + run_date + "\\Infected")
        infected_dir = currentpath + "\\RECLASS_RUN_" + run_date + "\\Infected"
        os.makedirs(currentpath + "\\RECLASS_RUN_" + run_date + "\\Normal")
        normal_dir = currentpath + "\\RECLASS_RUN_" + run_date + "\\Normal"
    
    # Prepare new datasets for re-classification
    for i in range(len(infected_df)):
        shutil.copy(currentpath + "\\" + infected_df['image'][i], str(infected_dir + "\\" +  infected_df['image'][i]))
    for i in range(len(normal_df)):
        shutil.copy(currentpath + "\\" + normal_df['image'][i], str(normal_dir + "\\" +  normal_df['image'][i]))

    return infected_dir, normal_dir, analyzed_dir