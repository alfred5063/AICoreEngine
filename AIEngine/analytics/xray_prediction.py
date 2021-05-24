# Covid-19 Detection from Chest Xray using Tensorflow
'''
Using pytorch detecting Covid-19 Infected Lungs from Normal Lungs with Chest X-Ray
From https://www.kaggle.com/ankan1998/chest-x-ray-to-detect-covid-19-with-pytorch
'''

import torch
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings('ignore')
from torchvision.datasets import ImageFolder
from torchvision import transforms
import torch.nn as nn
import torch.nn.functional as F
import torch.optim as optim
from pathlib import Path
from torch.utils.data.sampler import SubsetRandomSampler #samples randomly from given indices
from torch.utils.data.dataloader import DataLoader # loads the data from sampler

path_covid = Path("D:\ALFRED - Workspace\Xray Images\patients_covid")
path_normal = Path("D:\ALFRED - Workspace\Xray Images\patients_normal")

# Works similiar as ImageData Generator of Keras
# Defining transform to resize 1024x1024 to 128x128
# To change to Tensor
transform=transforms.Compose([
                              transforms.Resize([64,64]),
                              transforms.ToTensor()
])

train_val_path = r"D:\ALFRED - Workspace\Xray Images\train_dataset"
dataset=ImageFolder(train_val_path,transform=transform)
# Checking For Samples
img0,label0=dataset[10]
print(img0.shape,label0)
img1,label1=dataset[300]
print(img1.shape,label1)
print("*"*60)
print(dataset.classes)#list out all the classes

# Splitting the data into train and validation set
def split_train_val(tot_img,val_percentage=0.2,rnd=23):
  # Here indices are randomly permuted 
  number_of_val=int(tot_img*val_percentage)
  np.random.seed(rnd)
  indexs=np.random.permutation(tot_img)
  return indexs[number_of_val:],indexs[:number_of_val]

randomness=12
val_per=0.5
train_indices,validation_indices=split_train_val(len(dataset),val_per,randomness)
print(validation_indices[:5])

# Subset random sampler takes the indices to pick the data
# dataloader loads with the main dataset, with batch size and the sampler object
batch_size=16
# Training Part
train_sampler=SubsetRandomSampler(train_indices)
train_ds=DataLoader(dataset,batch_size,sampler=train_sampler)

# Validation Part
val_sampler=SubsetRandomSampler(validation_indices)
val_ds=DataLoader(dataset,batch_size,sampler=val_sampler)

# Applying CNN
# 1st layer of Conv2d
# 1st Argument is number of color channel for RGB=3, for BW=1
# 2nd Argument if number of filters, 3rd is filter size
# how to calculate its output directly to Linear Unit
# image=3x64x64-->64-3(filter_size)+1=62. So, output is 62x62xnumber of filter
# then 62x62xnum_of_filter-->maxpool(2,2)-->62/2=31-->8*31*31
# for same layer number of filter of output is input of new channel
# for conv to linear layer the above calculation is reqd.

"""
# Use nn.Sequential for implementation
# Though I myself doesnt like it.
model=nn.Sequential(
        nn.Conv2d(3,8,3), 
        nn.ReLU(),
        nn.MaxPool2d(2,2),
        nn.Flatten(start_dim=1), #.view(-1,)
        nn.Linear(8*31*31,2)

)
"""
# Recommended to use Object Oriented Neural Network
# pytorch nn library provide with 2 component on abstarct level.
# (i) is transformation i.e code (ii) Collection of weight - data
# class Module base class for all nn module
# Every neural network inherits from nn.Module class
class ConvNet(nn.Module):
  def __init__(self):
    # super here used access method of parent class
    # dont worry much just boiler plate
    super(ConvNet,self).__init__()
    # In conv layer in_channels== input; out_channels=output; kernel_size=filter size
    self.conv1=nn.Conv2d(in_channels=3,out_channels=8,kernel_size=3)
    # Linear layer in_features is input, how 8*31*31 came is explained in above comment
    # out_features= output
    self.fc1=nn.Linear(in_features=8*31*31,out_features=32)
    self.out=nn.Linear(in_features=32,out_features=2)

  def forward(self,l):
    # this method implements forward propagation
    # So, layers are structured as such

    # 1 Conv layer
    # may be thinking self.conv1 is an layer object instance how can we call as if it a function
    # Checkout python documents __call__ this special method is used, so that instances behaves like function
    # __call__ this special method invokes anytime the object instance is called. This interacts with forward method.
    l=self.conv1(l)
    l=F.relu(l)
    l=F.max_pool2d(l,kernel_size=2)

    # linear and final layer
    # -1 indicates, give any number of batch size
    l=l.reshape(-1,8*31*31)
    l=self.fc1(l)
    l=self.out(l)

    return l

model=ConvNet()
# If gpu present then use it or else use cpu
# if gpu not present dont run this cell
def default_device():
    
    if torch.cuda.is_available():
        return torch.device("cuda:0")
    else:
        return torch.device("cpu")

device=default_device()

# Loading model on GPU
model.to(device)

# Define loss and optimizer
loss_type = nn.CrossEntropyLoss()
# Adam optimizer is the combination of momentum with RMSprop and is more powerful
optimizer = optim.Adam(model.parameters(), lr=0.0005)

loss_val=[]
for epoch in range(12):  
# loop over the dataset multiple times
    print("Epoch count-->",epoch)
    running_loss=0.0
    for i, data in enumerate(train_ds):
        
        inputs, labels = data
        # Loading inputs,labels on GPU
        inputs,labels=inputs.to(device),labels.to(device)
        #inputs,labels=inputs,labels
        # zero the parameter gradients
        optimizer.zero_grad()

        # Passing input into the model
        outputs = model(inputs)
        
        # Caculating loss with crossentropy
        loss = loss_type(outputs, labels)
        
        # calculates the gradient 
        loss.backward()
        
        # update the weights
        optimizer.step()
        
        running_loss=running_loss+loss.item()* inputs.size(0)
        
    loss_val.append(running_loss / len(train_ds))
        
    print(running_loss)
        
plt.plot(loss_val,label="loss")
plt.legend()
print('Finished Training')

right = 0
total = 0

with torch.no_grad():
# Switching off the gradient part, so that backpropagation doesnt take place
    for data in val_ds:
        images, labels = data
        #images,labels=images,labels
        inputs,labels=inputs.to(device),labels.to(device)
        outputs = model(images)
        
        _, predicted = torch.max(outputs,dim=1)
        total += labels.size(0)
        # Caculating number of right prediction
        right += (predicted == labels).sum()

print('Accuracy on the validation images: %d %%' % (
    100 * right / total))


# This is completely from different source from train,val data
test_path = "/tmp/Xray_test_data"
transform=transforms.Compose([
                              transforms.Resize([64,64]),
                              transforms.ToTensor()])
test_dataset=ImageFolder(test_path,transform=transform)

# Checking For Samples
img0,label0=test_dataset[0]
print(img0.shape,label0)
img1,label1=test_dataset[150]
print(img1.shape,label1)
print("*"*60)
print(test_dataset.classes)#list out all the classes

batch_size=32
# Training Part
test_ds=DataLoader(test_dataset,batch_size)

right_test = 0
total_test = 0
with torch.no_grad():
    for data in test_ds:
        images, labels_test = data
        images,labels_test=images.to(device),labels_test.to(device)
        #images,labels_test=images,labels_test
        outputs_test = model(images)
        _, predicted_test = torch.max(outputs_test, 1)
        total_test += labels_test.size(0)
        right_test += (predicted_test == labels_test).sum()

print('Accuracy on the Test images: %d %%' % (
    100 * right_test / total_test))

