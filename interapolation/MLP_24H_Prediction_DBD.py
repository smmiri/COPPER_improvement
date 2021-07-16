import numpy as np
from numpy.lib.financial import ipmt
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Dense, Activation, Dropout
from tensorflow.keras.utils import to_categorical, plot_model
from tensorflow.keras.datasets import mnist
import pandas as pd
import random
import tensorflow as tf


data = pd.read_excel ('exp.xlsx', engine='openpyxl')
df = pd.DataFrame(data, columns= ['DATE','VALUE'])

daily_values = df['VALUE'].to_numpy()
length_daily_values = len(daily_values)

daily_values_array = []
daily_values_label = []

for i in range(23,length_daily_values,24):

    if np.isnan(daily_values[i]):
        continue
    
    if np.isnan(daily_values[i:(i+24)]).any():
        continue
    daily_values_label.append(round(i/24))
    daily_values_array.append(daily_values[i:(i+24)])

print(len(daily_values_array))

daily_values_label = np.array(daily_values_label)

daily_values_array_train = []
daily_values_label_train = []

for i in range(0,50):
    n_index = random.randint(0,63)
    
    while daily_values_label[n_index] in daily_values_label_train:
        n_index = random.randint(0,63)

    daily_values_label_train.append(daily_values_label[n_index])
    daily_values_array_train.append(daily_values_array[n_index])

daily_values_label_test = []
daily_values_array_test = []
for i in range(0,63):
    if daily_values_label[i] in daily_values_label_train:
        continue
    daily_values_label_test.append(daily_values_label[i])
    daily_values_array_test.append(daily_values_array[i])

input_size = 1
output_size = 24
# network parameters
batch_size = 5
hidden_units = 128
dropout = 0.45

# model is a 3-layer MLP with ReLU and dropout after each layer
model = Sequential()
model.add(Dense(hidden_units, input_dim=input_size))
model.add(Activation('relu'))
model.add(Dropout(dropout))
model.add(Dense(hidden_units))
model.add(Activation('relu'))
model.add(Dropout(dropout))
model.add(Dense(output_size))
model.add(Activation('relu'))
model.add(Dropout(dropout))
model.add(Dense(output_size))
# this is the output for one-hot vector

model.summary()
#plot_model(model, to_file='mlp-24h-prediction-dbd.png', show_shapes=True)

model.compile(loss='categorical_crossentropy',
              optimizer='adam',
              metrics=['accuracy'])

for i in range(0,50):
    print(daily_values_array_train[i])

'''
print(len(np.array(daily_values_label_train)))
print(type(np.array(daily_values_label_train)))

print(len(daily_values_array_train))
print(type(daily_values_array_train))
'''


model.fit(np.array(daily_values_label_train), np.array(daily_values_array_train), epochs=100, batch_size=5)

# validate the model on test dataset to determine generalization
loss, acc = model.evaluate(np.array(daily_values_label_test), np.array(daily_values_array_test), batch_size=batch_size)
print("\nTest accuracy: %.1f%%" % (100.0 * acc))


'''
print(len(daily_values_label_test))
print(len(set(daily_values_label_test)))

print(len(daily_values_label))
print(len(daily_values_array))
print((daily_values_label))
print(daily_values_label_train)
'''