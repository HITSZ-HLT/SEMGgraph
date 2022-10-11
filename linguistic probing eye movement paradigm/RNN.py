import  os
import  tensorflow as tf
import  numpy as np
from tensorflow import keras
# from general import *

#import  tensorflow as tf
#from tensorflow import keras
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import *
#from general import *
from sklearn.metrics import mean_squared_error, mean_absolute_error

tf.random.set_seed(22)
np.random.seed(22)
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'
assert tf.__version__.startswith('2.')

import numpy as np
import pandas as pd
import tensorflow as tf

# from tensorflow import keras
from tensorflow.keras import layers

from sklearn import datasets
from sklearn.model_selection import train_test_split


#生成回归训练集
def generate_regression_train_data(dataset_path,label,col_list):

    column_names = ['number_of_characters', 'is_entity_critical_word', 'number_of_dominated_nodes',
               'complexity_score', 'max_dependency_distance', 'number_of_senses_in_wordnet','avg_word_first_duration']
    # raw_dataset = pd.read_csv(dataset_path, skipinitialspace=True)

    # dataset = raw_dataset.copy()
    dataset = pd.read_excel(dataset_path, usecols=col_list)
    # dataset = dataset.dropna()
    # origin = dataset.pop('avg_word_first_duration')
    # dataset['USA'] = (origin == 1)*1.0
    # dataset['Europe'] = (origin == 2)*1.0
    # dataset['Japan'] = (origin == 3)*1.0

    train_dataset = dataset.sample(frac=0.8,random_state=0)
    test_dataset = dataset.drop(train_dataset.index)
    # print(train_dataset)
    # print('==========')

    train_stats = train_dataset.describe()
    train_stats.pop(label)
    train_stats = train_stats.transpose()
    y_train = train_dataset.pop(label)
    y_test = test_dataset.pop(label)
    def norm(x):
        return (x - train_stats['mean']) / train_stats['std']

    X_train = norm(train_dataset)
    X_test = norm(test_dataset)
    x_train = np.array(X_train)
    x_test = np.array(X_test)
    y_train = np.array(y_train)
    y_test = np.array(y_test)

    return x_train, y_train, x_test, y_test



#Func构建方式
class FuncMRegressor(keras.Model):
    def __init__(self, units, num_classes, num_layers):
        super(FuncMRegressor, self).__init__()

        self.cells = [keras.layers.LSTMCell(units) for _ in range(num_layers)]
        #
        self.rnn = keras.layers.RNN(self.cells, unroll=True)
        # self.rnn = keras.layers.RNN(units, return_sequences=True)
                        # return_sequences设置lstm中单个细胞中返回时间点的类型。非最后一层，return_sequences=True才能满足三维输入的数据格式要求
                        #默认 False。在输出序列中，返回单个 hidden state值还是返回全部time step 的 hidden state值。 False 返回单个， true 返回全部。

        # self.rnn2 = keras.layers.RNN(units) #最后一层，return_sequences=False，对于单个样本而言，返回的是一维（对于所有样本是二维）
        self.fc3 = keras.layers.Dense(1)  #最后一层全连接层。对于N分类问题，最后一层全连接输出个数为N个；对于回归问题，最后一层全连接层的输出为1

    #定义输入输出
    def call(self, inputs, training=None, mask=None):  ###*** 该函数重写了RNN的call()方法，该方法是模型输出表达式（即，模型）
        y = self.rnn(inputs)   #要求输入到lstm模型中的数据必须是三维
        #print("y2 shape:", y.shape)
        # y = self.rnn2(y)
        y = self.fc3(y)
        return y

#Seque构建方式（推荐）
class SequeRegressor():
    def __init__(self, units):
        self.units = units
        self.model = None

    #构建神经网络模型：（根据各层输入输出的shape）搭建网络结构、确定损失函数、确定优化器
    def build_model(self, loss, optimizer, metrics):
        self.model = Sequential()
        self.model.add(RNN(self.units, return_sequences=True))
        self.model.add(RNN(self.units))
        self.model.add(Dense(1))

        self.model.compile(loss=loss,
              optimizer=optimizer,   #优化器的选择很重要
              metrics=metrics)



def lstm_1(dataset_path,label,col_list):
    print("**********************【Func搭建方式】********************")
    # 获取二维数组（特征提取过的）
    x_train, y_train, x_test, y_test = generate_regression_train_data(dataset_path,label,col_list)
    x_train = x_train[:, :, np.newaxis]
    x_test = x_test[:, :, np.newaxis]

    #训练模型
    units = 64   #细胞个数
    num_classes = 1  #回归问题，该参数为1
    batch_size = 32
    epochs = 35
    model = FuncMRegressor(units, num_classes, num_layers=2)
    optimizer = tf.keras.optimizers.RMSprop(0.001)

    model.compile(loss='mse',
                optimizer=optimizer,
                metrics=['mae', 'mse'])


    # model.compile(optimizer=keras.optimizers.Adam(0.001),
    #                   loss=keras.losses.BinaryCrossentropy(from_logits=False),
    #                   metrics=['accuracy', 'mse'])
    model.fit(x_train, y_train, batch_size=batch_size, epochs=epochs,
                  validation_data=(x_test, y_test), verbose=1)

    # 模型应用预测
    #model.predict只返回y_pred
    out = model.predict(x_train)
    y_pred = model.predict(x_test)
    RMSE=np.sqrt(mean_squared_error(y_test, y_pred))
    print('RMSE:',RMSE)
    #print("out:", out)
    #evaluate用于评估您训练的模型。它的输出是准确度或损失，而不是对输入数据的预测。
    scores = model.evaluate(x_test, y_test, batch_size, verbose=1)
    print("Final test loss and accuracy :", scores)


label='avg_word_first_duration'
# Provo
print('==========provo============')
col_list=[4, 8, 9, 10, 11, 12, 13] # 13 first 14 pass 15 total
lstm_1('Gaze_data/Provo_avg.xlsx',label,col_list)
# GECO
print('==========geco============')
col_list=[4, 8, 9, 10, 11, 12, 15] # 15 first 16 pass 17 total
lstm_1('Gaze_data/clean_GECO_avg_4.xlsx',label,col_list)
# EA
print('==========ea============')
col_list=[4,5,6,7, 8, 9, 10] # 10 first 11 total 12 pass
lstm_1('Gaze_data/EA_avg.xlsx',label,col_list)

