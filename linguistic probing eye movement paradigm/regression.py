import sklearn
from sklearn.linear_model import LinearRegression as LR,Ridge,Lasso,RidgeCV
from sklearn.model_selection import train_test_split
from sklearn.model_selection import cross_val_score
from sklearn.metrics import mean_squared_error as MSE, r2_score
from sklearn.datasets import fetch_california_housing as feth
import pandas as pd
import numpy as np
from sklearn.metrics import mean_squared_error

def RR(Xtrain, Xtest, Ytrain, Ytest,X,y):
    Ridge_ = RidgeCV(alphas=np.arange(0.1, 100, 0.2)
                     , scoring="neg_mean_squared_error"
                     , store_cv_values=True
                     # ,cv=5
                     ).fit(Xtrain, Ytrain)
    # 无关交叉验证的岭回归结果
    print(Ridge_.score(Xtrain, Ytrain))
    # 调用所有交叉验证的结果
    # print(Ridge_.cv_values_.shape)
    # 进行平均后可以查看每个正则化系数取值下的交叉验证结果
    # print(Ridge_.cv_values_.mean(axis=0))
    # print(Ridge_.cv_values_)
    # 查看被选择出来的最佳正则化系数
    # print('L2:',Ridge_.alpha_)

    # 岭回归
    reg = Ridge(alpha=Ridge_.alpha_).fit(Xtrain, Ytrain)  # alpha为正则项系数，等于0相当于没有惩罚，和线性回归一样的

    yhat = reg.predict(Xtest)  # 测值
    print([*zip(Xtrain.columns, reg.coef_)])  # 系数
    print('b:',reg.intercept_)  # 截距
    # 模型评估
    mean = mean_squared_error(Ytest, yhat)
    rmse = np.sqrt(mean_squared_error(Ytest, yhat))
    print('RMSE:',rmse)
    return rmse
    # print(MSE(yhat, Ytest))  # 均方误差MSE

'''
    number_of_characters
    is_entity_critical_word	
    number_of_dominated_nodes	
    complexity_score	
    max_dependency_distance	
    number_of_senses_in_wordnet	
    
    avg_word_first_duration	avg_word_go_past_time	avg_word_total_reading_time	pos

'''

def get_rmse(dataset_path,label,train_list):
    train=pd.read_excel(dataset_path)
    # train.dropna(inplace=True)
    y = train[label]
    X = train[train_list]
    Xtrain, Xtest, Ytrain, Ytest = train_test_split(X, y, test_size=0.1, random_state=420)
    print(RR(Xtrain, Xtest, Ytrain, Ytest,X,y))

label=['avg_word_first_duration','avg_word_total_reading_time']
# FFD_list=['number_of_characters', 'start_with_capital_letter','have_alphanumeric_letters','capital_letters_only']
# TRT_list=['is_entity_critical_word','number_of_dominated_nodes','complexity_score', 'max_dependency_distance', 'number_of_senses_in_wordnet']
FFD_list=['number_of_characters', 'start_with_capital_letter','is_entity_critical_word', 'number_of_senses_in_wordnet','sentiment']
TRT_list=['number_of_characters', 'start_with_capital_letter','is_entity_critical_word', 'number_of_senses_in_wordnet','sentiment','number_of_dominated_nodes_norm','complexity_score_norm', 'max_dependency_distance_norm']

print('Provo avg_word_first_duration')
get_rmse('update_data/Provo_2.xlsx',label[0],FFD_list)
print('Provo avg_word_total_reading_time')
get_rmse('update_data/Provo_2.xlsx',label[1],TRT_list)
print('===========================')
print('GECO avg_word_first_duration')
get_rmse('update_data/GECO_2.xlsx',label[0],FFD_list)
print('GECO avg_word_total_reading_time')
get_rmse('update_data/GECO_2.xlsx',label[1],TRT_list)
print('===========================')
print('EA avg_word_first_duration')
get_rmse('update_data/EA_3.xlsx',label[0],FFD_list)
print('EA avg_word_total_reading_time')
get_rmse('update_data/EA_3.xlsx',label[1],TRT_list)
# get_rmse('Gaze_data/EA_avg.xlsx',label)
# label='avg_word_total_reading_time'
# get_rmse('Gaze_data/clean_GECO_avg_4.xlsx',label)
# 'avg_word_first_duration'
# 'avg_word_regression_time'
# 'avg_word_total_reading_time'
# 'avg_word_go_past_time'