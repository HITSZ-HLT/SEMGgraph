from sklearn.datasets import load_boston
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LinearRegression, SGDRegressor
from sklearn.metrics import mean_squared_error
from sklearn.linear_model import Ridge, ElasticNet, Lasso,RidgeCV
# Rideg表示岭回归，ElasticNet表示弹性网络，Lasso表示Lasso回归
import numpy as np

def linear_model1(x_train, x_test, y_train, y_test):
    """
    线性回归:正规方程
    :return:None
    """
    # 1.获取数据
    # data = load_boston()

    # 2.数据集划分
    # x_train, x_test, y_train, y_test = train_test_split(data.data, data.target, random_state=22)

    # 3.特征工程-标准化
    transfer = StandardScaler()
    x_train = transfer.fit_transform(x_train)
    x_test = transfer.fit_transform(x_test)

    # 4.机器学习-线性回归(正规方程)
    estimator = LinearRegression()
    estimator.fit(x_train, y_train)

    # 5.模型评估
    # 5.1 获取系数等值
    y_predict = estimator.predict(x_test)
    # print("预测值为:\n", y_predict)
    print("模型中的系数为:\n", estimator.coef_)
    print("模型中的偏置为:\n", estimator.intercept_)

    # 5.2 评价
    # 均方误差
    error = mean_squared_error(y_test, y_predict)
    print("误差为:\n", error)
    RMSE = np.sqrt(error)
    print('RMSE:', RMSE)
    return estimator.intercept_, estimator.coef_,error,RMSE


def linear_model2(x_train, x_test, y_train, y_test):
    """
    线性回归:梯度下降法
    :return:None
    """
    transfer = StandardScaler()
    x_train = transfer.fit_transform(x_train)
    x_test = transfer.fit_transform(x_test)

    # 4.机器学习-线性回归(特征方程)
    estimator = SGDRegressor(max_iter=1000)
    estimator.fit(x_train, y_train)

    # 5.模型评估
    # 5.1 获取系数等值
    y_predict = estimator.predict(x_test)
    # print("预测值为:\n", y_predict)
    print("模型中的系数为:\n", estimator.coef_)
    print("模型中的偏置为:\n", estimator.intercept_)

    error = mean_squared_error(y_test, y_predict)
    print("误差为:\n", error)
    RMSE = np.sqrt(error)
    print('RMSE:', RMSE)
    return estimator.intercept_, estimator.coef_,error,RMSE

def linear_model3(x_train, x_test, y_train, y_test):
    """
    线性回归:岭回归
    :return:
    """
    transfer = StandardScaler()
    x_train = transfer.fit_transform(x_train)
    x_test = transfer.fit_transform(x_test)

    Ridge_ = RidgeCV(alphas=np.arange(1, 1001, 2)
                     , scoring="neg_mean_squared_error"
                     , store_cv_values=True
                     # ,cv=5
                     ).fit(x_train, y_train)
    # 无关交叉验证的岭回归结果
    # print(Ridge_.score(x_train, y_train))

    # 4.机器学习-线性回归(岭回归)
    estimator = Ridge(alpha=Ridge_.alpha_).fit(x_train, y_train)
    estimator = RidgeCV(alphas=(0.1, 1, 10))
    estimator.fit(x_train, y_train)

    # 5.模型评估
    # 5.1 获取系数等值
    y_predict = estimator.predict(x_test)
    # print("预测值为:\n", y_predict)
    print("模型中的系数为:\n", estimator.coef_)
    print("模型中的偏置为:\n", estimator.intercept_)
    print('-------------------')
    print(type(estimator.intercept_))
    print(type(estimator.coef_))
    # 5.2 评价
    # 均方误差
    error = mean_squared_error(y_test, y_predict)
    print("误差为:\n", error)

    RMSE = np.sqrt(error)
    print('RMSE:', RMSE)
    return estimator.intercept_, estimator.coef_, error, RMSE


if __name__ == '__main__':
    import pandas as pd
    import xlwt
    wb=xlwt.Workbook()
    ws_1=wb.add_sheet('Provo')
    ws_2=wb.add_sheet('GECO')
    ws_3 = wb.add_sheet('EA')
    FFD_list=['number_of_characters', 'is_entity_critical_word','start_with_capital_letter','have_alphanumeric_letters','capital_letters_only']
    TRT_list=['number_of_dominated_nodes','complexity_score', 'max_dependency_distance', 'number_of_senses_in_wordnet']
    # labels = ['avg_word_first_duration','avg_word_go_past_time','avg_word_total_reading_time']
    labels = ['avg_word_first_duration']
    # dataset_path=['Gaze_data/Provo_avg.xlsx','Gaze_data/clean_GECO_avg_4.xlsx','Gaze_data/EA_avg.xlsx']
    dataset_path=['Provo.xlsx']
    for each_dataset in dataset_path:
        print('======='+each_dataset+'=======')
        train = pd.read_excel(each_dataset)
        # train.dropna(inplace=True)
        for i in range(len(labels)):
            print('======='+labels[i]+'=======')
            y = train[labels[i]]

            X = train[FFD_list]
            Xtrain, Xtest, Ytrain, Ytest = train_test_split(X, y, test_size=0.2, random_state=420)
            print('========方程式LR========')
            LR_intercept, LR_coef,LR_error,LR_RMSE=linear_model1(Xtrain, Xtest, Ytrain, Ytest)
            print('========梯度下降GR========')
            GR_intercept, GR_coef,GR_error,GR_RMSE=linear_model2(Xtrain, Xtest, Ytrain, Ytest)
            print('========岭回归RR========')
            RR_intercept, RR_coef,RR_error,RR_RMSE=linear_model3(Xtrain, Xtest, Ytrain, Ytest)

            if each_dataset=='Gaze_data/Provo_avg.xlsx':
                ws=ws_1
            elif each_dataset=='Gaze_data/clean_GECO_avg_4.xlsx':
                ws=ws_2
            else:
                ws=ws_3

            ws.write(0,i,labels[i])
            ws.write(1,i,'LR')
            ws.write(2, i, str(LR_intercept))
            ws.write(3,i,LR_coef[0])
            ws.write(4, i, LR_coef[1])
            ws.write(5, i, LR_coef[2])
            ws.write(6, i, LR_coef[3])
            ws.write(7, i, LR_coef[4])
            # ws.write(8, i, LR_coef[5])
            ws.write(9,i,LR_error)
            ws.write(10,i,LR_RMSE)


            ws.write(11,i,'GR')
            ws.write(12, i,str(GR_intercept))
            ws.write(13,i,GR_coef[0])
            ws.write(14, i, GR_coef[1])
            ws.write(15, i, GR_coef[2])
            ws.write(16, i, GR_coef[3])
            ws.write(17, i, GR_coef[4])
            #ws.write(18, i, GR_coef[5])
            ws.write(19,i,GR_error)
            ws.write(20,i,GR_RMSE)


            ws.write(21,i,'RR')
            ws.write(22, i, str(RR_intercept))
            ws.write(23,i,RR_coef[0])
            ws.write(24, i, RR_coef[1])
            ws.write(25, i, RR_coef[2])
            ws.write(26, i, RR_coef[3])
            ws.write(27, i, RR_coef[4])
            #ws.write(28, i, RR_coef[5])
            ws.write(29,i,RR_error)
            ws.write(30,i,RR_RMSE)
    wb.save('Gaze_data/param_1.xlsx')
