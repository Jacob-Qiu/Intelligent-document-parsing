"""
    第一版本训练器
"""


import joblib
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, recall_score, precision_score, f1_score


def train(feature_path: str, save_path: str, model_name: str) -> None:
    """基于特征数据集训练随机森林分类器

    Args:
        path·(str): 特征数据集(.csv)的路径
        save_path(str): 模型文件保存的路径
        model_name·(str): 模型文件名.pkl

    Return:
        None
   """
    df = pd.read_csv(feature_path)
    df.drop(columns=df.columns[0], axis=1, inplace=True)
    x = df.drop(columns=['Text', 'Style'])
    y = df['Style']

    # 获取训练集和验证集
    x_train, x_test, y_train, y_test = train_test_split(x, y, test_size=0.1, random_state=42, stratify=y)

    # 创建随机森林分类器
    average = None
    clf = RandomForestClassifier(max_depth=10, random_state=2, class_weight='balanced')
    clf.fit(x_train, y_train)

    # 评价模型性能
    train_acc = accuracy_score(y_train, clf.predict(x_train))  # 训练集准确性
    test_acc = accuracy_score(y_test, clf.predict(x_test))  # 验证集准确性
    train_rec = recall_score(y_train, clf.predict(x_train), average=average)
    test_rec = recall_score(y_test, clf.predict(x_test), average=average)
    train_pre = precision_score(y_train, clf.predict(x_train), average=average)
    test_pre = precision_score(y_test, clf.predict(x_test), average=average)
    train_f1 = f1_score(y_train, clf.predict(x_train), average=average)
    test_f1 = f1_score(y_test, clf.predict(x_test), average=average)
    feature_imp = pd.Series(clf.feature_importances_, index=x_train.columns).sort_values(ascending=False)  # 各特征的重要性
    print(f'train_acc:{train_acc}')
    print(f'test_acc:{test_acc}\n')
    print(f'train_rec:{train_rec}')
    print(f'test_rec:{test_rec}\n')
    print(f'train_pre:{train_pre}')
    print(f'test_pre:{test_pre}\n')
    print(f'train_f1:{train_f1}')
    print(f'test_f1:{test_f1}\n')
    print(feature_imp)

    # 存储模型文件
    file_path = save_path + '/' + model_name
    joblib.dump(clf, open(file_path, "wb"))


# 示例
if __name__ == '__main__':
    feature_path = './feature/path'
    save_path = './model/save/path'
    model_name = 'model_name.pkl'
    train(feature_path, save_path, model_name)
