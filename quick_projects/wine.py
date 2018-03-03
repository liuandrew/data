# -*- coding: utf-8 -*-
"""
Created on Sat Mar  3 10:27:02 2018

@author: Andy
"""

import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn import preprocessing

from sklearn.pipeline import make_pipeline
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import GridSearchCV

from sklearn.metrics import mean_squared_error, r2_score
from sklearn.model_selection import cross_val_score
dataset_url = 'http://mlr.cs.umass.edu/ml/machine-learning-databases/wine-quality/winequality-red.csv'

data = pd.read_csv(dataset_url, sep=';')
y = data['quality']
X = data.drop('quality', axis=1)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, 
                                                    random_state=0, stratify=y)

scaler = preprocessing.StandardScaler()
scaler.fit(X_train)

pipeline = make_pipeline(preprocessing.StandardScaler(),
                         RandomForestRegressor(n_estimators=100))

hyperparameters = {
  'randomforestregressor__max_features': ['auto', 'sqrt', 'log2'],
  'randomforestregressor__max_depth': [None, 5, 3, 1]    
}

clf = GridSearchCV(pipeline, hyperparameters, cv=10)

y_pred = clf.predict(X_test)

print(r2_score(y_test, y_pred))
print(mean_squared_error(y_test, y_pred))


from sklearn.decomposition import PCA
import matplotlib.pyplot as plt
import seaborn as sns