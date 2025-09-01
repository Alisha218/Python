import pandas as pd
#working on actual data 

dataframe=pd.read_csv("student_mark.csv")
#checking first few rows
print("First few rows")
print(dataframe.head)
#checking for missing values 
print("Missing values ")
print(dataframe.isnull().sum())
'''The 0 next to each column name indicates that all columns have integer values (int64).
dtype: int64 means all the data in the columns is of type int64, which is a 64-bit integer.'''


#dropping values 
print("Dropping vlaues ")
dataframe=dataframe.dropna()
#group data togather 
print("Grouping data togather")
grouped = dataframe.groupby("Student_ID").mean()
'''
.mean=avgmarks.
.sum() — sums the values in each group.
.count() — counts the non-null values in each group.
.max() — gives the maximum value in each group.
.min() — gives the minimum value in each group.'''
print(grouped)
