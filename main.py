import temperature as tp

resultado= tp.convert_to_celsius(120,"fahrenheit")
print(resultado)


import datetime as dt
timestamp = dt.datetime(2026, 9, 6, 19, 15)
print(timestamp.day )


import numpy as np

array1 = np.array([10, 100, 1000.])

array2 = np.array([[1., 2., 3],
                   [4., 5., 6.]])
print(array2 * array2)
print("-"*30)
print(array2.T)
print("-"*30)
print(array2 @ array2.T) 
#print(1 + array2) 
print("-"*30)
print(array2.sum(axis=0)) # Referece ao eixo da linhas 
print("-"*30)
print(array2.sum(axis=1)) # Referece ao eixo da colunas
print("-"*30)
print(array2.sum()) # Soma tudo
print("-"*30)