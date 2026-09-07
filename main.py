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
# print(array2 * array2)
# print("-"*30)
# print(array2.T)
# print("-"*30)
# print(array2 @ array2.T) 
# #print(1 + array2) 
# print("-"*30)
# print(array2.sum(axis=0)) # Referece ao eixo da linhas 
# print("-"*30)
# print(array2.sum(axis=1)) # Referece ao eixo da colunas
# print("-"*30)
# print(array2.sum()) # Soma tudo
# print("-"*30)
'''
Aqui está o que cada um faz:

    inicio (start): É a posição onde o seu corte começa. Este número é inclusivo, ou seja, o elemento desta exata posição entra no resultado. 
    Se você não colocar nada (ex: [:5]), o Python entende que deve começar do zero.

    fim (stop): É a posição onde o corte deve parar. Este número é exclusivo, ou seja, o Python para de cortarum item antes dessa posição 
    (ele não entra no resultado final). Se você não colocar nada (ex: [2:]), o Python vai até o último elemento.

    passo (step): É o intervalo ou o "pulo" entre os itens. O padrão é 1 (pegar de um em um). Se você colocar 2, 
    ele pega um, pula um, pega outro. Se você usar um número negativo (como -1), ele lê a sequência de trás para frente.

    matriz[inicio:fim:passo, inicio:fim:passo]
'''

print(array2)
print("-"*30)
print(array2[1, :2])
print("-"*30)
print(array2[:, 1:])
print("-"*30)
print(array2[:, 1])

print("-"*30)

print(np.arange(2 * 5).reshape(5, 2))
print("-"*30)
print(np.random.rand(2, 3))