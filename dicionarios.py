dicionario = {
    "name" : "nome",
    "see": "olhe",
    "love": "amor"
     
}
dicionary = {
    "mensagem" : "text",
    "flime" : "movie"
}
#dic_midia = {**dicionario, **dicionary}#Descompaquitamenento com **
dic_midia = dicionario | dicionary# pipe como operador de mesclagem, python3.9 pra cima
print(dic_midia.values())


importante = False
print("importante") if importante else print("náo e imporante")#operadore ternario