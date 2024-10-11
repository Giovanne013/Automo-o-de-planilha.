#importa bibliotecas
import pandas as pd
import time
import pyautogui

# Lendo o arquivo CSV
tabela = pd.read_csv("Aulunosautomatização.csv",encoding="utf-8", delimiter=";")

# exibir tabela
print(tabela)

# remover colunas com drop(colums="nome da coluna") para dados que não irei ultilizar
tabela = tabela.drop(columns=["Matrícula","N°","Categoria","Telefone","ASO","CIP"])#ajusta corforme a necessidade
#exibbir colunas importantes
print(tabela)

# abrindo excel
pyautogui.press("win")
time.sleep(2)
pyautogui.write("excel")
time.sleep(2)
pyautogui.press("enter")
time.sleep(13)

#escrever colunas
pyautogui.click(x=2469,y=309)
time.sleep(2)

# montando tabela da planilha, das colunas na qual eu quero.
pyautogui.write("username")
pyautogui.press("tab")
time.sleep(1)

#tabela passowrd
pyautogui.write("password")
pyautogui.press("tab")
time.sleep(1)


#tabela firstname
pyautogui.write("firstname")
pyautogui.press("tab")
time.sleep(1)

#tabela lastname
pyautogui.write("lastname")
pyautogui.press("tab")
time.sleep(1)

#tabela email
pyautogui.write("email")
pyautogui.press("tab")
time.sleep(1)

#tabela qual o curso vai cadastrar na platforma Moodle.
pyautogui.write("course1")#CPTP
pyautogui.press("tab")
time.sleep(1)

#qual vai ser a função do cadastrado, no caso aqui vai ser estudante pq é uma planilha de estudantes.
pyautogui.write("role1")#student
pyautogui.press("tab")
time.sleep(1)


#aqui vai a turma na qual ele vai ficar cadastrado no moodle
pyautogui.write("group1")#TURMA - 2024
pyautogui.press("tab")
time.sleep(1)

#essa fica em branco pois não é necessaria
pyautogui.write("profile_field_CPF")
pyautogui.click(x=2468,y=334)#pegar a posição da primeira linha com o click

#pegar o index da tabela
linha = 0 #de onde o index vai começar na planilha no caso primeiro item. item 0 
#exemplo = 
# 0 JOSE AUGUSTO BATISTA RODRIGUES 2422883000 augustojaber@gmail.com (esse seria o index 0)
# 1 VALDELINO NUNES DA SILVA rodrigs 972145220 valdeilva48@hotmail.com ( index 1) e por ai vai !!!


for linha in tabela.index:
    username = tabela.loc[linha,"CPF"]
    pyautogui.write(str(username))
    pyautogui.press("tab")

    #localizando a primeira linha da coluna CPF
    username = tabela.loc[linha,"CPF"]
    pyautogui.write(str(username))
    pyautogui.press("tab")

    #localizando a primeira linha da coluna Nome
    fristname = tabela.loc[linha,"Nome"]
    pyautogui.write(str(fristname))
    pyautogui.press("tab")

    #localizando a primeira linha da coluna lastname
    lastname = tabela.loc[linha,"lastname"]
    pyautogui.write(str(lastname))
    pyautogui.press("tab")

    #localizando a primeira linha da coluna E-mail
    email = tabela.loc[linha,"E-mail"]
    pyautogui.write(str(email))
    pyautogui.press("tab")

    # escrevendo tipo do curso, caso o curso for outro, apenas alterar o curso aqui!! pelo nome real
    pyautogui.write("CAPTP")#CPTP
    pyautogui.press("tab")

    # definindo o tipo do cadastrado
    pyautogui.write("student")#student
    pyautogui.press("tab")

    # em caso de novas turmas, alterar o nome da turna aqui!! Ex:TURMA - 2025, TURMA - 2026 Exatamente igual ao que esta, apenas mucas o numero!!
    pyautogui.write("2024 - TURMA 25")#TURMA - 2024
    pyautogui.press("tab")

    #loop para retorna ao ínicio da planilha.
    username = tabela.loc[linha,"CPF"]
    pyautogui.write(str(username))
    pyautogui.press("left")
    pyautogui.press("left")
    pyautogui.press("left")
    pyautogui.press("left")
    pyautogui.press("left")
    pyautogui.press("left")
    pyautogui.press("left")
    pyautogui.press("left")
    pyautogui.press('down')
    # Pressiona a seta para a esquerda 5 vezes
    # repetir isso para todas as coluans e linhas.