import pyautogui as pygui
import openpyxl
import time

# Carregar a planilha
planilha_clientes = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = planilha_clientes['Sheet1']

# Abrir o site
pygui.click(523,1052, duration=0.2)  # Ajuste a posição conforme necessário
pygui.write(r"C:\Users\marco\Desktop\codes\automatizacao\index.htm")
pygui.press('enter')

# Adicionar um atraso para garantir que a página carregue
time.sleep(1)  # Ajuste o tempo conforme necessário

# Iterar sobre as linhas da planilha e preencher o formulário
for linha in pagina_clientes.iter_rows(min_row=2, max_row=3, values_only=True):
    if linha:  # Verificar se a linha não está vazia
        nome, telefone, email, endereco, cidade, estado, cep, data_nascimento = linha

        # Preencher o formulário
        pygui.click(680,233, duration=0.4)  # Ajuste a posição do primeiro campo
        pygui.write(nome)
        pygui.press('tab')
        pygui.write(telefone)
        pygui.press('tab')
        pygui.write(email)
        pygui.press('tab')
        pygui.write(endereco)
        pygui.press('tab')
        pygui.write(cidade)
        pygui.press('tab')
        pygui.write(estado)
        pygui.press('tab')
        pygui.write(cep)
        pygui.press('tab')
        data = data_nascimento.replace('/', '')
        pygui.write(data, interval=1)
        
        # Adicionar um atraso para garantir que o formulário seja preenchido corretamente
        time.sleep(1)
        
        # Enviar o formulário
        pygui.click(698,774)

        #print(f"Cadastro enviado: {nome}, {telefone}, {email}, {endereco}, {cidade}, {estado}, {cep}, {data_nascimento}")
    else:
        print("Linha vazia encontrada")

print("Todos os cadastros foram enviados.")