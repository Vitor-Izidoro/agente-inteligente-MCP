import speech_recognition as sr
import pyttsx3
import openpyxl
import requests
import langchain
import pyaudio
# Se usar Ollama:
import ollama
import requests
import webbrowser

# Função para criar/abrir planilha de produtos (agora com coluna preço)
def carregar_planilha(nome_arquivo="lista_compras.xlsx"):
    try:
        wb = openpyxl.load_workbook(nome_arquivo)
        ws = wb["Produtos"]
        # Garante que as colunas existam
        if ws.max_row == 0 or ws.cell(row=1, column=1).value != "Produto":
            ws.append(["Produto", "Preço"])
            wb.save(nome_arquivo)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Produtos"
        ws.append(["Produto", "Preço"])
        wb.save(nome_arquivo)
    return wb

# Função para adicionar produto com preço opcional
def adicionar_produto(produto, preco=None, nome_arquivo="lista_compras.xlsx"):
    wb = carregar_planilha(nome_arquivo)
    ws = wb["Produtos"]
    if preco is not None:
        ws.append([produto, preco])
    else:
        ws.append([produto, ""])
    wb.save(nome_arquivo)

# Função para listar produtos e preços
def listar_produtos(nome_arquivo="lista_compras.xlsx"):
    wb = carregar_planilha(nome_arquivo)
    ws = wb["Produtos"]
    produtos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            produtos.append(f"{row[0]} - R$ {row[1] if row[1] else 'N/A'}")
    return produtos

# Função para remover produto (por nome)
def remover_produto(produto, nome_arquivo="lista_compras.xlsx"):
    wb = carregar_planilha(nome_arquivo)
    ws = wb["Produtos"]
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] and produto.lower() == row[0].lower():
            ws.delete_rows(idx)
            wb.save(nome_arquivo)
            break

# Função para gerar link WhatsApp (mostra nome e preço)
def gerar_link_whatsapp(produtos):
    texto = "Lista de compras: " + ", ".join(produtos)
    numero = "5541988039241"
    link = f"https://wa.me/{numero}?text={requests.utils.quote(texto)}"
    webbrowser.open(link)
    return link

# Função para reconhecer comando de voz
def reconhecer_comando():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Fale seu comando...")
        audio = r.listen(source)
    try:
        comando = r.recognize_google(audio, language="pt-BR")
        print(f"Você disse: {comando}")
        return comando.lower()
    except sr.UnknownValueError:
        print("Não entendi o comando.")
        return ""
    except sr.RequestError:
        print("Erro ao acessar o serviço de reconhecimento.")
        return ""

# Função para reconhecer palavra de ativação

def ouvir_ate_ativacao(wake_word="agente"):
    r = sr.Recognizer()
    while True:
        with sr.Microphone() as source:
            print(f"Diga '{wake_word}' para ativar o agente...")
            audio = r.listen(source)
        try:
            texto = r.recognize_google(audio, language="pt-BR")
            print(f"Você disse: {texto}")
            if wake_word in texto.lower():
                print("agente ativada! Fale seu comando...")
                return True
        except sr.UnknownValueError:
            print("Não entendi. Tente novamente.")
        except sr.RequestError:
            print("Erro ao acessar o serviço de reconhecimento.")

# Função principal (exemplo de fluxo básico)
def main():
    while True:
        ouvir_ate_ativacao()
        comando = reconhecer_comando()
        if "adicionar" in comando:
            dados = comando.replace("adicionar", "").strip().split()
            produto = None
            preco = None
            # Tenta encontrar um valor que seja preço (float)
            for i, item in enumerate(dados):
                try:
                    valor = float(item.replace(",", "."))
                    preco = valor
                    produto = " ".join(dados[:i])
                    break
                except ValueError:
                    pass
            if produto is None:
                produto = " ".join(dados)
            if produto:
                adicionar_produto(produto, preco)
                if preco is not None:
                    print(f"Produto '{produto}' adicionado com preço R$ {preco}.")
                else:
                    print(f"Produto '{produto}' adicionado sem preço.")
        elif "remover" in comando:
            produto = comando.replace("remover", "").strip()
            if produto:
                remover_produto(produto)
                print(f"Produto '{produto}' removido.")
        elif "listar" in comando:
            produtos = listar_produtos()
            print("Produtos:", produtos)
        elif "enviar" in comando:
            produtos = listar_produtos()
            link = gerar_link_whatsapp(produtos)
            print("Link para WhatsApp:", link)
        elif "sair" in comando:
            print("Encerrando agente.")
            break
        else:
            print("Comando não reconhecido. Tente novamente.")
def gerar_link_whatsapp(produtos):
    texto = "Lista de compras: " + ", ".join(produtos)
    numero = "5541988039241"
    link = f"https://wa.me/{numero}?text={requests.utils.quote(texto)}"
    webbrowser.open(link)
    return f"https://wa.me/{numero}?text={requests.utils.quote(texto)}"
if __name__ == "__main__":
    main()