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

# Função para criar/abrir planilha de produtos
def carregar_planilha(nome_arquivo="lista_compras.xlsx"):
    try:
        wb = openpyxl.load_workbook(nome_arquivo)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        wb.active.title = "Produtos"
        wb.save(nome_arquivo)
    return wb

# Função para adicionar produto
def adicionar_produto(produto, nome_arquivo="lista_compras.xlsx"):
    wb = carregar_planilha(nome_arquivo)
    ws = wb["Produtos"]
    ws.append([produto])
    wb.save(nome_arquivo)

# Função para listar produtos
def listar_produtos(nome_arquivo="lista_compras.xlsx"):
    wb = carregar_planilha(nome_arquivo)
    ws = wb["Produtos"]
    return [row[0].value for row in ws.iter_rows() if row[0].value]

# Função para remover produto
def remover_produto(produto, nome_arquivo="lista_compras.xlsx"):
    wb = carregar_planilha(nome_arquivo)
    ws = wb["Produtos"]
    produtos = [row[0].value for row in ws.iter_rows() if row[0].value]
    if produto in produtos:
        idx = produtos.index(produto) + 1
        ws.delete_rows(idx)
        wb.save(nome_arquivo)

# Função para gerar link WhatsApp
def gerar_link_whatsapp(produtos):
    texto = "Lista de compras: " + ", ".join(produtos)
    return f"https://wa.me/?text={requests.utils.quote(texto)}"

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

# Função principal (exemplo de fluxo básico)
def main():
    while True:
        comando = reconhecer_comando()
        if "adicionar" in comando:
            produto = comando.replace("adicionar", "").strip()
            if produto:
                adicionar_produto(produto)
                print(f"Produto '{produto}' adicionado.")
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