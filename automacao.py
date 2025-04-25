import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import unicodedata

def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(c for c in texto if not unicodedata.combining(c))
    texto = texto.replace(" ", "")  # remove todos os espaços (até os entre letras)
    return texto


def encontrar_header(caminho_arquivo):
    for i in range(15):  # aumentamos de 10 para 15
        try:
            df_temp = pd.read_excel(caminho_arquivo, header=i, nrows=1)
            colunas = [normalizar(col) for col in df_temp.columns]

            if any("produto" in col or "descricao" in col for col in colunas) and \
               any("preco" in col or "valor" in col for col in colunas):
                return i
        except Exception:
            continue
    raise Exception("Cabeçalho não encontrado nas primeiras 15 linhas.")



def padronizar_colunas(caminho_arquivo):
    try:
        # Lê o arquivo com o cabeçalho na linha correta
        header_index = encontrar_header(caminho_arquivo)
        df = pd.read_excel(caminho_arquivo, header=header_index)


        colunas_originais = list(df.columns)
        colunas_normalizadas = [normalizar(col) for col in colunas_originais]

        mapa = {
    "produto": ["produto", "item", "nomedoproduto", "descricao", "descricaodoitem", "produtobruto", "produtonatural"],
    "preco": ["preco", "valor", "precounitario", "r$/kg", "peso/un", "r/kg", "preco1kg", "preco10kg"]
}



        colunas_padrao = {}

        for chave, variacoes in mapa.items():
            for var in variacoes:
                for original, normalizada in zip(colunas_originais, colunas_normalizadas):
                    if var in normalizada:
                        colunas_padrao[chave] = original
                        break
                if chave in colunas_padrao:
                    break

        if not all(k in colunas_padrao for k in ["produto", "preco"]):
            raise Exception("Colunas obrigatórias não foram encontradas.")

        df_renomeado = df.rename(columns={
            colunas_padrao["produto"]: "Produto",
            colunas_padrao["preco"]: "Preço"
        })

        

     # Extrai nome da empresa
        nome_empresa = os.path.basename(caminho_arquivo).split('.')[0].upper()
        df_renomeado["Empresa"] = nome_empresa

        return df_renomeado[["Produto", "Preço", "Empresa"]]

    except Exception as e:
        raise Exception(f"Erro no arquivo: {caminho_arquivo}\n{str(e)}")



# Função para selecionar os arquivos
def selecionar_planilhas():
    arquivos = filedialog.askopenfilenames(
        title="Selecione as planilhas",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivos:
        processar_planilhas(arquivos)

# Função que processa todas as planilhas
def processar_planilhas(arquivos):
    todas = []
    for arquivo in arquivos:
        try:
            df_padronizado = padronizar_colunas(arquivo)
            todas.append(df_padronizado)
        except Exception as e:
            messagebox.showwarning("Erro", f"Erro no arquivo: {arquivo}\n{e}")

    if not todas:
        messagebox.showerror("Erro", "Nenhuma planilha válida foi processada.")
        return

    resultado = pd.concat(todas, ignore_index=True)
    resultado = resultado.dropna(subset=["Preço"])
    resultado["Preço"] = pd.to_numeric(resultado["Preço"], errors="coerce")
    resultado = resultado.dropna(subset=["Preço"])
    resultado_ordenado = resultado.sort_values(by=["Produto", "Preço"])

    resultado_ordenado.to_excel("cotacao_organizada.xlsx", index=False)
    messagebox.showinfo("Sucesso", "Planilha 'cotacao_organizada.xlsx' criada com sucesso!")


# Interface gráfica
janela = tk.Tk()
janela.title("Cotador Inteligente de Preços")
janela.geometry("400x200")

label = tk.Label(janela, text="Selecione as planilhas dos fornecedores:", font=("Arial", 12))
label.pack(pady=20)

botao = tk.Button(janela, text="Selecionar Planilhas", command=selecionar_planilhas,
                  font=("Arial", 12), bg="#007ACC", fg="white")
botao.pack(pady=10)

janela.mainloop()
