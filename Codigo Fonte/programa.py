import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import webbrowser
from io import StringIO

def selecionar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(title="Selecione o arquivo Excel",
                                                 filetypes=[("Excel files", "*.xlsx *.xls")])
    if caminho_arquivo:
        carregar_dados(caminho_arquivo)

def carregar_dados(caminho):
    try:
        df = pd.read_excel(caminho, na_filter=False)
        exibir_prompt(df)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")

def exibir_prompt(df):
    def abrir_chatgpt():
        prompt_usuario = entrada_prompt.get("1.0", tk.END).strip()
        dados_excel = "\n".join(["\t".join(str(x) if pd.notna(x) else "" for x in row) for row in df.values])
        prompt_completo = f"{prompt_usuario}\nFaça com que todas as informações sejam organizadas, mesmo que deixe a planilha desigual, e preencha os espaços desiguais com os caracteres  !!.!! \n\n{dados_excel}"

        # Copiar o prompt completo para a área de transferência
        root.clipboard_clear()
        root.clipboard_append(prompt_completo)
        messagebox.showinfo("Atenção", "O prompt foi copiado para a área de transferência. Abra o ChatGPT, cole o prompt e obtenha a planilha organizada.")

        # Abrir ChatGPT no navegador
        webbrowser.open("https://chat.openai.com/")

        # Abre um prompt pedindo que o usuário cole a planilha organizada manualmente
        colar_planilha_janela = tk.Toplevel(root)
        colar_planilha_janela.title("Cole a Planilha Organizada")

        label_instrucoes = tk.Label(colar_planilha_janela, text="Cole a planilha organizada pelo ChatGPT abaixo e clique em 'Salvar'")
        label_instrucoes.pack()

        entrada_planilha = tk.Text(colar_planilha_janela, height=10, width=60)
        entrada_planilha.pack()

        botao_salvar = tk.Button(colar_planilha_janela, text="Salvar Planilha", command=lambda: salvar_arquivo(entrada_planilha.get("1.0", tk.END), df))
        botao_salvar.pack()

    prompt_janela = tk.Toplevel(root)
    prompt_janela.title("Instruções de Organização")
    prompt_label = tk.Label(prompt_janela, text="Digite como deseja organizar a planilha:")
    prompt_label.pack()

    entrada_prompt = tk.Text(prompt_janela, height=5, width=50)
    entrada_prompt.insert(tk.END, "")  # Não insere texto inicial agora
    entrada_prompt.pack()

    botao_enviar = tk.Button(prompt_janela, text="Enviar para ChatGPT", command=abrir_chatgpt)
    botao_enviar.pack()

def salvar_arquivo(texto_planilha, df):
    # Usa StringIO para transformar o texto colado em um DataFrame e salvar como Excel
    try:
        dados = StringIO(texto_planilha)
        df_novo = pd.read_csv(dados, sep="\t")  # Assumindo que o ChatGPT organiza os dados em formato tabulado

        # Limpa as células que contêm os caracteres '/!!.!!\' para manter a estrutura
        df_novo.replace("!!.!!", "", inplace=True)

        # Abre a tela para o usuário salvar o arquivo
        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if caminho:
            df_novo.to_excel(caminho, index=False)  # Salva diretamente em formato Excel
            messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")

root = tk.Tk()
root.title("Organizador de Planilhas com ChatGPT")

botao_selecionar = tk.Button(root, text="Selecionar Arquivo Excel", command=selecionar_arquivo)
botao_selecionar.pack(pady=20)

root.mainloop()
