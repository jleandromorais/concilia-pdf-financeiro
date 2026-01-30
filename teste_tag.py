from pathlib import Path
import pdfplumber
from tkinter import filedialog, Tk

# Esconde a janela principal do Tkinter
root = Tk()
root.withdraw()

print("Selecione UM arquivo PDF da TAG (aqueles que começam com OAC...)")
arquivo = filedialog.askopenfilename(title="Selecione um PDF da TAG", filetypes=[("PDF", "*.pdf")])

if arquivo:
    print(f"\n--- LENDO ARQUIVO: {Path(arquivo).name} ---")
    try:
        with pdfplumber.open(arquivo) as pdf:
            # Lê a primeira página
            texto = pdf.pages[0].extract_text()
            print("--- INÍCIO DO TEXTO ---")
            print(texto)
            print("--- FIM DO TEXTO ---")
            
            print("\n\n>>> COPIE O TEXTO ACIMA E MANDE NO CHAT <<<")
    except Exception as e:
        print(f"Erro ao ler: {e}")
else:
    print("Nenhum arquivo selecionado.")