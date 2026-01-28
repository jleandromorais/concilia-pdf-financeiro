from __future__ import annotations

import logging
import os
import re
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple
from tkinter import filedialog, messagebox

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

# ==========================================
# CONFIGURAÇÃO DA "IA" DE LEITURA (OCR)
# ==========================================
# Aqui está o caminho exato do seu usuário que você encontrou antes
CAMINHO_TESSERACT_JOSE = r"C:\Users\jose.demorais\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

OCR_ATIVADO = False
try:
    import pytesseract
    # Tenta configurar o caminho
    if os.path.exists(CAMINHO_TESSERACT_JOSE):
        pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESSERACT_JOSE
        OCR_ATIVADO = True
    else:
        # Se não achar no seu user, tenta nos padrões do Windows
        caminhos_padrao = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
        ]
        for p in caminhos_padrao:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                OCR_ATIVADO = True
                break
except ImportError:
    OCR_ATIVADO = False

APP_TITLE = "ConciliaPDF — Artilharia Pesada (OCR + Regex)"
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s")

@dataclass(frozen=True)
class PdfItem:
    file_name: str
    file_path: str
    category: str
    amount: float
    status: str
    method: str

def br_money_to_float(raw: str) -> float:
    if not raw: return 0.0
    # Mantém apenas números, vírgulas e pontos
    clean = re.sub(r"[^\d,\.]", "", str(raw))
    if not clean: return 0.0
    # Remove ponto de milhar e troca vírgula decimal
    clean = clean.replace(".", "").replace(",", ".")
    try:
        return float(clean)
    except ValueError:
        return 0.0

def format_br(value: float) -> str:
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def clean_ocr_text(text: str) -> str:
    """Limpa erros comuns de leitura de OCR"""
    t = text.replace("|", "").replace("!", "1").replace("l", "1")
    t = t.replace("$=", " ").replace("=", " = ")
    return t

# ==========================================
# FUNÇÃO DE LEITURA (HÍBRIDA)
# ==========================================
def ler_conteudo_pdf(pdf_path: Path) -> Tuple[str, str]:
    """
    Tenta ler texto digital. Se falhar ou for pouco, usa OCR (Tesseract) na imagem.
    """
    texto_final = ""
    origem = ""

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 1. Tenta extrair texto normal (rápido)
            paginas_texto = []
            for p in pdf.pages:
                paginas_texto.append(p.extract_text() or "")
            texto_digital = "\n".join(paginas_texto)

            # Se tiver bastante texto, confiamos nele
            if len(texto_digital.strip()) > 50:
                return texto_digital, "TEXTO_DIGITAL"

            # 2. Se chegou aqui, o PDF é imagem (Scan). Vamos usar o OCR.
            if not OCR_ATIVADO:
                return "", "FALHA: É imagem e Tesseract não foi achado."

            origem = "OCR (IA Visual)"
            paginas_ocr = []
            # Lê apenas a 1ª página (geralmente o total está nela) para ser rápido
            for i, p in enumerate(pdf.pages):
                if i > 0: break # Limita a 1 página
                # Aumenta resolução para ler letras pequenas
                imagem = p.to_image(resolution=300).original 
                texto_lido = pytesseract.image_to_string(imagem, lang="por")
                paginas_ocr.append(texto_lido)
            
            texto_final = "\n".join(paginas_ocr)
            
    except Exception as e:
        return "", f"ERRO LEITURA: {str(e)}"

    return texto_final, origem

# ==========================================
# CÉREBRO: PROCURAR VALORES
# ==========================================
def extrair_valor(text: str) -> Tuple[float, str]:
    # Limpa o texto para facilitar a busca
    text = clean_ocr_text(text)
    
    # --- REGRA 1: AMBEV / CBA (Receitas) ---
    # Padrão: "Total ... = R$ 142.000,00" ou "Total ... 142.000,00"
    # Procura a palavra "Total" seguida (mesmo que longe) por um "=" e um número
    match_receita = re.search(r"Total.*?=\s*(?:R\$\s*)?([\d\.]+(?:,\d{2}))", text, re.IGNORECASE | re.DOTALL)
    if match_receita:
        val = br_money_to_float(match_receita.group(1))
        return val, "Padrão 'Total ... ='"

    # --- REGRA 2: ENEVA / PETROLINA (Despesas) ---
    # Padrão CSV ou Nota de Débito
    # Procura "Valor Total Débito" ou "Líquido"
    match_despesa = re.search(r"(?:Valor\s+Total|L[ií]quido|Total\s+a\s+Pagar).*?([\d\.]+(?:,\d{2}))", text, re.IGNORECASE | re.DOTALL)
    if match_despesa:
        val = br_money_to_float(match_despesa.group(1))
        return val, "Padrão 'Valor Total/Líquido'"

    # --- REGRA 3: SALVA-VIDAS (Maior Valor) ---
    # Se o texto foi lido, mas a regra específica falhou, pega o maior valor monetário encontrado.
    # Isso ajuda muito em OCRs que comem palavras mas leem números.
    todos_valores = re.findall(r"(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})", text)
    if todos_valores:
        lista_floats = []
        for v in todos_valores:
            f = br_money_to_float(v)
            # Ignora valores pequenos (unitários) ou datas (2025)
            if f > 50 and f != 2025 and f != 2026:
                lista_floats.append(f)
        
        if lista_floats:
            return max(lista_floats), "Maior Valor Encontrado (Automático)"

    return 0.0, "Valor não identificado"

# ==========================================
# LÓGICA DE PROCESSAMENTO
# ==========================================
def processar_lista(arquivos: List[Path], categoria: str) -> List[PdfItem]:
    resultados = []
    for arquivo in arquivos:
        try:
            # 1. Ler (Texto ou OCR)
            texto, metodo_leitura = ler_conteudo_pdf(arquivo)
            
            if not texto:
                resultados.append(PdfItem(arquivo.name, str(arquivo), categoria, 0.0, "REVISAR", metodo_leitura))
                continue

            # 2. Extrair Valor
            valor, metodo_extracao = extrair_valor(texto)
            
            if valor > 0:
                status = "OK"
                detalhe = f"{metodo_leitura} -> {metodo_extracao}"
            else:
                status = "REVISAR"
                detalhe = f"{metodo_leitura} -> Texto lido, mas sem valor claro."

            resultados.append(PdfItem(arquivo.name, str(arquivo), categoria, valor, status, detalhe))

        except Exception as e:
            resultados.append(PdfItem(arquivo.name, str(arquivo), categoria, 0.0, "ERRO", str(e)))
            
    return resultados

def salvar_excel(caminho: Path, itens: List[PdfItem]):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatorio"
    ws.append(["Arquivo", "Categoria", "Valor", "Status", "Método", "Caminho"])
    
    for i in itens:
        ws.append([
            i.file_name, i.category, f"R$ {format_br(i.amount)}",
            i.status, i.method, i.file_path
        ])
    
    # Formatação bonita
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="003366") # Azul escuro
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["E"].width = 40
    
    wb.save(caminho)

# ==========================================
# INTERFACE GRÁFICA
# ==========================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("650x500")
        self.path_rec = None
        self.path_desp = None

        # Cabeçalho
        tk.Label(self, text="Conciliador com OCR (Leitura Visual)", font=("Arial", 14, "bold"), fg="#333").pack(pady=15)
        
        # Status do Tesseract
        if OCR_ATIVADO:
            status_ocr = "✅ Módulo de Leitura (OCR) ATIVADO!"
            cor_ocr = "green"
            msg_extra = f"Lendo de: {pytesseract.pytesseract.tesseract_cmd}"
        else:
            status_ocr = "❌ Módulo de Leitura NÃO encontrado!"
            cor_ocr = "red"
            msg_extra = f"O sistema procurou em: {CAMINHO_TESSERACT_JOSE}"

        tk.Label(self, text=status_ocr, font=("Arial", 10, "bold"), fg=cor_ocr).pack()
        tk.Label(self, text=msg_extra, font=("Arial", 8), fg="gray").pack(pady=5)

        # Botões
        frame = tk.Frame(self)
        frame.pack(pady=20)

        tk.Button(frame, text="1. Pasta RECEITAS", width=25, height=2, command=self.sel_rec).grid(row=0, column=0, padx=5, pady=5)
        self.lbl_rec = tk.Label(frame, text="Nenhuma pasta selecionada")
        self.lbl_rec.grid(row=0, column=1, sticky="w")

        tk.Button(frame, text="2. Pasta DESPESAS", width=25, height=2, command=self.sel_desp).grid(row=1, column=0, padx=5, pady=5)
        self.lbl_desp = tk.Label(frame, text="Nenhuma pasta selecionada")
        self.lbl_desp.grid(row=1, column=1, sticky="w")

        # Botão Principal
        btn_run = tk.Button(self, text="PROCESSAR ARQUIVOS", font=("Arial", 12, "bold"), bg="#008CBA", fg="white", width=30, height=2, command=self.run)
        btn_run.pack(pady=20)

        self.status = tk.Label(self, text="Aguardando...", fg="blue")
        self.status.pack()

    def sel_rec(self):
        p = filedialog.askdirectory()
        if p: 
            self.path_rec = Path(p)
            qtd = len(list(self.path_rec.rglob("*.pdf")))
            self.lbl_rec.config(text=f"{qtd} arquivos encontrados", fg="green")

    def sel_desp(self):
        p = filedialog.askdirectory()
        if p: 
            self.path_desp = Path(p)
            qtd = len(list(self.path_desp.rglob("*.pdf")))
            self.lbl_desp.config(text=f"{qtd} arquivos encontrados", fg="green")

    def run(self):
        if not self.path_rec and not self.path_desp:
            return messagebox.showwarning("Ops", "Selecione as pastas primeiro!")
        
        destino = filedialog.askdirectory(title="Onde salvar o Relatório Final?")
        if not destino: return

        self.status.config(text="Lendo arquivos... (OCR pode demorar um pouco nas Receitas)")
        self.update()

        arquivos_rec = list(self.path_rec.rglob("*.pdf")) if self.path_rec else []
        arquivos_desp = list(self.path_desp.rglob("*.pdf")) if self.path_desp else []

        # Processa tudo
        itens = processar_lista(arquivos_rec, "Receita") + processar_lista(arquivos_desp, "Despesa")

        # Salva
        timestamp = datetime.now().strftime("%H%M%S")
        caminho_excel = Path(destino) / f"Relatorio_Final_{timestamp}.xlsx"
        salvar_excel(caminho_excel, itens)

        # Totais para mensagem
        tot_rec = sum(i.amount for i in itens if i.category == "Receita" and i.status == "OK")
        tot_desp = sum(i.amount for i in itens if i.category == "Despesa" and i.status == "OK")
        pendencias = sum(1 for i in itens if i.status != "OK")

        msg = (f"Concluído!\n\n"
               f"Receitas OK: {format_br(tot_rec)}\n"
               f"Despesas OK: {format_br(tot_desp)}\n"
               f"SALDO: {format_br(tot_rec - tot_desp)}\n\n"
               f"Itens para Revisar: {pendencias}")
        
        messagebox.showinfo("Sucesso", msg)
        self.status.config(text="Processo finalizado.")

if __name__ == "__main__":
    App().mainloop()