import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import sys
import io

from processing import processar_planilha

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("HokiSeg - Atualizador de Planilha")
        self.root.geometry("700x550")

        self.pdf_path = ""
        self.excel_path = ""

        # Frame para os botões e labels
        frame_botoes = tk.Frame(root, pady=10)
        frame_botoes.pack(fill='x')

        # Botão PDF
        self.btn_pdf = tk.Button(frame_botoes, text="1. Selecionar PDF", command=self.selecionar_pdf, width=20, height=2)
        self.btn_pdf.grid(row=0, column=0, padx=10, pady=5)
        self.lbl_pdf = tk.Label(frame_botoes, text="Nenhum PDF selecionado", fg="red")
        self.lbl_pdf.grid(row=0, column=1, padx=5, sticky="w")

        # Botão Planilha
        self.btn_excel = tk.Button(frame_botoes, text="2. Selecionar Planilha", command=self.selecionar_excel, width=20, height=2)
        self.btn_excel.grid(row=1, column=0, padx=10, pady=5)
        self.lbl_excel = tk.Label(frame_botoes, text="Nenhuma planilha selecionada", fg="red")
        self.lbl_excel.grid(row=1, column=1, padx=5, sticky="w")

        # Botão Processar
        self.btn_processar = tk.Button(root, text="3. Processar e Atualizar Planilha", command=self.iniciar_processamento, height=3, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.btn_processar.pack(pady=10, fill='x', padx=10)

        # Área de Log (para mostrar os resultados)
        self.log_area = scrolledtext.ScrolledText(root, height=20, width=80, wrap=tk.WORD)
        self.log_area.pack(pady=10, padx=10, fill="both", expand=True)
        self.log_area.insert(tk.END, "Bem-vindo! Por favor, siga os passos:\n\n1. Selecione o arquivo PDF.\n2. Selecione a planilha Excel para atualizar.\n3. Clique em 'Processar'.\n\n")
        self.log_area.config(state="disabled")

    def selecionar_pdf(self):
        path = filedialog.askopenfilename(
            title="Selecione o PDF de extrato",
            filetypes=[("PDF files", "*.pdf")]
        )
        if path:
            self.pdf_path = path
            self.lbl_pdf.config(text=os.path.basename(path), fg="green")
            self.log_message(f"Arquivo PDF selecionado: {path}\n")

    def selecionar_excel(self):
        path = filedialog.askopenfilename(
            title="Selecione a Planilha Financeira",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.excel_path = path
            self.lbl_excel.config(text=os.path.basename(path), fg="green")
            self.log_message(f"Planilha Excel selecionada: {path}\n")

    def log_message(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, message)
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.root.update_idletasks()

    def iniciar_processamento(self):
        if not self.pdf_path or not self.excel_path:
            messagebox.showerror("Arquivos Faltando", "Por favor, selecione o arquivo PDF E a planilha Excel antes de processar.")
            return

        self.log_message("\n========================================\n")
        self.log_message("Iniciando processamento...\n")
        self.btn_processar.config(text="Processando...", state="disabled")

        log_stream = io.StringIO()
        sys.stdout = log_stream
        sys.stderr = log_stream

        try:
            processar_planilha(self.pdf_path, self.excel_path)

            log_output = log_stream.getvalue()
            self.log_message(log_output)
            
            if "✅" in log_output:
                messagebox.showinfo("Sucesso!", "Planilha atualizada com sucesso. Verifique o log para detalhes.")
            elif "⚠️" in log_output or "❌" in log_output:
                 messagebox.showwarning("Aviso", "Processamento concluído com avisos ou erros. Verifique o log.")
            else:
                 messagebox.showinfo("Concluído", "Processamento finalizado sem dados novos.")

        except Exception as e:
            log_output = log_stream.getvalue()
            self.log_message(log_output)
            self.log_message(f"\n❌ ERRO CRÍTICO: {e}\n")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{e}")
        finally:
            # Restaura o stdout e reativa o botão
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
            self.btn_processar.config(text="3. Processar e Atualizar Planilha", state="normal")
            self.log_message("========================================\n")
