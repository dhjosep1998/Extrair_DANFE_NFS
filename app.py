import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os

from extrator import processar_notas


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Extrator de Notas Fiscais")
        self.geometry("520x320")
        self.resizable(False, False)

        ttk.Label(
            self,
            text="Extrator de Notas Fiscais",
            font=("Segoe UI", 16, "bold")
        ).pack(pady=15)

        self.pasta_var = tk.StringVar()

        ttk.Entry(
            self,
            textvariable=self.pasta_var,
            width=60
        ).pack(pady=5)

        ttk.Button(
            self,
            text="Selecionar Pasta",
            command=self.selecionar_pasta
        ).pack(pady=5)

        self.progress = ttk.Progressbar(
            self,
            orient="horizontal",
            length=450,
            mode="determinate"
        )
        self.progress.pack(pady=20)

        self.status = ttk.Label(self, text="Aguardando...")
        self.status.pack()

        self.botao = ttk.Button(
            self,
            text="Processar PDFs",
            command=self.executar
        )
        self.botao.pack(pady=15)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_var.set(pasta)

    def executar(self):
        pasta = self.pasta_var.get()

        if not pasta or not os.path.isdir(pasta):
            messagebox.showerror("Erro", "Selecione uma pasta válida")
            return

        self.botao.config(state="disabled")
        self.progress["value"] = 0
        self.status.config(text="Processando...")

        thread = threading.Thread(
            target=self.processar,
            args=(pasta,),
            daemon=True
        )
        thread.start()

    def processar(self, pasta):
        try:
            ok = processar_notas(
                pasta,
                progresso_callback=self.atualizar_barra
            )
            self.after(0, lambda: self.finalizar(ok))
        except Exception as e:
            self.after(0, lambda: self.erro(e))

    def atualizar_barra(self, valor):
        self.after(0, lambda: self.progress.config(value=valor * 100))

    def finalizar(self, ok):
        self.botao.config(state="normal")
        self.status.config(text="Concluído")

        if ok:
            messagebox.showinfo(
                "Sucesso",
                "Arquivo NOTAS_FISCAIS.xlsx gerado com sucesso!"
            )
        else:
            messagebox.showwarning(
                "Aviso",
                "Nenhum PDF encontrado"
            )

    def erro(self, e):
        self.botao.config(state="normal")
        self.status.config(text="Erro")
        messagebox.showerror("Erro", str(e))


if __name__ == "__main__":
    App().mainloop()
