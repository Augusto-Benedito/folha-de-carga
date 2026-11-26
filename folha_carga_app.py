"""
Folha de Carga - janela para gerar a folha (Produto | Quantidade) a partir do
PDF "Relatório de produtos - Pré-Venda" do sistema.

Uso normal: abrir o programa, escolher o PDF do dia e clicar em "Gerar folha".
Atalho: arrastar o PDF em cima do ícone do programa (.exe) já abre e gera.

Este arquivo precisa estar na mesma pasta do gerar_folha_carga.py
(o criar_exe.bat junta os dois em um único FolhaDeCarga.exe).
"""

import json
import os
import subprocess
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import gerar_folha_carga as core

ORDENS = {
    "Mapa atual (separado por seções)": "mapa",
    "Maior volume primeiro (2L, 1L, 600ML...)": "volume",
    "Ordem do sistema (a do PDF)": "sistema",
    "Alfabética": "alfabetica",
}


# Configuração (último cadastro e ordem escolhida) fica em %APPDATA%\FolhaDeCarga.
def caminho_config():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    pasta = os.path.join(base, "FolhaDeCarga")
    try:
        os.makedirs(pasta, exist_ok=True)
    except OSError:
        pasta = os.path.expanduser("~")
    return os.path.join(pasta, "config.json")


def le_config():
    try:
        with open(caminho_config(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def salva_config(cfg):
    try:
        with open(caminho_config(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def abre_no_sistema(caminho):
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", caminho])
        else:
            subprocess.Popen(["xdg-open", caminho])
    except Exception as e:
        messagebox.showerror("Não consegui abrir", str(e))


def texto_detalhes(r):
    """Texto de apoio mostrado embaixo do resultado."""
    L = []
    L.append(f"Itens lidos do PDF: {r['n_itens_pdf']}")
    L.append(f"Linhas na folha: {r['n_linhas']}  ({r['por_pagina']} por página, frente e verso)")
    L.append(f"Soma das quantidades: {r['soma']:g}")
    if r["total_pdf"] is not None:
        L.append(f"Total do rodapé do sistema: {r['total_pdf']:g}")
    if r["renomeado"]:
        L.append("")
        L.append("O arquivo do dia estava aberto no Excel, então salvei com outro nome.")
    if r["nao_cadastrados"]:
        L.append("")
        L.append("ATENÇÃO - produtos que NÃO estão no cadastro (troque por um cadastro mais novo):")
        L += [f"  - {n}" for n in r["nao_cadastrados"]]
    if r["sem_tipo"]:
        L.append("")
        L.append("ATENÇÃO - embalagem não reconhecida (aparece com '?' na folha):")
        L += [f"  - {n}" for n in r["sem_tipo"]]
    if r.get("sem_mapa"):
        L.append("")
        L.append("ATENÇÃO: não achei o mapa_ordem.csv; separei pelas seções, mas sem a ordem do mapa.")
    if r.get("fora_do_mapa"):
        L.append("")
        L.append("Itens que NÃO estão no mapa atual (ganharam uma linha no fim da seção, em itálico):")
        L += [f"  - {n}" for n in r["fora_do_mapa"]]
    if r["fonte_pequena"]:
        L.append("")
        L.append("ATENÇÃO: muitos itens, a fonte ficou pequena. Considere imprimir em 3 páginas.")
    if r.get("aviso_historico"):
        L.append("")
        L.append("ATENÇÃO: " + r["aviso_historico"])
    L.append("")
    L.append(f"Arquivo: {r['saida']}")
    if r.get("historico"):
        L.append(f"Anotado no histórico: {r['historico']}")
    return "\n".join(L)


# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, pdf_inicial=None):
        super().__init__()
        self.title("Folha de Carga")
        self.minsize(640, 520)
        self.cfg = le_config()
        self.resultado = None

        self.var_pdf = tk.StringVar(value="")
        self.var_cad = tk.StringVar(value=self.cfg.get("cadastro", ""))
        ordem_salva = self.cfg.get("ordem_v2", "mapa")
        nome_ordem = next((k for k, v in ORDENS.items() if v == ordem_salva), list(ORDENS)[0])
        self.var_ordem = tk.StringVar(value=nome_ordem)

        self._monta()
        if pdf_inicial and os.path.exists(pdf_inicial):
            self.var_pdf.set(pdf_inicial)
            self.after(200, self.gerar)

    def _monta(self):
        pad = {"padx": 12, "pady": 4}
        quadro = ttk.Frame(self)
        quadro.pack(fill="both", expand=True, padx=8, pady=8)
        quadro.columnconfigure(1, weight=1)

        ttk.Label(quadro, text="1. PDF do dia (do sistema)", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", **pad
        )
        ttk.Entry(quadro, textvariable=self.var_pdf, state="readonly").grid(
            row=1, column=0, columnspan=2, sticky="ew", **pad
        )
        ttk.Button(quadro, text="Escolher PDF...", command=self.escolhe_pdf).grid(
            row=1, column=2, sticky="e", **pad
        )

        ttk.Label(quadro, text="2. Cadastro de produtos (.xlsx)", font=("Segoe UI", 10, "bold")).grid(
            row=2, column=0, columnspan=3, sticky="w", **pad
        )
        ttk.Entry(quadro, textvariable=self.var_cad, state="readonly").grid(
            row=3, column=0, columnspan=2, sticky="ew", **pad
        )
        ttk.Button(quadro, text="Trocar cadastro...", command=self.escolhe_cadastro).grid(
            row=3, column=2, sticky="e", **pad
        )

        ttk.Label(quadro, text="3. Ordem das linhas", font=("Segoe UI", 10, "bold")).grid(
            row=4, column=0, columnspan=3, sticky="w", **pad
        )
        ttk.Combobox(
            quadro, textvariable=self.var_ordem, values=list(ORDENS), state="readonly"
        ).grid(row=5, column=0, columnspan=3, sticky="ew", **pad)

        self.btn_gerar = ttk.Button(quadro, text="GERAR FOLHA", command=self.gerar)
        self.btn_gerar.grid(row=6, column=0, columnspan=3, sticky="ew", padx=12, pady=(14, 6), ipady=8)

        self.lbl_status = tk.Label(quadro, text="", font=("Segoe UI", 13, "bold"), anchor="w")
        self.lbl_status.grid(row=7, column=0, columnspan=3, sticky="ew", padx=12, pady=(6, 2))

        self.txt = tk.Text(quadro, height=12, wrap="word", font=("Consolas", 9), state="disabled")
        self.txt.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=12, pady=4)
        quadro.rowconfigure(8, weight=1)

        barra = ttk.Frame(quadro)
        barra.grid(row=9, column=0, columnspan=3, sticky="ew", padx=12, pady=(4, 8))
        self.btn_abrir = ttk.Button(barra, text="Abrir planilha", command=self.abrir_planilha, state="disabled")
        self.btn_abrir.pack(side="left")
        self.btn_pasta = ttk.Button(barra, text="Abrir pasta", command=self.abrir_pasta, state="disabled")
        self.btn_pasta.pack(side="left", padx=8)

    # ----------------------------------------------------------------------
    def escolhe_pdf(self):
        p = filedialog.askopenfilename(
            title="Escolha o PDF do sistema", filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
        )
        if p:
            self.var_pdf.set(p)

    def escolhe_cadastro(self):
        p = filedialog.askopenfilename(
            title="Escolha o cadastro de produtos (.xlsx)",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if p:
            self.var_cad.set(p)
            self.cfg["cadastro"] = p
            salva_config(self.cfg)

    def mostra(self, texto, cor, detalhes=""):
        self.lbl_status.config(text=texto, fg=cor)
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", detalhes)
        self.txt.config(state="disabled")

    def gerar(self):
        pdf = self.var_pdf.get().strip()
        if not pdf:
            messagebox.showinfo("Falta o PDF", "Clique em 'Escolher PDF...' e selecione o PDF do dia.")
            return
        cad = self.var_cad.get().strip() or None
        ordem = ORDENS[self.var_ordem.get()]
        self.cfg["ordem_v2"] = ordem
        salva_config(self.cfg)

        self.config(cursor="watch")
        self.btn_gerar.config(state="disabled")
        self.mostra("Gerando...", "#444444")
        self.update_idletasks()
        try:
            r = core.processar(pdf, cadastro=cad, ordem=ordem)
        except RuntimeError as e:
            self.resultado = None
            self.mostra("Não foi possível gerar", "#b00020", str(e))
            self.btn_abrir.config(state="disabled")
            self.btn_pasta.config(state="disabled")
            if "cadastro" in str(e).lower():
                messagebox.showwarning("Cadastro", str(e))
        except Exception:
            self.resultado = None
            self.mostra(
                "Erro inesperado",
                "#b00020",
                "Copie este texto e envie para quem faz manutenção:\n\n" + traceback.format_exc(),
            )
            self.btn_abrir.config(state="disabled")
            self.btn_pasta.config(state="disabled")
        else:
            self.resultado = r
            if r["cadastro"] and r["cadastro"] != self.var_cad.get():
                self.var_cad.set(r["cadastro"])
                self.cfg["cadastro"] = r["cadastro"]
                salva_config(self.cfg)
            if r["ok"]:
                cor = "#1b7a2b"
                msg = f"CONFERÊNCIA OK  -  total {r['soma']:g}"
            else:
                cor = "#b00020"
                msg = f"NÃO BATEU: folha {r['soma']:g} x sistema {r['total_pdf']}  -  não imprima, confira"
            if r["nao_cadastrados"] or r["sem_tipo"]:
                cor = "#b00020" if not r["ok"] else "#b36b00"
                msg += "  (veja os avisos)"
            self.mostra(msg, cor, texto_detalhes(r))
            self.btn_abrir.config(state="normal")
            self.btn_pasta.config(state="normal")
        finally:
            self.btn_gerar.config(state="normal")
            self.config(cursor="")

    def abrir_planilha(self):
        if self.resultado:
            abre_no_sistema(self.resultado["saida"])

    def abrir_pasta(self):
        if self.resultado:
            abre_no_sistema(os.path.dirname(self.resultado["saida"]))


# --------------------------------------------------------------------------
def main():
    pdf = None
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".pdf"):
        pdf = sys.argv[1]
    App(pdf).mainloop()


if __name__ == "__main__":
    main()
