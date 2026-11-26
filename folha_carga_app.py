"""
Folha de Carga - janela para gerar a folha (Produto | Quantidade) a partir do
PDF "Relatório de produtos - Pré-Venda" do sistema.

Uso normal: abrir o programa, escolher o PDF do dia e clicar em "Gerar folha".
A aba "Histórico" mostra as folhas já geradas (do "Histórico de cargas.xlsx").
Atalho: arrastar o PDF em cima do ícone do programa (.exe) já abre e gera.

Este arquivo precisa estar na mesma pasta do gerar_folha_carga.py
(o criar_exe.bat junta os dois em um único FolhaDeCarga.exe).
"""

import json
import os
import subprocess
import sys
import traceback
import datetime
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
    if r.get("vasilhames"):
        L.append("")
    for v in r.get("vasilhames") or []:
        produto, qtd = core.texto_vasilhame(v)
        L.append(f"{produto}: {qtd}")
        if v["sem_tamanho"]:
            L.append(f"ATENÇÃO - {v['nome']} sem tamanho de caixa no cadastro (não entrou nas garrafas):")
            L += [f"  - {n}" for n in v["sem_tamanho"]]
    if r.get("aviso_historico"):
        L.append("")
        L.append("ATENÇÃO: " + r["aviso_historico"])
    L.append("")
    L.append(f"Arquivo: {r['saida']}")
    if r.get("historico"):
        L.append(f"Anotado no histórico: {r['historico']}")
    return "\n".join(L)


def texto_hist(v):
    """Valor de uma célula do histórico como texto para a tabela da aba Histórico."""
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%d/%m/%Y %H:%M")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, pdf_inicial=None):
        super().__init__()
        self.title("Folha de Carga")
        self.minsize(820, 520)
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
        self.abas = ttk.Notebook(self)
        self.abas.pack(fill="both", expand=True, padx=8, pady=8)
        quadro = ttk.Frame(self.abas)
        self.abas.add(quadro, text="  Gerar folha  ")
        self.aba_hist = ttk.Frame(self.abas)
        self.abas.add(self.aba_hist, text="  Histórico  ")
        self._monta_gerar(quadro)
        self._monta_historico(self.aba_hist)
        # Ao abrir a aba Histórico, a lista é lida de novo do arquivo.
        self.abas.bind("<<NotebookTabChanged>>", self._mudou_aba)

    def _monta_gerar(self, quadro):
        pad = {"padx": 12, "pady": 4}
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
    # Aba Histórico: lista das folhas geradas, lida do "Histórico de cargas.xlsx"
    # da pasta das folhas (a pasta da última folha gerada, ou a escolhida aqui).
    # (coluna do histórico, título na tela, largura, alinhamento)
    COLUNAS_HIST = [
        ("Gerado em", "Gerado em", 140, "center"),
        ("Data do mapa", "Data do mapa", 110, "center"),
        ("Mapa", "Mapa", 180, "w"),
        ("Linhas na folha", "Linhas", 65, "center"),
        ("Volumes", "Volumes", 80, "center"),
        ("Total do sistema", "Sistema", 80, "center"),
        ("Conferência", "Conferência", 110, "center"),
    ]

    def _monta_historico(self, aba):
        topo = ttk.Frame(aba)
        topo.pack(fill="x", padx=12, pady=(10, 4))
        self.var_pasta_hist = tk.StringVar(value=self.cfg.get("pasta_historico", ""))
        ttk.Label(topo, text="Pasta das folhas:", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Entry(topo, textvariable=self.var_pasta_hist, state="readonly").pack(
            side="left", fill="x", expand=True, padx=8
        )
        ttk.Button(topo, text="Trocar pasta...", command=self.escolhe_pasta_hist).pack(side="left")

        self.lbl_hist = tk.Label(aba, text="", font=("Segoe UI", 10), anchor="w")
        self.lbl_hist.pack(fill="x", padx=12, pady=(4, 2))

        meio = ttk.Frame(aba)
        meio.pack(fill="both", expand=True, padx=12, pady=4)
        nomes = [c[0] for c in self.COLUNAS_HIST]
        self.tabela = ttk.Treeview(meio, columns=nomes, show="headings", selectmode="browse")
        for nome, titulo, largura, alinha in self.COLUNAS_HIST:
            self.tabela.heading(nome, text=titulo)
            self.tabela.column(nome, width=largura, anchor=alinha, stretch=(nome == "Mapa"))
        self.tabela.tag_configure("nao_bateu", foreground="#b00020")
        self.tabela.tag_configure("sem_arquivo", foreground="#888888")
        rolagem = ttk.Scrollbar(meio, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=rolagem.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")
        self.tabela.bind("<Double-1>", lambda e: self.abrir_folha_hist())

        barra = ttk.Frame(aba)
        barra.pack(fill="x", padx=12, pady=(4, 8))
        ttk.Button(barra, text="Abrir folha selecionada", command=self.abrir_folha_hist).pack(side="left")
        ttk.Button(barra, text="Abrir histórico no Excel", command=self.abrir_planilha_hist).pack(
            side="left", padx=8
        )
        ttk.Button(barra, text="Atualizar", command=self.carrega_historico).pack(side="right")
        self.folhas_hist = {}

    def _mudou_aba(self, _evento=None):
        if self.abas.select() == str(self.aba_hist):
            self.carrega_historico()

    def escolhe_pasta_hist(self):
        p = filedialog.askdirectory(title="Escolha a pasta onde ficam as folhas de carga")
        if p:
            self._usa_pasta_hist(p)
            self.carrega_historico()

    def _usa_pasta_hist(self, pasta):
        pasta = os.path.normpath(pasta)
        self.var_pasta_hist.set(pasta)
        if self.cfg.get("pasta_historico") != pasta:
            self.cfg["pasta_historico"] = pasta
            salva_config(self.cfg)

    def carrega_historico(self):
        self.tabela.delete(*self.tabela.get_children())
        self.folhas_hist = {}
        pasta = self.var_pasta_hist.get()
        if not pasta:
            self.lbl_hist.config(
                text="Gere uma folha ou clique em 'Trocar pasta...' para ver o histórico.", fg="#444444"
            )
            return
        try:
            folhas = core.le_historico(pasta)
        except Exception as e:
            self.lbl_hist.config(text=f"Não consegui ler o histórico: {e}", fg="#b00020")
            return
        if not folhas:
            self.lbl_hist.config(text="Ainda não há folhas no histórico desta pasta.", fg="#444444")
            return
        for f in folhas:
            valores = [texto_hist(f.get(c[0])) for c in self.COLUNAS_HIST]
            tags = []
            if f.get("Conferência") == "NÃO BATEU":
                tags.append("nao_bateu")
            if not (f["caminho"] and os.path.exists(f["caminho"])):
                tags.append("sem_arquivo")
            item = self.tabela.insert("", "end", values=valores, tags=tags)
            self.folhas_hist[item] = f
        # Ordem cronológica: a folha mais nova fica no fim; já deixa ela à vista e selecionada.
        ultimo = self.tabela.get_children()[-1]
        self.tabela.see(ultimo)
        self.tabela.selection_set(ultimo)
        nao_bateu = sum(1 for f in folhas if f.get("Conferência") == "NÃO BATEU")
        texto = f"{len(folhas)} folha(s) no histórico. Dois cliques numa linha abrem a folha."
        if nao_bateu:
            texto += f"  {nao_bateu} não bateu com o sistema (em vermelho)."
        self.lbl_hist.config(text=texto, fg="#b00020" if nao_bateu else "#1b7a2b")

    def abrir_folha_hist(self):
        sel = self.tabela.selection()
        if not sel:
            messagebox.showinfo("Histórico", "Selecione uma folha na lista.")
            return
        caminho = self.folhas_hist[sel[0]]["caminho"]
        if not caminho or not os.path.exists(caminho):
            messagebox.showwarning(
                "Histórico", f"Não achei a folha:\n{caminho}\n\nEla pode ter sido apagada ou movida."
            )
            return
        abre_no_sistema(caminho)

    def abrir_planilha_hist(self):
        pasta = self.var_pasta_hist.get()
        caminho = os.path.join(pasta, core.NOME_HISTORICO) if pasta else ""
        if not caminho or not os.path.exists(caminho):
            messagebox.showinfo("Histórico", "Ainda não há histórico nesta pasta.")
            return
        abre_no_sistema(caminho)

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
            if r.get("historico"):
                self._usa_pasta_hist(os.path.dirname(r["historico"]))
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
