"""
Gera a folha de carga (Produto | Quantidade) a partir do PDF
"Relatório de produtos - Pré-Venda" do sistema.

Uso:
    python gerar_folha_carga.py MAPA_A_2026_11_10.pdf
    python gerar_folha_carga.py MAPA_A_2026_11_10.pdf --ordem sistema      (ordem do PDF)
    python gerar_folha_carga.py MAPA_A_2026_11_10.pdf --ordem alfabetica
    python gerar_folha_carga.py MAPA_A_2026_11_10.pdf --ordem volume       (maior volume primeiro)

Por padrão a folha sai separada e ordenada como o MAPA_ATUAL (usa o mapa_ordem.csv).
Itens que não estão no mapa ou no cadastro ganham uma linha no fim da seção certa.
    python gerar_folha_carga.py MAPA_A_2026_11_10.pdf --cadastro relatorio_Lista_de_produtos_2026-11-10.xlsx

Precisa de: pip install pdfplumber openpyxl

Apelidos (opcional): crie um arquivo apelidos.csv ao lado do script, com duas
colunas (sem acento no cabeçalho):
    base,nome_curto
    CERV ANTARTICA PILSEN 350ML,Antarctica Pilsen 350ML
A coluna "base" é o nome do cadastro SEM a embalagem (FD 12X1, CX 24X1, UN...).
"""

import argparse
import csv
import datetime
import glob
import math
import os
import re
import sys
import unicodedata

import pdfplumber
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

RANK_TIPO = {"FD": 0, "CX": 1, "UN": 2}
SIGLAS_MAIUSCULAS = {"RET", "NS", "PET", "KS", "LT", "ICE", "H2O"}

RE_EMBALAGEM = re.compile(
    r"\s+(FD|CX|UN)\.?(?:\s*(\d+)\s*[Xx]\s*(\d+))?\s*$", re.IGNORECASE
)


# --------------------------------------------------------------------------
def normaliza(s):
    return re.sub(r"\s+", " ", str(s).strip())


def separa_nome(nome):
    """'SPRITE 350ML FD 12X1' -> ('SPRITE 350ML', 'FD', 12)"""
    nome = normaliza(nome)
    m = RE_EMBALAGEM.search(nome)
    if not m:
        return nome, None, None
    arranjo = int(m.group(2)) if m.group(2) else None
    return nome[: m.start()].strip(), m.group(1).upper(), arranjo


def chave(base):
    """Chave de agrupamento: maiúsculas, sem pontos, espaços únicos."""
    return normaliza(base.upper().replace(".", " "))


def nome_bonito(base):
    palavras = []
    for w in base.split():
        if any(c.isdigit() for c in w) or w.upper().strip(".") in SIGLAS_MAIUSCULAS:
            palavras.append(w.upper())
        else:
            palavras.append(w.capitalize())
    return " ".join(palavras)


# --------------------------------------------------------------------------
def _como_texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def le_cadastro(caminho):
    wb = load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    linhas = ws.iter_rows(values_only=True)
    try:
        cab = [_como_texto(c).lower() for c in next(linhas)]
    except StopIteration:
        raise RuntimeError("O cadastro está vazio.")
    if "referencia" not in cab or "nome" not in cab:
        raise RuntimeError("O cadastro não tem as colunas 'referencia' e 'nome'.")
    i_ref, i_nome = cab.index("referencia"), cab.index("nome")
    por_ref = {}
    variantes = {}
    for lin in linhas:
        if lin is None or len(lin) <= max(i_ref, i_nome):
            continue
        ref, nome = _como_texto(lin[i_ref]), _como_texto(lin[i_nome])
        if not ref or not nome:
            continue
        base, tipo, arr = separa_nome(nome)
        por_ref[ref] = (base, tipo, arr)
        if tipo and arr:
            variantes.setdefault((chave(base), tipo), set()).add(arr)
    wb.close()
    return por_ref, variantes


# --------------------------------------------------------------------------
def le_apelidos(caminho):
    apelidos = {}
    if caminho and os.path.exists(caminho):
        with open(caminho, encoding="utf-8-sig", newline="") as f:
            for linha in csv.DictReader(f):
                if linha.get("base") and linha.get("nome_curto"):
                    apelidos[chave(linha["base"])] = linha["nome_curto"].strip()
    return apelidos


# --------------------------------------------------------------------------
def num(txt):
    """Número no formato brasileiro: 1.234,50 -> 1234.5 (usado no total do rodapé)."""
    return float(txt.replace(".", "").replace(",", "."))


def le_pdf(caminho):
    linhas = []
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            linhas += (pagina.extract_text() or "").splitlines()

    padrao_ref = re.compile(r"^\d\w*$")
    padrao_linha = re.compile(
        r"^(?P<nome>.+?)\s+(?P<qtd>\d[\d.]*)\s+(?P<vm>[\d.]*,\d\d)\s+(?P<vt>[\d.]*,\d\d)$"
    )

    itens, ref, total, data, grupo = [], None, None, None, ""
    espera_total = False
    for l in linhas:
        l = l.strip()
        if espera_total:
            m = re.match(r"([\d.]*,\d+)", l)
            if m:
                total = num(m.group(1))
                espera_total = False
                continue
        m = re.match(r"Quantidade total:\s*([\d.]*,\d+)?", l)
        if m:
            if m.group(1):
                total = num(m.group(1))
            else:
                espera_total = True
            continue
        m = re.match(r"Intervalo:\s*(\d{2})/(\d{2})/(\d{4})", l)
        if m:
            data = f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        m = re.match(r"Grupo do Pedido:\s*(.+)$", l)
        if m:
            grupo = m.group(1).strip()
        if l.startswith(("Ano/m", "Produto ", "Gerado")):
            continue
        if padrao_ref.match(l) and len(l) >= 6:
            ref = l
            continue
        m = padrao_linha.match(l)
        if m and ref:
            qtd = float(m.group("qtd"))
            itens.append(
                {"ref": ref, "nome": normaliza(m.group("nome")), "qtd": qtd}
            )
            ref = None
    return itens, total, data, grupo


# --------------------------------------------------------------------------
RE_VOLUME = re.compile(r"(?<![\w.,])(\d+(?:[.,]\d+)?)\s*(ML|L)\b", re.IGNORECASE)


def volume_ml(base):
    """Volume do produto em ml, lido do nome ('2L' -> 2000, '1,5L' -> 1500, '350ML' -> 350)."""
    m = RE_VOLUME.search(base)
    if not m:
        return None
    v = float(m.group(1).replace(",", "."))
    return v * 1000 if m.group(2).upper() == "L" else v


# Seções do MAPA_ATUAL, na ordem em que aparecem na folha (código, título impresso).
SECOES = [
    ("600ML", "600ML"),
    ("LATAS", "LATAS"),
    ("200ML", "200ML"),
    ("AGUAS", "ÁGUAS"),
    ("ENERGETICOS", "ENERGÉTICOS"),
    ("SUCOS", "SUCOS"),
    ("2L", "2L / 2,25L"),
    ("RETORNAVEIS", "RETORNÁVEIS"),
    ("CERVEJAS", "CERVEJAS"),
    ("CERVEJAS RET", "CERVEJAS RETORNÁVEIS"),
    ("DESTILADOS", "DESTILADOS"),
    ("ICES", "ICES"),
    ("LIMPEZA", "LIMPEZA"),
    ("OUTROS", "OUTROS"),
]
ORDEM_SECOES = [c for c, _ in SECOES]
TITULO_SECAO = dict(SECOES)


def secao_da_base(nome, vol=None):
    """Em qual seção do mapa um produto entra, pelo nome (usado para produtos
    que não estão no mapa ou nem no cadastro)."""
    t = _sem_acento(nome).upper()
    if vol is None:
        vol = volume_ml(nome)
    ret = bool(re.search(r"\bRET\b|\bKS\b|\bNS\b", t.replace(".", " ")))
    if "SANITARIA" in t:
        return "LIMPEZA"
    if re.search(r"AGUA DE COCO|KERO ?COCO", t):
        return "SUCOS"
    if t.startswith("AGUA") or t.startswith("H2O"):
        return "AGUAS"
    if re.search(r"MONSTER|ENERGETICO|RED BULL", t):
        return "ENERGETICOS"
    if re.search(r"GATORADE|POWERADE|DEL VALLE|KAPO|TODDYNHO|TAMPICO|NATIVO|GUARAVITON|GROSELHA|ICE TEA|SUCO", t):
        return "SUCOS"
    if re.search(r"^(51 )?ICE\b", t):
        return "ICES"
    if re.search(r"^CERV|SPATEN|ESTRELLA|THEREZOPOLIS|HEINEKEN|BUDWEISER|STELLA|CORONA|AMSTEL|SKOL|BRAHMA|CHOPP", t):
        return "CERVEJAS RET" if ret or (vol is not None and vol in (300, 600, 1000)) else "CERVEJAS"
    if re.search(r"CACHACA|VODKA|CONHAQUE|COQUETEL|COROTE|RECORD|VINHO|CAPANGA|CYNAR|CAMPARI|JURUBEBA|CATUABA|AGUARDENTE|APERITIVO|WHISKY|GIN\b|PARATUDO|YPIOCA|PAIZANO",
                 t):
        return "DESTILADOS"
    if re.search(r"CAIXA PLASTICA|GARRAFA DE VIDRO", t):
        return "OUTROS"
    if ret:
        return "RETORNAVEIS"
    if vol is None:
        return "OUTROS"
    if vol >= 1500:
        return "2L"
    if vol >= 500:
        return "600ML"
    if vol >= 250:
        return "LATAS"
    return "200ML"


# --------------------------------------------------------------------------
def monta_linhas(itens, por_ref, variantes, apelidos, ordem, ordem_mapa=None):
    """Devolve (linhas, nao_cadastrados, sem_tipo, fora_do_mapa).
    linhas = [(nome, texto, [quantidades numéricas], volume_ml, chave_do_produto)]."""
    grupos = {}
    avisos_nao_cadastrados = []
    for it in itens:
        if it["ref"] in por_ref:
            base, tipo, arr = por_ref[it["ref"]]
        else:
            base, tipo, arr = separa_nome(it["nome"])
            avisos_nao_cadastrados.append(it["nome"])
        k = chave(base)
        g = grupos.setdefault(k, {"base": base, "itens": []})
        g["itens"].append({"tipo": tipo, "arr": arr, "qtd": it["qtd"]})

    linhas = []
    sem_tipo = []
    for k, g in grupos.items():
        partes = []
        for i in g["itens"]:
            tipo, arr = i["tipo"], i["arr"]
            if tipo is None:
                rotulo = "?"
                sem_tipo.append(g["base"])
                plain = True
            elif tipo == "UN":
                rotulo, plain = "UN", True
            elif tipo == "FD":
                # FD na folha = fardo de 12 (nos produtos de 1L ou mais, fardo de 6).
                # Fardo com outra quantidade sai como (C/6), (C/18)...
                # Ex.: Sprite Zero 350ML só tem FD 6X1 -> sai (C/6), nunca (FD).
                vol = volume_ml(g["base"])
                padrao = 6 if vol is not None and vol >= 1000 else 12
                if arr is None:
                    rotulo, plain = "FD", True
                else:
                    plain = arr == padrao
                    rotulo = "FD" if plain else f"C/{arr}"
            else:
                vs = variantes.get((k, tipo), set())
                if len(vs) <= 1 or arr is None:
                    rotulo, plain = tipo, True
                else:
                    padrao = 12 if 12 in vs else max(vs)
                    plain = arr == padrao
                    rotulo = tipo if plain else f"C/{arr}"
            q = int(i["qtd"]) if float(i["qtd"]).is_integer() else i["qtd"]
            partes.append(
                (RANK_TIPO.get(tipo, 9), 0 if plain else 1, arr or 0, q, f"{q}({rotulo})")
            )
        partes.sort()
        nome = apelidos.get(k) or nome_bonito(g["base"])
        linhas.append(
            (
                nome,
                "; ".join(p[4] for p in partes),
                [p[3] for p in partes],
                volume_ml(g["base"]),
                k,
            )
        )

    # "sistema": mantém a ordem em que os produtos aparecem no PDF.
    if ordem == "volume":
        # Maior volume primeiro; produtos sem volume no nome vão para o fim.
        # Empate: ordem alfabética.
        linhas.sort(key=lambda x: (x[3] is None, -(x[3] or 0), x[0].upper()))
    elif ordem == "alfabetica":
        linhas.sort(key=lambda x: x[0].upper())
    fora_do_mapa = []
    if ordem == "mapa":
        om = ordem_mapa or {}
        # Itens fora do mapa entram no fim da seção deles (posição .5 depois do último).
        fim_secao = {}
        for pos, sec in om.values():
            fim_secao[sec] = max(fim_secao.get(sec, pos), pos)
        maior = max([p for p, _ in om.values()] or [0])
        novas = []
        for x in linhas:
            if x[4] in om:
                pos, sec = om[x[4]]
                novo = False
            else:
                sec = secao_da_base(grupos[x[4]]["base"], x[3])
                if sec in fim_secao:
                    pos = fim_secao[sec] + 0.5
                else:
                    pos = maior + 1 + ORDEM_SECOES.index(sec if sec in ORDEM_SECOES else "OUTROS")
                novo = True
                fora_do_mapa.append(x[0])
            novas.append((x + (sec, novo), pos))
        novas.sort(key=lambda t: (t[1], -(t[0][3] or 0), t[0][0].upper()))
        linhas = [t[0] for t in novas]
    return linhas, avisos_nao_cadastrados, sem_tipo, fora_do_mapa


# --------------------------------------------------------------------------
# Vasilhames: a garrafa retornável que vai cheia no caminhão volta vazia.
# A folha mostra, no fim da seção dos produtos (RETORNÁVEIS, CERVEJAS RETORNÁVEIS, DESTILADOS),
# quanto de cada vasilhame deve voltar.
# (nome na folha, regra pelo nome do produto, garrafas por casco).
# Garrafas por casco = None: cada CX/FD vendida é um casco (o tamanho do cadastro).
# Com um número: soma as garrafas e divide pelo casco (o que sobra sai como UN).
# A ordem aqui é a ordem na folha (a do mapa).
TIPOS_VASILHAME = [
    # Todo refrigerante 2L retornável (Coca, Coca Zero, Fanta...). O casco leva 12 garrafas,
    # mesmo vendendo em FD 6X1.
    ("Refri 2L", re.compile(r"^(?!CERV\b)(?=.*\b2L\b)(?=.*\bRET\b)"), 12),
    # Tudo que é KS: Coca, Coca Zero, Fanta, Sprite... KS 290ML.
    ("KS", re.compile(r"\bKS\b"), None),
    # Guaraná 300ML retornável (não confundir com a cerveja litrinho 300ML).
    ("Guaraná 300ML", re.compile(r"^GUARANA\b.*\b300ML\b.*\bRET\b"), None),
    # Cerveja 600ML retornável (Skol, Brahma, Antarctica, Original, Amstel, Spaten RET...).
    # O RET pode vir antes ou depois do 600ML (CERV SPATEN RET. 600ML).
    # A Amstel 600ML volta junto com as da Ambev; a Heineken tem casco próprio (abaixo).
    ("Cerveja 600ML", re.compile(r"^CERV\b(?!.*\bHEINEKEN\b)(?=.*\b600ML\b)(?=.*\bRET\b)"), None),
    # Heineken 600ML: casco próprio. No mapa ela não tem RET no nome (CERV HEINEKEN 600ML),
    # então conta toda Heineken 600ML, menos a descartável (DESC).
    ("Heineken 600ML", re.compile(r"^CERV\b(?=.*\bHEINEKEN\b)(?=.*\b600ML\b)(?!.*\bDESC\b)"), None),
    # Cerveja litrão 1L retornável. O Amstel Litrão tem casco próprio (abaixo).
    ("Litrão 1L", re.compile(r"^CERV\b(?!.*\bAMSTEL\b)(?=.*\b(1L|1000ML)\b)(?=.*\bRET\b)"), None),
    ("Amstel Litrão 1L", re.compile(r"^CERV\b(?=.*\bAMSTEL\b)(?=.*\b(1L|1000ML)\b)(?=.*\bRET\b)"), None),
    # Destilados retornáveis (garrafa de vidro), todos numa linha só, casco de 12:
    # Cachaça 51 965ML (normal, Ouro, Seleção), Velho Barreiro 910ML, Villa Velha 965ML
    # e os Record de 900/960ML (o Corote Record e a Vodka Record 500ML não contam).
    (
        "Destilados",
        re.compile(
            r"^(CACHACA|AGUARDENTE)\b.*\b51\b.*\b965ML\b"
            r"|\bVELHO BARREIRO\b.*\b910ML\b"
            r"|\bVILL?A VELHA\b.*\b965ML\b"
            r"|^RECORD\b.*\b(900|960)ML\b"
        ),
        12,  # o casco leva 12 garrafas
    ),
]


def tipo_vasilhame(base):
    """Nome do vasilhame do produto ('KS', 'Refri 2L'...), ou None se não é um deles."""
    k = _sem_acento(chave(base))
    for nome, regra, _ in TIPOS_VASILHAME:
        if regra.search(k):
            return nome
    return None


def conta_vasilhames(itens, por_ref):
    """Soma os vasilhames do PDF. Devolve uma lista (na ordem de TIPOS_VASILHAME,
    só os que aparecem) de dicts:
    nome = 'KS', 'Refri 2L'...; por_tipo = {'CX': engradados, 'FD': ..., 'UN': garrafas soltas};
    garrafas = total de garrafas (pelo tamanho da caixa no cadastro);
    chaves = produtos daquele vasilhame (para achar o lugar na folha);
    sem_tamanho = produtos cuja embalagem não diz quantas garrafas tem."""
    contas = {}
    for it in itens:
        if it["ref"] in por_ref:
            base, tipo, arr = por_ref[it["ref"]]
        else:
            base, tipo, arr = separa_nome(it["nome"])
        nome = tipo_vasilhame(base)
        if not nome:
            continue
        v = contas.setdefault(
            nome, {"nome": nome, "por_tipo": {}, "garrafas": 0, "chaves": set(), "sem_tamanho": []}
        )
        v["chaves"].add(chave(base))
        q = it["qtd"]
        if tipo == "UN":
            v["por_tipo"]["UN"] = v["por_tipo"].get("UN", 0) + q
            v["garrafas"] += q
        elif tipo in ("CX", "FD") and arr:
            v["por_tipo"][tipo] = v["por_tipo"].get(tipo, 0) + q
            v["garrafas"] += q * arr
        else:
            v["sem_tamanho"].append(nome_bonito(base))
    inteiro = lambda x: int(x) if float(x).is_integer() else x
    saida = []
    for nome, _, por_casco in TIPOS_VASILHAME:
        if nome in contas:
            v = contas[nome]
            v["garrafas"] = inteiro(v["garrafas"])
            if por_casco:
                # Ex.: Refri 2L: 7 FD 6X1 = 42 garrafas = 3 cascos de 12 + 6 garrafas soltas.
                cascos, soltas = divmod(v["garrafas"], por_casco)
                v["por_tipo"] = {t: inteiro(q) for t, q in (("CX", cascos), ("UN", soltas)) if q}
            else:
                v["por_tipo"] = {t: inteiro(q) for t, q in v["por_tipo"].items()}
            v["sem_tamanho"] = sorted(set(v["sem_tamanho"]))
            saida.append(v)
    return saida


def texto_vasilhame(v):
    """(coluna Produto, coluna Quantidade) da linha de um vasilhame."""
    partes = [f"{v['por_tipo'][t]} {t}" for t in ("CX", "FD", "UN") if v["por_tipo"].get(t)]
    return f"Vasilhame {v['nome']} que volta ({v['garrafas']} garrafas)", " + ".join(partes) or "0"


# --------------------------------------------------------------------------
# Quadradinho de cada linha, para quem carrega ir marcando o que já subiu no caminhão.
CAIXINHA = "☐"
LARGURA_CAIXINHA = 5
# Soma máxima das larguras A+B+C que cabe no A4 em pé com margens de 0,5".
LARGURA_MAX_FOLHA = 94


def escreve_xlsx(linhas, saida, data, grupo, total_pdf, titulo=None, vasilhames=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Carga"

    # Monta a folha: títulos de seção (quando há) + itens.
    com_secao = any(len(l) > 5 for l in linhas)
    folha = []
    sec_atual = None
    for l in linhas:
        if com_secao and l[5] != sec_atual:
            sec_atual = l[5]
            folha.append(("secao", TITULO_SECAO.get(sec_atual, sec_atual)))
        folha.append(("item", l))

    # Linhas dos vasilhames: no fim da seção onde estão os produtos de cada um
    # (Refri 2L, KS, Guaraná em RETORNÁVEIS; cervejas em CERVEJAS RETORNÁVEIS).
    # Sem seções, logo depois do último produto. Não entram no TOTAL.
    onde = {}
    for v in vasilhames or []:
        pos = [i for i, (t, c) in enumerate(folha) if t == "item" and c[4] in v["chaves"]]
        if not pos:
            continue
        fim = pos[-1] + 1
        if com_secao:
            while fim < len(folha) and folha[fim][0] == "item":
                fim += 1
        onde.setdefault(fim, []).append(("vasilhame", texto_vasilhame(v)))
    for fim in sorted(onde, reverse=True):  # de trás para frente, para não mudar as posições
        folha[fim:fim] = onde[fim]

    n = len(folha)
    linha_total = n + 2
    por_pagina = math.ceil((n + 1) / 2)
    altura = min(20.0, 730.0 / (por_pagina + 1))
    fonte = max(7, min(12, round(altura * 0.62)))
    aviso_fonte = altura < 11

    borda = Side(style="thin", color="000000")
    caixa = Border(left=borda, right=borda, top=borda, bottom=borda)
    fundo_secao = PatternFill("solid", start_color="D9D9D9", end_color="D9D9D9")

    # Coluna C: quadradinho para conferir a carga (entra na impressão).
    # Colunas extras (fora da área de impressão) com as quantidades numéricas.
    max_partes = max(len(l[2]) for l in linhas)
    col_ini = 4
    col_fim = col_ini + max_partes - 1
    letra_ini, letra_fim = get_column_letter(col_ini), get_column_letter(col_fim)

    ws.append(["Produto", "Quantidade", "OK"])
    linhas_secao = []
    linhas_novas = []
    linhas_vasilhame = []
    for tipo, conteudo in folha:
        if tipo == "secao":
            ws.append([conteudo])
            linhas_secao.append(ws.max_row)
            continue
        if tipo == "vasilhame":
            ws.append([conteudo[0], conteudo[1], CAIXINHA])
            linhas_vasilhame.append(ws.max_row)
            continue
        l = conteudo
        ws.append([l[0], l[1], CAIXINHA] + l[2])
        if len(l) > 6 and l[6]:
            linhas_novas.append(ws.max_row)

    # Linha de total
    ws.cell(row=linha_total, column=1, value="TOTAL")
    ws.cell(
        row=linha_total,
        column=2,
        value=f"=SUM({letra_ini}2:{letra_fim}{n + 1})",
    )

    # As linhas de vasilhame são em negrito (letra mais larga): conta com uma folga a mais.
    vas = [c for t, c in folha if t == "vasilhame"]
    larg_a = min(60, max([len(l[0]) + 2 for l in linhas] + [math.ceil(len(c[0]) * 1.12) + 3 for c in vas]))
    larg_b = min(34, max([len(l[1]) + 2 for l in linhas + [("", "Quantidade", [], 0, "")]] + [len(c[1]) + 4 for c in vas]))
    # A + B + C precisam caber na largura do A4 em pé: se passar, encolhe só a coluna A.
    larg_a = max(min(larg_a, LARGURA_MAX_FOLHA - LARGURA_CAIXINHA - max(larg_b, 12)), 12)
    ws.column_dimensions["A"].width = larg_a
    ws.column_dimensions["B"].width = max(larg_b, 12)
    ws.column_dimensions["C"].width = LARGURA_CAIXINHA

    for r in range(1, linha_total + 1):
        ws.row_dimensions[r].height = altura
        negrito = r in (1, linha_total) or r in linhas_secao or r in linhas_vasilhame
        for c in (1, 2, 3):
            cel = ws.cell(row=r, column=c)
            # O ☐ não existe na Arial; a Segoe UI Symbol (Windows) desenha ele certinho.
            fonte_nome = "Segoe UI Symbol" if c == 3 and r > 1 else "Arial"
            cel.font = Font(name=fonte_nome, size=fonte, bold=negrito, italic=r in linhas_novas and c < 3)
            cel.border = caixa
            cel.alignment = Alignment(
                horizontal="center" if c >= 2 or negrito else "left",
                vertical="center",
                indent=0 if c >= 2 or negrito else 1,
            )
    fundo_vasilhame = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    for r in linhas_vasilhame:
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).fill = fundo_vasilhame
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for r in linhas_secao:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).fill = fundo_secao
    ws.cell(row=linha_total, column=1).alignment = Alignment(
        horizontal="right", vertical="center", indent=1
    )

    # Colunas de apoio em cinza (não entram na impressão).
    cinza = Font(name="Arial", size=9, color="808080")
    ws.cell(row=1, column=col_ini, value="Quantidades numéricas por embalagem (não imprime)")
    ws.cell(row=1, column=col_ini).font = cinza
    for r in range(2, n + 2):
        for c in range(col_ini, col_fim + 1):
            ws.cell(row=r, column=c).font = cinza
    for c in range(col_ini, col_fim + 1):
        ws.column_dimensions[get_column_letter(c)].width = 9

    col_conf = col_fim + 2
    L1, L2 = get_column_letter(col_conf), get_column_letter(col_conf + 1)
    ws.column_dimensions[L1].width = 22
    ws.column_dimensions[L2].width = 14
    if total_pdf is not None:
        ws.cell(row=1, column=col_conf, value="Total do PDF (sistema)")
        ws.cell(row=1, column=col_conf + 1, value=int(total_pdf))
        ws.cell(row=2, column=col_conf, value="Confere com a folha?")
        ws.cell(
            row=2,
            column=col_conf + 1,
            value=f'=IF(B{linha_total}={L2}1,"OK","DIFERENTE")',
        )
        for r in (1, 2):
            for c in (col_conf, col_conf + 1):
                ws.cell(row=r, column=c).font = Font(name="Arial", size=9, bold=True, color="808080")

    # Quebra de página na metade (frente e verso), sem deixar título de seção sozinho no fim da página.
    # Se a quebra cair logo depois de um título de seção, sobe uma linha.
    quebra = por_pagina + 1
    while quebra in linhas_secao and quebra > 2:
        quebra -= 1

    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.page_margins.header = ws.page_margins.footer = 0.2
    ws.print_title_rows = "1:1"
    ws.print_area = f"A1:C{linha_total}"
    ws.print_options.horizontalCentered = True
    if n > 1:
        ws.row_breaks.append(Break(id=quebra))
    cab = titulo or f"Carga {data or ''}  {grupo or ''}".strip()
    ws.oddHeader.center.text = cab
    ws.oddHeader.center.size = 9
    wb.calculation.fullCalcOnLoad = True
    try:
        wb.save(saida)
    except PermissionError:
        # O arquivo do dia está aberto no Excel: salva com outro nome.
        base, ext = os.path.splitext(saida)
        n_alt = 2
        while True:
            alt = f"{base}_{n_alt}{ext}"
            try:
                wb.save(alt)
                break
            except PermissionError:
                n_alt += 1
                if n_alt > 20:
                    raise RuntimeError(
                        "Não consegui salvar. Feche o arquivo Carga_... no Excel e tente de novo."
                    )
        saida = alt
    return por_pagina, aviso_fonte, saida


# --------------------------------------------------------------------------
DIAS_SEMANA = {
    "SEGUNDA": "Segunda",
    "TERCA": "Terça",
    "QUARTA": "Quarta",
    "QUINTA": "Quinta",
    "SEXTA": "Sexta",
    "SABADO": "Sábado",
    "DOMINGO": "Domingo",
}


def _sem_acento(t):
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")


def rotulo_do_arquivo(pdf):
    """Lê o nome do PDF (ex.: MAPA_PALMEIRAS_QUARTA_2026_11_10.pdf) e devolve
    (rotulo, dia_da_semana): ('Palmeiras', 'Quarta'). MAPA_A_2026_11_10.pdf -> ('A', None)."""
    nome = os.path.splitext(os.path.basename(pdf))[0]
    rotulo, dia = [], None
    for parte in re.split(r"[\s_\-]+", nome):
        limpa = re.sub(r"[^\w]", "", parte)
        if not limpa or limpa.isdigit():
            continue
        u = _sem_acento(limpa).upper()
        if u in ("MAPA", "MAPAS", "FEIRA"):
            continue
        if u in DIAS_SEMANA:
            dia = DIAS_SEMANA[u]
            continue
        rotulo.append(limpa)
    return " ".join(rotulo).title(), dia


def nome_seguro(t):
    return re.sub(r'[\\/:*?"<>|]+', " ", t).strip()


def monta_rotulo(pdf, grupo_pdf):
    """Rótulo do mapa: o nome escrito no arquivo (Palmeiras) ou, se for só uma letra
    (MAPA_A), o 'Grupo do Pedido' do relatório (Grupo A)."""
    rot, dia = rotulo_do_arquivo(pdf)
    if len(rot) >= 2:
        nome = rot
    elif grupo_pdf:
        nome = grupo_pdf.strip().title()
    elif rot:
        nome = f"Grupo {rot.upper()}"
    else:
        nome = "Carga"
    return nome_seguro(nome), dia


# --------------------------------------------------------------------------
# Histórico: um arquivo só, na pasta das folhas, com uma linha por folha gerada
# (a mais nova sempre no fim, então fica em ordem cronológica).
NOME_HISTORICO = "Histórico de cargas.xlsx"
CABECALHO_HISTORICO = [
    "Gerado em",
    "Data do mapa",
    "Mapa",
    "Linhas na folha",
    "Volumes",
    "Total do sistema",
    "Conferência",
    "Arquivo",
]
LARGURAS_HISTORICO = [19, 14, 26, 15, 10, 16, 14, 50]


def _numero(v):
    if v is None:
        return None
    return int(v) if float(v).is_integer() else float(v)


def registra_historico(pasta, saida, mapa, data, n_linhas, soma, total_pdf, ok, agora=None):
    """Acrescenta uma linha no 'Histórico de cargas.xlsx' da pasta (cria o arquivo na
    primeira vez) e devolve o caminho dele. Se o histórico estiver aberto no Excel,
    levanta PermissionError; o arquivo existente nunca é apagado nem recriado."""
    caminho = os.path.join(pasta, NOME_HISTORICO)
    agora = agora or datetime.datetime.now().replace(microsecond=0)
    if os.path.exists(caminho):
        wb = load_workbook(caminho)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Histórico"
        ws.append(CABECALHO_HISTORICO)
        for c, larg in enumerate(LARGURAS_HISTORICO, start=1):
            ws.cell(row=1, column=c).font = Font(name="Arial", size=10, bold=True)
            ws.column_dimensions[get_column_letter(c)].width = larg
        ws.freeze_panes = "A2"
        # Se imprimir: deitado, todas as colunas numa página de largura.
        ws.page_setup.paperSize = 9
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = "1:1"

    if total_pdf is None:
        conferencia = "SEM TOTAL"
    else:
        conferencia = "OK" if ok else "NÃO BATEU"
    arquivo = os.path.basename(saida)
    ws.append(
        [agora, data or "", mapa, n_linhas, _numero(soma), _numero(total_pdf), conferencia, arquivo]
    )
    r = ws.max_row
    for c in range(1, len(CABECALHO_HISTORICO) + 1):
        ws.cell(row=r, column=c).font = Font(name="Arial", size=10)
    ws.cell(row=r, column=1).number_format = "dd/mm/yyyy hh:mm"
    if conferencia == "NÃO BATEU":
        ws.cell(row=r, column=7).font = Font(name="Arial", size=10, bold=True, color="B00020")
    # Clicar no nome abre a folha (ela fica na mesma pasta do histórico).
    ws.cell(row=r, column=8).hyperlink = arquivo
    ws.cell(row=r, column=8).font = Font(name="Arial", size=10, color="0563C1", underline="single")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CABECALHO_HISTORICO))}{r}"
    wb.save(caminho)
    return caminho


def le_historico(pasta):
    """Lê o 'Histórico de cargas.xlsx' da pasta e devolve a lista de folhas geradas,
    em ordem cronológica (a mais nova no fim). Cada item é um dict com as colunas do
    histórico e 'caminho' = caminho completo da folha. Sem histórico na pasta: []."""
    caminho = os.path.join(pasta, NOME_HISTORICO)
    if not os.path.exists(caminho):
        return []
    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()
    if not linhas:
        return []
    cab = [_como_texto(c) for c in linhas[0]]
    folhas = []
    for lin in linhas[1:]:
        if not lin or all(c is None for c in lin):
            continue
        d = {nome: (lin[i] if i < len(lin) else None) for i, nome in enumerate(cab)}
        arquivo = _como_texto(d.get("Arquivo"))
        d["caminho"] = os.path.join(pasta, arquivo) if arquivo else None
        folhas.append(d)
    return folhas


# --------------------------------------------------------------------------
def pasta_do_app():
    """Pasta onde estão o script (ou o .exe), usada para achar apelidos.csv e o cadastro."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def acha_mapa_csv():
    """mapa_ordem.csv: primeiro ao lado do programa (editável), depois o que vai dentro do .exe."""
    cands = [os.path.join(pasta_do_app(), "mapa_ordem.csv")]
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        cands.append(os.path.join(sys._MEIPASS, "mapa_ordem.csv"))
    cands.append(os.path.join(os.getcwd(), "mapa_ordem.csv"))
    for c in cands:
        if os.path.exists(c):
            return c
    return None


def le_ordem_mapa(caminho):
    """chave do produto -> (posição no mapa atual, seção)."""
    ordem = {}
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            try:
                pos = float(r["ordem"])
            except (KeyError, ValueError):
                continue
            k = chave(r.get("base_cadastro", ""))
            sec = (r.get("secao") or "").strip().upper() or "OUTROS"
            if sec not in ORDEM_SECOES:
                sec = "OUTROS"
            if k and (k not in ordem or pos < ordem[k][0]):
                ordem[k] = (pos, sec)
    return ordem


def acha_cadastro(pasta_pdf, pasta_script):
    for pasta in (pasta_pdf, os.getcwd(), pasta_script):
        achados = glob.glob(os.path.join(pasta, "relatorio_Lista_de_produtos*.xlsx"))
        if achados:
            return max(achados, key=os.path.getmtime)
    return None


def processar(pdf, cadastro=None, apelidos_csv=None, ordem="mapa", saida_dir=None, mapa_csv=None):
    """Faz tudo: lê PDF e cadastro, monta a folha e salva o Excel.
    Devolve um dicionário com o resultado (usado pelo terminal e pela janela).
    Levanta RuntimeError com uma mensagem clara quando não dá para continuar."""
    pasta_pdf = os.path.dirname(os.path.abspath(pdf))
    cadastro = cadastro or acha_cadastro(pasta_pdf, pasta_do_app())
    if not cadastro or not os.path.exists(cadastro):
        raise RuntimeError(
            "Não achei o cadastro (relatorio_Lista_de_produtos...xlsx). Escolha o arquivo do cadastro."
        )
    apelidos_csv = apelidos_csv or os.path.join(pasta_do_app(), "apelidos.csv")

    por_ref, variantes = le_cadastro(cadastro)
    apelidos = le_apelidos(apelidos_csv)
    itens, total_pdf, data, grupo = le_pdf(pdf)
    if not itens:
        raise RuntimeError(
            "Não consegui ler nenhum produto desse PDF. Ele é o 'Relatório de produtos - Pré-Venda' do sistema?"
        )

    ordem_mapa = None
    sem_mapa = False
    if ordem == "mapa":
        mapa_csv = mapa_csv or acha_mapa_csv()
        if mapa_csv:
            ordem_mapa = le_ordem_mapa(mapa_csv)
        else:
            ordem_mapa, sem_mapa = {}, True
    linhas, nao_cad, sem_tipo, fora_do_mapa = monta_linhas(
        itens, por_ref, variantes, apelidos, ordem, ordem_mapa
    )
    vasilhames = conta_vasilhames(itens, por_ref)

    soma = sum(i["qtd"] for i in itens)
    rotulo, dia = monta_rotulo(pdf, grupo)
    data_arq = data.replace("/", "-") if data else "sem-data"
    partes_nome = ["Carga", rotulo] if rotulo != "Carga" else ["Carga"]
    nome_arq = " ".join(partes_nome) + (f" - {dia}" if dia else "") + f" - {data_arq}.xlsx"
    titulo = " ".join(partes_nome) + (f" - {dia}" if dia else "") + (f" - {data}" if data else "")
    saida_dir = saida_dir or pasta_pdf
    os.makedirs(saida_dir, exist_ok=True)
    saida = os.path.join(saida_dir, nome_arq)
    por_pagina, fonte_pequena, saida_final = escreve_xlsx(
        linhas, saida, data, grupo, total_pdf, titulo=titulo, vasilhames=vasilhames
    )
    ok = total_pdf is not None and abs(soma - total_pdf) < 0.001

    # A folha já está salva; um problema no histórico só vira aviso.
    historico, aviso_historico = None, None
    mapa = rotulo + (f" - {dia}" if dia else "")
    try:
        historico = registra_historico(
            saida_dir, saida_final, mapa, data, len(linhas), soma, total_pdf, ok
        )
    except PermissionError:
        aviso_historico = (
            f"Não anotei esta folha no '{NOME_HISTORICO}' porque ele está aberto no Excel. "
            "Feche o histórico e gere a folha de novo para ela entrar na lista."
        )
    except Exception as e:
        aviso_historico = f"Não consegui anotar esta folha no '{NOME_HISTORICO}': {e}"

    return {
        "saida": saida_final,
        "renomeado": saida_final != saida,
        "n_itens_pdf": len(itens),
        "n_linhas": len(linhas),
        "por_pagina": por_pagina,
        "soma": soma,
        "total_pdf": total_pdf,
        "ok": ok,
        "historico": historico,
        "vasilhames": vasilhames,
        "aviso_historico": aviso_historico,
        "nao_cadastrados": nao_cad,
        "sem_tipo": sorted(set(sem_tipo)),
        "fora_do_mapa": [] if sem_mapa else fora_do_mapa,
        "sem_mapa": sem_mapa,
        "fonte_pequena": fonte_pequena,
        "data": data,
        "grupo": grupo,
        "cadastro": cadastro,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("pdf")
    ap.add_argument("--cadastro")
    ap.add_argument("--apelidos")
    ap.add_argument(
        "--ordem",
        choices=["mapa", "volume", "sistema", "alfabetica"],
        default="mapa",
        help="mapa (padrão): seções e ordem do MAPA_ATUAL; volume: maior volume primeiro; sistema: ordem do PDF; alfabetica",
    )
    ap.add_argument("--saida", help="pasta de saída (padrão: a mesma do PDF)")
    ap.add_argument("--mapa", help="arquivo mapa_ordem.csv (padrão: o que está ao lado do programa)")
    a = ap.parse_args()

    try:
        r = processar(a.pdf, a.cadastro, a.apelidos, a.ordem, a.saida, a.mapa)
    except RuntimeError as e:
        sys.exit(str(e))

    if r["renomeado"]:
        print("Atenção: o arquivo do dia estava aberto no Excel; salvei com outro nome.")
    print(f"PDF: {r['n_itens_pdf']} itens, soma das quantidades = {r['soma']:g}")
    print(f"Folha: {r['n_linhas']} linhas ({r['por_pagina']} por página, frente e verso)")
    if r["ok"]:
        print(f"CONFERÊNCIA OK: total do PDF = {r['total_pdf']:g}")
    else:
        print("!!! CONFERÊNCIA NÃO BATEU !!!")
        print(f"    soma lida = {r['soma']:g} | total no rodapé do PDF = {r['total_pdf']}")
        print("    Confira a folha antes de imprimir.")
    if r["nao_cadastrados"]:
        print("\nAtenção - produtos que NÃO estão no cadastro (use o cadastro mais novo):")
        for n in r["nao_cadastrados"]:
            print("  -", n)
    if r["sem_tipo"]:
        print("\nAtenção - embalagem não reconhecida (marcada com '?'):")
        for n in r["sem_tipo"]:
            print("  -", n)
    if r.get("sem_mapa"):
        print("\nAtenção: não achei o mapa_ordem.csv; separei pelas seções, mas sem a ordem do mapa.")
    if r["fora_do_mapa"]:
        print("\nItens que NÃO estão no mapa atual (ganharam uma linha no fim da seção, em itálico):")
        for n in r["fora_do_mapa"]:
            print("  -", n)
    if r["fonte_pequena"]:
        print("\nAtenção: muitos itens, a fonte ficou pequena. Considere imprimir em 3 páginas.")
    if r["vasilhames"]:
        print()
    for v in r["vasilhames"]:
        print(" = ".join(texto_vasilhame(v)))
        if v["sem_tamanho"]:
            print(f"Atenção - {v['nome']} sem tamanho de caixa no cadastro (não entrou nas garrafas):")
            for n in v["sem_tamanho"]:
                print("  -", n)
    if r["aviso_historico"]:
        print("\nAtenção: " + r["aviso_historico"])
    print(f"\nArquivo gerado: {r['saida']}")
    if r["historico"]:
        print(f"Anotado no histórico: {r['historico']}")
    sys.exit(0 if r["ok"] else 1)


if __name__ == "__main__":
    main()
