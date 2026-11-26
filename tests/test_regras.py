"""
Testes das regras da folha de carga. Rodar com:  python -m pytest
Servem para garantir que uma mudança nova não quebre o que já funciona.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import gerar_folha_carga as core  # noqa: E402


def linha_de(itens, cadastro, ordem="sistema", ordem_mapa=None, apelidos=None):
    """Monta as linhas da folha a partir de itens do PDF e de um cadastro {ref: nome}."""
    por_ref, variantes = {}, {}
    for ref, nome in cadastro.items():
        base, tipo, arr = core.separa_nome(nome)
        por_ref[ref] = (base, tipo, arr)
        if tipo and arr:
            variantes.setdefault((core.chave(base), tipo), set()).add(arr)
    return core.monta_linhas(itens, por_ref, variantes, apelidos or {}, ordem, ordem_mapa)


def test_separa_nome():
    assert core.separa_nome("SPRITE 350ML FD 12X1") == ("SPRITE 350ML", "FD", 12)
    assert core.separa_nome("CACHACA 51 965ML UN") == ("CACHACA 51 965ML", "UN", None)
    assert core.separa_nome("PRODUTO SEM EMBALAGEM") == ("PRODUTO SEM EMBALAGEM", None, None)


def test_fardo_de_12_sai_como_fd_e_outros_como_c():
    cadastro = {"1001": "SPRITE 350ML FD 12X1", "1002": "SPRITE 350ML FD 6X1"}
    itens = [
        {"ref": "1001", "nome": "SPRITE 350ML FD 12X1", "qtd": 12.0},
        {"ref": "1002", "nome": "SPRITE 350ML FD 6X1", "qtd": 2.0},
    ]
    linhas, *_ = linha_de(itens, cadastro)
    assert linhas[0][0] == "Sprite 350ML"
    assert linhas[0][1] == "12(FD); 2(C/6)"


def test_fardo_de_6_abaixo_de_1l_nunca_vira_fd():
    # Caso real: Sprite Zero 350ML só tem FD 6X1 no cadastro.
    cadastro = {"2001": "SPRITE ZERO 350ML FD 6X1"}
    itens = [{"ref": "2001", "nome": "SPRITE ZERO 350ML FD 6X1", "qtd": 21.0}]
    linhas, *_ = linha_de(itens, cadastro)
    assert linhas[0][1] == "21(C/6)"


def test_fardo_de_6_de_2l_e_o_fd_normal():
    cadastro = {"3001": "COCA COLA 2L FD 6X1"}
    itens = [{"ref": "3001", "nome": "COCA COLA 2L FD 6X1", "qtd": 5.0}]
    linhas, *_ = linha_de(itens, cadastro)
    assert linhas[0][1] == "5(FD)"


def test_un_e_cx_na_mesma_linha():
    cadastro = {"4001": "CACHACA 51 965ML UN", "4002": "CACHACA 51 965ML CX 12X1"}
    itens = [
        {"ref": "4001", "nome": "CACHACA 51 965ML UN", "qtd": 2.0},
        {"ref": "4002", "nome": "CACHACA 51 965ML CX 12X1", "qtd": 4.0},
    ]
    linhas, *_ = linha_de(itens, cadastro)
    assert linhas[0][1] == "4(CX); 2(UN)"


def test_produto_fora_do_cadastro_gera_aviso():
    itens = [{"ref": "9999", "nome": "PRODUTO NOVO 600ML FD 12X1", "qtd": 1.0}]
    linhas, nao_cad, sem_tipo, fora = linha_de(itens, {})
    assert nao_cad == ["PRODUTO NOVO 600ML FD 12X1"]
    assert linhas[0][1] == "1(FD)"


def test_apelido_troca_o_nome():
    cadastro = {"5001": "CERV ANTARTICA PILSEN 350ML FD 12X1"}
    itens = [{"ref": "5001", "nome": "CERV ANTARTICA PILSEN 350ML FD 12X1", "qtd": 3.0}]
    apelidos = {core.chave("CERV ANTARTICA PILSEN 350ML"): "Antarctica Pilsen 350ML"}
    linhas, *_ = linha_de(itens, cadastro, apelidos=apelidos)
    assert linhas[0][0] == "Antarctica Pilsen 350ML"


def test_secao_por_nome():
    assert core.secao_da_base("AGUA CRYSTAL S/ GAS 500ML") == "AGUAS"
    assert core.secao_da_base("MONSTER ENERGY 473ML") == "ENERGETICOS"
    assert core.secao_da_base("CERV SKOL 350ML") == "CERVEJAS"
    assert core.secao_da_base("CERV SKOL 600ML RET") == "CERVEJAS RET"
    assert core.secao_da_base("VODKA SMIRNOFF 998ML") == "DESTILADOS"
    assert core.secao_da_base("GUARANA ANTARCTICA 2L") == "2L"


def test_ordem_do_mapa_e_item_fora_do_mapa():
    cadastro = {
        "1": "COCA COLA 600ML FD 12X1",
        "2": "FANTA UVA 600ML FD 12X1",
        "3": "GUARANA NOVO 600ML FD 12X1",
    }
    itens = [
        {"ref": "3", "nome": "", "qtd": 1.0},
        {"ref": "2", "nome": "", "qtd": 1.0},
        {"ref": "1", "nome": "", "qtd": 1.0},
    ]
    ordem_mapa = {
        core.chave("COCA COLA 600ML"): (1.0, "600ML"),
        core.chave("FANTA UVA 600ML"): (2.0, "600ML"),
    }
    linhas, _, _, fora = linha_de(itens, cadastro, ordem="mapa", ordem_mapa=ordem_mapa)
    assert [l[0] for l in linhas] == ["Coca Cola 600ML", "Fanta Uva 600ML", "Guarana Novo 600ML"]
    assert fora == ["Guarana Novo 600ML"]


def test_mapa_ordem_csv_do_repositorio_carrega():
    pasta = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ordem = core.le_ordem_mapa(os.path.join(pasta, "mapa_ordem.csv"))
    assert len(ordem) > 250
    assert ordem[core.chave("COCA COLA 600ML")] == (1.0, "600ML")


def test_escreve_xlsx(tmp_path):
    cadastro = {"1": "COCA COLA 600ML FD 12X1"}
    itens = [{"ref": "1", "nome": "", "qtd": 3.0}]
    linhas, *_ = linha_de(itens, cadastro, ordem="mapa", ordem_mapa={core.chave("COCA COLA 600ML"): (1.0, "600ML")})
    saida = str(tmp_path / "Carga.xlsx")
    _, _, final = core.escreve_xlsx(linhas, saida, "25/11/2026", "A", 3.0)
    assert os.path.exists(final)


def test_quadradinho_para_conferir_a_carga(tmp_path):
    from openpyxl import load_workbook

    cadastro = {"1": "COCA COLA 600ML FD 12X1", "2": "FANTA UVA 600ML FD 12X1"}
    itens = [{"ref": "1", "nome": "", "qtd": 3.0}, {"ref": "2", "nome": "", "qtd": 2.0}]
    ordem_mapa = {
        core.chave("COCA COLA 600ML"): (1.0, "600ML"),
        core.chave("FANTA UVA 600ML"): (2.0, "600ML"),
    }
    linhas, *_ = linha_de(itens, cadastro, ordem="mapa", ordem_mapa=ordem_mapa)
    _, _, final = core.escreve_xlsx(linhas, str(tmp_path / "Carga.xlsx"), "25/11/2026", "A", 5.0)
    ws = load_workbook(final).active

    # Linha 1: cabeçalho; linha 2: título da seção; linhas 3 e 4: itens.
    assert [ws.cell(row=1, column=c).value for c in (1, 2, 3)] == ["Produto", "Quantidade", "OK"]
    assert ws.cell(row=2, column=3).value is None  # título de seção não tem quadradinho
    assert [ws.cell(row=r, column=3).value for r in (3, 4)] == ["☐", "☐"]
    # O quadradinho imprime; as colunas de apoio continuam fora da impressão.
    assert ws.print_area.endswith("$A$1:$C$5")
    # O total continua somando só as quantidades numéricas (agora a partir da coluna D).
    assert ws.cell(row=5, column=2).value == "=SUM(D2:D4)"
    assert ws.cell(row=3, column=4).value == 3


def test_historico_cria_e_acrescenta_em_ordem_cronologica(tmp_path):
    import datetime

    from openpyxl import load_workbook

    pasta = str(tmp_path)
    dia1 = datetime.datetime(2026, 11, 25, 7, 42)
    dia2 = datetime.datetime(2026, 11, 26, 7, 35)
    core.registra_historico(pasta, "Carga Palmeiras - 25-11-2026.xlsx", "Palmeiras", "25/11/2026",
                            48, 312.0, 312.0, True, agora=dia1)
    caminho = core.registra_historico(pasta, "Carga Grupo A - 26-11-2026.xlsx", "Grupo A", "26/11/2026",
                                      55, 270.0, 271.0, False, agora=dia2)

    ws = load_workbook(caminho).active
    linhas = [[c.value for c in row] for row in ws.iter_rows()]
    assert linhas[0] == core.CABECALHO_HISTORICO
    assert linhas[1] == [dia1, "25/11/2026", "Palmeiras", 48, 312, 312, "OK", "Carga Palmeiras - 25-11-2026.xlsx"]
    assert linhas[2] == [dia2, "26/11/2026", "Grupo A", 55, 270, 271, "NÃO BATEU", "Carga Grupo A - 26-11-2026.xlsx"]
    assert len(linhas) == 3


def _processar_com_pdf_falso(tmp_path, monkeypatch):
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.append(["referencia", "nome"])
    wb.active.append(["1", "COCA COLA 600ML FD 12X1"])
    cadastro = str(tmp_path / "relatorio_Lista_de_produtos.xlsx")
    wb.save(cadastro)
    itens = [{"ref": "1", "nome": "COCA COLA 600ML FD 12X1", "qtd": 3.0}]
    monkeypatch.setattr(core, "le_pdf", lambda pdf: (itens, 3.0, "26/11/2026", None))
    pdf = str(tmp_path / "MAPA_PALMEIRAS_QUINTA_2026_11_26.pdf")
    return core.processar(pdf, cadastro=cadastro)


def test_processar_anota_no_historico(tmp_path, monkeypatch):
    from openpyxl import load_workbook

    r = _processar_com_pdf_falso(tmp_path, monkeypatch)
    assert r["aviso_historico"] is None
    assert r["historico"] == str(tmp_path / core.NOME_HISTORICO)
    ws = load_workbook(r["historico"]).active
    assert ws.max_row == 2
    assert ws.cell(row=2, column=3).value == "Palmeiras - Quinta"
    assert ws.cell(row=2, column=7).value == "OK"
    assert ws.cell(row=2, column=8).value == os.path.basename(r["saida"])


def test_historico_aberto_no_excel_nao_impede_a_folha(tmp_path, monkeypatch):
    def aberto(*a, **k):
        raise PermissionError("arquivo em uso")

    monkeypatch.setattr(core, "registra_historico", aberto)
    r = _processar_com_pdf_falso(tmp_path, monkeypatch)
    assert os.path.exists(r["saida"])  # a folha saiu
    assert r["historico"] is None
    assert "aberto no Excel" in r["aviso_historico"]


CADASTRO_KS = {
    "10": "COCA COLA KS 290ML RET CX 24X1",
    "11": "COCA COLA ZERO KS 290ML RET CX 24X1",
    "12": "FANTA LARANJA KS 290ML RET CX 24X1",
    "13": "SPRITE KS 290ML RET UN",
    "20": "COCA COLA 2L RET FD 6X1",
    "30": "GUARANA ANTARCTICA 300ML RET CX 24X1",
}
ITENS_KS = [
    {"ref": "20", "nome": "", "qtd": 4.0},
    {"ref": "10", "nome": "", "qtd": 5.0},
    {"ref": "11", "nome": "", "qtd": 2.0},
    {"ref": "12", "nome": "", "qtd": 3.0},
    {"ref": "13", "nome": "", "qtd": 5.0},
    {"ref": "30", "nome": "", "qtd": 7.0},
]


def _por_ref(cadastro):
    return {ref: core.separa_nome(nome) for ref, nome in cadastro.items()}


def test_vasilhame_ks_conta_coca_fanta_sprite_ks():
    v = core.conta_vasilhame_ks(ITENS_KS, _por_ref(CADASTRO_KS))
    # 5 + 2 + 3 engradados de 24 = 240 garrafas, mais 5 garrafas soltas de Sprite KS.
    # Coca 2L e Guaraná 300ML são retornáveis mas NÃO são KS: não contam.
    assert (v["cx"], v["un"], v["garrafas"]) == (10, 5, 245)
    assert v["sem_tamanho"] == []
    assert core.texto_vasilhame(v) == ("Vasilhame KS que volta (245 garrafas)", "10 CX + 5 UN")


def test_sem_ks_nao_tem_vasilhame():
    itens = [{"ref": "20", "nome": "", "qtd": 4.0}]
    assert core.conta_vasilhame_ks(itens, _por_ref(CADASTRO_KS)) is None


def test_ks_sem_tamanho_de_caixa_gera_aviso():
    itens = [{"ref": "99", "nome": "COCA COLA KS 290ML RET CX", "qtd": 2.0}]
    v = core.conta_vasilhame_ks(itens, {})
    assert v["sem_tamanho"] == ["Coca Cola KS 290ML RET"]
    assert v["garrafas"] == 0


def test_linha_do_vasilhame_no_fim_de_retornaveis_e_fora_do_total(tmp_path):
    from openpyxl import load_workbook

    pasta = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ordem_mapa = core.le_ordem_mapa(os.path.join(pasta, "mapa_ordem.csv"))
    cadastro = dict(CADASTRO_KS, **{"40": "CERV SKOL 600ML RET CX 24X1"})
    itens = ITENS_KS + [{"ref": "40", "nome": "", "qtd": 1.0}]
    linhas, *_ = linha_de(itens, cadastro, ordem="mapa", ordem_mapa=ordem_mapa)
    v = core.conta_vasilhame_ks(itens, _por_ref(cadastro))
    _, _, final = core.escreve_xlsx(linhas, str(tmp_path / "Carga.xlsx"), "26/11/2026", "A", 27.0, vasilhame=v)

    ws = load_workbook(final).active
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    i_vas = col_a.index("Vasilhame KS que volta (245 garrafas)")
    # Fica depois de todos os retornáveis (incluindo o Guaraná 300ML) e antes da próxima seção.
    assert col_a[i_vas - 1] == "Guarana Antarctica 300ML RET"
    assert col_a[i_vas + 1] == "CERVEJAS RETORNÁVEIS"
    linha = i_vas + 1
    assert ws.cell(row=linha, column=2).value == "10 CX + 5 UN"
    assert ws.cell(row=linha, column=4).value is None  # não soma no TOTAL
    # O TOTAL continua batendo com o sistema: 4 + 5 + 2 + 3 + 5 + 7 + 1 = 27.
    numeros = [ws.cell(row=r, column=4).value for r in range(2, ws.max_row)]
    assert sum(n for n in numeros if isinstance(n, (int, float))) == 27
