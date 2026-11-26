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


CADASTRO_RET = {
    "10": "COCA COLA KS 290ML RET CX 24X1",
    "11": "COCA COLA ZERO KS 290ML RET CX 24X1",
    "12": "FANTA LARANJA KS 290ML RET CX 24X1",
    "13": "SPRITE KS 290ML RET UN",
    "20": "COCA COLA 2L RET FD 6X1",
    "21": "COCA COLA ZERO 2L RET FD 6X1",
    "22": "FANTA LARANJA 2L RET FD 6X1",
    "30": "GUARANA ANTARCTICA 300ML RET CX 24X1",
    "31": "GUARANA ANTARCTICA 300ML RET UN",
    "40": "CERV ANTARCTICA 300ML RET. CX 23X1",
}
ITENS_RET = [
    {"ref": "20", "nome": "", "qtd": 4.0},
    {"ref": "21", "nome": "", "qtd": 2.0},
    {"ref": "22", "nome": "", "qtd": 1.0},
    {"ref": "10", "nome": "", "qtd": 5.0},
    {"ref": "11", "nome": "", "qtd": 2.0},
    {"ref": "12", "nome": "", "qtd": 3.0},
    {"ref": "13", "nome": "", "qtd": 5.0},
    {"ref": "30", "nome": "", "qtd": 7.0},
    {"ref": "31", "nome": "", "qtd": 3.0},
    {"ref": "40", "nome": "", "qtd": 1.0},
]


def _por_ref(cadastro):
    return {ref: core.separa_nome(nome) for ref, nome in cadastro.items()}


def _vasilhames(itens=ITENS_RET, cadastro=CADASTRO_RET):
    return {v["nome"]: v for v in core.conta_vasilhames(itens, _por_ref(cadastro))}


def test_tipo_de_vasilhame_pelo_nome():
    assert core.tipo_vasilhame("COCA COLA KS 290ML RET") == "KS"
    assert core.tipo_vasilhame("SPRITE KS 290ML RET") == "KS"
    assert core.tipo_vasilhame("COCA COLA 2L RET") == "Refri 2L"
    assert core.tipo_vasilhame("COCA COLA ZERO 2L RET") == "Refri 2L"
    assert core.tipo_vasilhame("FANTA LARANJA 2L RET") == "Refri 2L"
    assert core.tipo_vasilhame("FANTA UVA 2L RET") == "Refri 2L"
    assert core.tipo_vasilhame("GUARANA ANTARCTICA 300ML RET") == "Guaraná 300ML"
    assert core.tipo_vasilhame("GUARANÁ ANTARCTICA 300ML RET") == "Guaraná 300ML"
    # Não são vasilhame (por enquanto):
    assert core.tipo_vasilhame("COCA COLA 2L") is None  # 2L descartável
    # Cerveja litrinho 300ML é vasilhame próprio, não entra no Guaraná 300ML:
    assert core.tipo_vasilhame("CERV ANTARCTICA 300ML RET.") == "Cerveja 300ML"
    assert core.tipo_vasilhame("CERV SKOL 300ML RET.") == "Cerveja 300ML"
    assert core.tipo_vasilhame("CERV BRAHMA CHOPP 300ML RET.") == "Cerveja 300ML"
    assert core.tipo_vasilhame("CERV ORIGINAL LITRINHO 300ML") == "Cerveja 300ML"  # sem RET no nome
    # Cervejas 600ML e litrão retornáveis:
    assert core.tipo_vasilhame("CERV SKOL 600ML RET") == "Cerveja 600ML"
    assert core.tipo_vasilhame("CERV BRAHMA CHOPP 600ML RET.") == "Cerveja 600ML"
    assert core.tipo_vasilhame("CERV ANTARTICA SUB-ZERO 600ML RET") == "Cerveja 600ML"
    assert core.tipo_vasilhame("CERV SPATEN RET. 600ML") == "Cerveja 600ML"  # RET antes do 600ML
    assert core.tipo_vasilhame("CERV SKOL LITRÃO 1L RET.") == "Litrão 1L"
    assert core.tipo_vasilhame("CERV BRAHMA LITRAO 1L RET.") == "Litrão 1L"
    assert core.tipo_vasilhame("CERV SPATEN DESC 600ML") is None  # descartável
    # Amstel 600ML volta junto com as da Ambev; Heineken 600ML e Amstel Litrão têm casco próprio.
    assert core.tipo_vasilhame("CERV AMSTEL 600ML RET.") == "Cerveja 600ML"
    assert core.tipo_vasilhame("CERV HEINEKEN 600ML") == "Heineken 600ML"
    assert core.tipo_vasilhame("CERV HEINEKEN 600ML RET") == "Heineken 600ML"
    assert core.tipo_vasilhame("CERV HEINEKEN DESC 600ML") is None
    assert core.tipo_vasilhame("CERV AMSTEL LITRÃO 1L RET.") == "Amstel Litrão 1L"
    assert core.tipo_vasilhame("CERV GLACIAL LITRÃO 1L RET.") == "Litrão 1L"


def test_vasilhame_ks_conta_coca_fanta_sprite_ks():
    v = _vasilhames()["KS"]
    # 5 + 2 + 3 engradados de 24 = 240 garrafas, mais 5 garrafas soltas de Sprite KS.
    assert (v["por_tipo"], v["garrafas"]) == ({"CX": 10, "UN": 5}, 245)
    assert core.texto_vasilhame(v) == ("Vasilhame KS que volta (245 garrafas)", "10 CX + 5 UN")


def test_vasilhame_refri_2l_casco_de_12():
    v = _vasilhames()["Refri 2L"]
    # Coca 4 FD + Coca Zero 2 FD + Fanta 1 FD, de 6 garrafas = 42 garrafas
    # = 3 cascos de 12 + 6 garrafas soltas.
    assert core.texto_vasilhame(v) == ("Vasilhame Refri 2L que volta (42 garrafas)", "3 CX + 6 UN")


def test_vasilhame_refri_2l_conta_exata_e_unidades():
    cadastro = {"20": "COCA COLA 2L RET FD 6X1", "23": "COCA COLA 2L RET UN"}
    itens = [{"ref": "20", "nome": "", "qtd": 4.0}]
    v = core.conta_vasilhames(itens, _por_ref(cadastro))[0]
    assert core.texto_vasilhame(v) == ("Vasilhame Refri 2L que volta (24 garrafas)", "2 CX")
    itens = [{"ref": "23", "nome": "", "qtd": 5.0}]
    v = core.conta_vasilhames(itens, _por_ref(cadastro))[0]
    assert core.texto_vasilhame(v) == ("Vasilhame Refri 2L que volta (5 garrafas)", "5 UN")


def test_vasilhame_guarana_300ml():
    v = _vasilhames()["Guaraná 300ML"]
    # 7 engradados de 24 = 168, mais 3 soltas. A cerveja Antarctica 300ML não conta.
    assert core.texto_vasilhame(v) == ("Vasilhame Guaraná 300ML que volta (171 garrafas)", "7 CX + 3 UN")


def test_vasilhames_na_ordem_do_mapa():
    assert [v["nome"] for v in core.conta_vasilhames(ITENS_RET, _por_ref(CADASTRO_RET))] == [
        "Refri 2L", "KS", "Guaraná 300ML", "Cerveja 300ML"
    ]


def test_sem_retornavel_nao_tem_vasilhame():
    cadastro = {"1": "COCA COLA 2L FD 6X1", "2": "SPRITE 350ML FD 12X1"}  # 2L descartável não volta
    itens = [{"ref": "1", "nome": "", "qtd": 4.0}, {"ref": "2", "nome": "", "qtd": 3.0}]
    assert core.conta_vasilhames(itens, _por_ref(cadastro)) == []


def test_vasilhame_sem_tamanho_de_caixa_gera_aviso():
    itens = [{"ref": "99", "nome": "COCA COLA KS 290ML RET CX", "qtd": 2.0}]
    v = core.conta_vasilhames(itens, {})[0]
    assert v["sem_tamanho"] == ["Coca Cola KS 290ML RET"]
    assert v["garrafas"] == 0


def test_linhas_de_vasilhame_no_fim_de_retornaveis_e_fora_do_total(tmp_path):
    from openpyxl import load_workbook

    pasta = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ordem_mapa = core.le_ordem_mapa(os.path.join(pasta, "mapa_ordem.csv"))
    linhas, *_ = linha_de(ITENS_RET, CADASTRO_RET, ordem="mapa", ordem_mapa=ordem_mapa)
    vas = core.conta_vasilhames(ITENS_RET, _por_ref(CADASTRO_RET))
    _, _, final = core.escreve_xlsx(linhas, str(tmp_path / "Carga.xlsx"), "26/11/2026", "A", 33.0, vasilhames=vas)

    ws = load_workbook(final).active
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    i = col_a.index("Vasilhame Refri 2L que volta (42 garrafas)")
    # As três linhas ficam juntas, depois de todos os retornáveis e antes da próxima seção.
    assert col_a[i - 1] == "Guarana Antarctica 300ML RET"
    assert col_a[i:i + 4] == [
        "Vasilhame Refri 2L que volta (42 garrafas)",
        "Vasilhame KS que volta (245 garrafas)",
        "Vasilhame Guaraná 300ML que volta (171 garrafas)",
        "CERVEJAS RETORNÁVEIS",
    ]
    for r in range(i + 1, i + 4):
        assert ws.cell(row=r, column=4).value is None  # não somam no TOTAL
    # O TOTAL continua batendo com o sistema: 4+2+1+5+2+3+5+7+3+1 = 33.
    numeros = [ws.cell(row=r, column=c).value for r in range(2, ws.max_row) for c in (4, 5)]
    assert sum(n for n in numeros if isinstance(n, (int, float))) == 33


def test_vasilhame_de_cerveja_no_fim_de_cervejas_retornaveis(tmp_path):
    from openpyxl import load_workbook

    cadastro = dict(CADASTRO_RET, **{
        "50": "CERV SKOL 600ML RET CX 24X1",
        "51": "CERV BRAHMA CHOPP 600ML RET. CX 24X1",
        "52": "CERV SPATEN RET. 600ML CX 24X1",
        "53": "CERV SPATEN DESC 600ML CX 12X1",
        "60": "CERV SKOL LITRÃO 1L RET. CX 12X1",
        "61": "CERV BRAHMA LITRÃO 1L RET. UN",
        "70": "VODKA SMIRNOFF 998ML UN",
    })
    itens = ITENS_RET + [
        {"ref": "50", "nome": "", "qtd": 3.0},
        {"ref": "51", "nome": "", "qtd": 2.0},
        {"ref": "52", "nome": "", "qtd": 1.0},
        {"ref": "53", "nome": "", "qtd": 4.0},
        {"ref": "60", "nome": "", "qtd": 2.0},
        {"ref": "61", "nome": "", "qtd": 6.0},
        {"ref": "70", "nome": "", "qtd": 1.0},
    ]
    vas = core.conta_vasilhames(itens, _por_ref(cadastro))
    assert [v["nome"] for v in vas] == ["Refri 2L", "KS", "Guaraná 300ML", "Cerveja 300ML", "Cerveja 600ML", "Litrão 1L"]
    # 600ML: 3 + 2 + 1 engradados de 24 (a Spaten descartável não conta).
    assert core.texto_vasilhame(vas[3]) == ("Vasilhame Cerveja 300ML que volta (23 garrafas)", "23 UN")
    assert core.texto_vasilhame(vas[4]) == ("Vasilhame Cerveja 600ML que volta (144 garrafas)", "6 CX")
    # Litrão: 2 engradados de 12 + 6 garrafas soltas.
    assert core.texto_vasilhame(vas[5]) == ("Vasilhame Litrão 1L que volta (30 garrafas)", "2 CX + 6 UN")

    pasta = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ordem_mapa = core.le_ordem_mapa(os.path.join(pasta, "mapa_ordem.csv"))
    linhas, *_ = linha_de(itens, cadastro, ordem="mapa", ordem_mapa=ordem_mapa)
    _, _, final = core.escreve_xlsx(linhas, str(tmp_path / "Carga.xlsx"), "26/11/2026", "A", 52.0, vasilhames=vas)
    col_a = [c.value for c in load_workbook(final).active["A"]]

    # Refri 2L, KS e Guaraná continuam no fim de RETORNÁVEIS...
    i = col_a.index("CERVEJAS RETORNÁVEIS")
    assert col_a[i - 1] == "Vasilhame Guaraná 300ML que volta (171 garrafas)"
    # ...e as cervejas ficam no fim de CERVEJAS RETORNÁVEIS, antes de DESTILADOS.
    j = col_a.index("DESTILADOS")
    assert col_a[j - 3:j] == [
        "Vasilhame Cerveja 300ML que volta (23 garrafas)",
        "Vasilhame Cerveja 600ML que volta (144 garrafas)",
        "Vasilhame Litrão 1L que volta (30 garrafas)",
    ]


def test_heineken_e_amstel_litrao_em_linhas_proprias():
    cadastro = {
        "1": "CERV SKOL 600ML RET CX 24X1",
        "2": "CERV AMSTEL 600ML RET. CX 24X1",
        "3": "CERV HEINEKEN 600ML CX 24X1",
        "4": "CERV SKOL LITRÃO 1L RET. CX 12X1",
        "5": "CERV AMSTEL LITRÃO 1L RET. CX 12X1",
    }
    itens = [{"ref": r, "nome": "", "qtd": q} for r, q in (("1", 3.0), ("2", 2.0), ("3", 4.0), ("4", 1.0), ("5", 2.0))]
    vas = core.conta_vasilhames(itens, _por_ref(cadastro))
    assert [core.texto_vasilhame(v) for v in vas] == [
        ("Vasilhame Cerveja 600ML que volta (120 garrafas)", "5 CX"),  # Skol + Amstel
        ("Vasilhame Heineken 600ML que volta (96 garrafas)", "4 CX"),
        ("Vasilhame Litrão 1L que volta (12 garrafas)", "1 CX"),
        ("Vasilhame Amstel Litrão 1L que volta (24 garrafas)", "2 CX"),
    ]


def test_le_historico(tmp_path):
    import datetime

    pasta = str(tmp_path)
    assert core.le_historico(pasta) == []  # ainda não tem histórico
    dia1 = datetime.datetime(2026, 11, 25, 7, 42)
    dia2 = datetime.datetime(2026, 11, 26, 7, 35)
    core.registra_historico(pasta, "Carga A.xlsx", "Palmeiras", "25/11/2026", 48, 312.0, 312.0, True, agora=dia1)
    core.registra_historico(pasta, "Carga B.xlsx", "Grupo A", "26/11/2026", 55, 270.0, 271.0, False, agora=dia2)
    folhas = core.le_historico(pasta)
    assert [f["Mapa"] for f in folhas] == ["Palmeiras", "Grupo A"]  # ordem cronológica
    assert folhas[0]["Gerado em"] == dia1
    assert folhas[1]["Conferência"] == "NÃO BATEU"
    assert folhas[1]["caminho"] == os.path.join(pasta, "Carga B.xlsx")


def test_vasilhame_de_destilados():
    sim = [
        "CACHAÇA 51 965ML", "CACHACA 51 OURO 965ML", "CACHAÇA 51 SELEÇÃO 965ML",
        "CACHAÇA VELHO BARREIRO 910ML", "AGUARDENTE VILLA VELHA CARVALHO 965ML",
        "RECORD CANA E MEL 900ML", "RECORD CANELINHA 900ML", "RECORD FERNET 960ML", "RECORD COCO 960ML",
    ]
    nao = [
        "COROTE RECORD AMENDOIM 500ML", "VODKA RECORD LIMAO 500ML", "GROSELHA RECORD 500ML",
        "51 ICE BALADA 275ML", "VODKA SMIRNOFF 998ML",
    ]
    for nome in sim:
        assert core.tipo_vasilhame(nome) == "Destilados", nome
    for nome in nao:
        assert core.tipo_vasilhame(nome) is None, nome


def test_linha_de_destilados_no_fim_de_destilados(tmp_path):
    from openpyxl import load_workbook

    cadastro = {
        "1": "CACHAÇA 51 965ML CX 12X1",
        "2": "CACHAÇA 51 965ML UN",
        "3": "CACHAÇA VELHO BARREIRO 910ML CX 12X1",
        "4": "RECORD FERNET 960ML UN",
        "5": "VODKA SMIRNOFF 998ML UN",
        "6": "51 ICE BALADA 275ml FD 6X1",
    }
    itens = [{"ref": r, "nome": "", "qtd": q} for r, q in (("1", 2.0), ("2", 3.0), ("3", 1.0), ("4", 4.0), ("5", 1.0), ("6", 2.0))]
    vas = core.conta_vasilhames(itens, _por_ref(cadastro))
    # 2 + 1 caixas de 12 = 36 garrafas, mais 3 + 4 soltas = 43.
    assert [core.texto_vasilhame(v) for v in vas] == [
        ("Vasilhame Destilados que volta (43 garrafas)", "3 CX + 7 UN")
    ]
    pasta = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ordem_mapa = core.le_ordem_mapa(os.path.join(pasta, "mapa_ordem.csv"))
    linhas, *_ = linha_de(itens, cadastro, ordem="mapa", ordem_mapa=ordem_mapa)
    _, _, final = core.escreve_xlsx(linhas, str(tmp_path / "Carga.xlsx"), "26/11/2026", "A", 13.0, vasilhames=vas)
    col_a = [c.value for c in load_workbook(final).active["A"]]
    j = col_a.index("ICES")
    assert col_a[j - 1] == "Vasilhame Destilados que volta (43 garrafas)"


def test_destilados_casco_de_12_garrafas():
    cadastro = {"1": "CACHAÇA 51 965ML UN", "2": "RECORD FERNET 960ML CX 6X1"}
    # 12 garrafas soltas de 51 = 1 casco cheio; 3 caixas de 6 de Record = 18 garrafas.
    itens = [{"ref": "1", "nome": "", "qtd": 12.0}, {"ref": "2", "nome": "", "qtd": 3.0}]
    v = core.conta_vasilhames(itens, _por_ref(cadastro))[0]
    assert core.texto_vasilhame(v) == ("Vasilhame Destilados que volta (30 garrafas)", "2 CX + 6 UN")


def test_guarana_300ml_casco_de_24_garrafas():
    cadastro = {"1": "GUARANA ANTARCTICA 300ML RET UN", "2": "GUARANA ANTARCTICA 300ML RET CX 12X1"}
    # 30 garrafas soltas + 1 caixa de 12 = 42 garrafas = 1 casco de 24 + 18 soltas.
    itens = [{"ref": "1", "nome": "", "qtd": 30.0}, {"ref": "2", "nome": "", "qtd": 1.0}]
    v = core.conta_vasilhames(itens, _por_ref(cadastro))[0]
    assert core.texto_vasilhame(v) == ("Vasilhame Guaraná 300ML que volta (42 garrafas)", "1 CX + 18 UN")


def test_cerveja_300ml_casco_de_24_garrafas():
    cadastro = {
        "1": "CERV SKOL 300ML RET. CX 23X1",
        "2": "CERV ORIGINAL LITRINHO 300ML CX 23X1",
        "3": "CERV BRAHMA CHOPP 300ML RET. UN",
    }
    # 1 caixa de 23 + 1 caixa de 23 + 2 soltas = 48 garrafas = 2 cascos de 24.
    itens = [{"ref": "1", "nome": "", "qtd": 1.0}, {"ref": "2", "nome": "", "qtd": 1.0}, {"ref": "3", "nome": "", "qtd": 2.0}]
    v = core.conta_vasilhames(itens, _por_ref(cadastro))[0]
    assert core.texto_vasilhame(v) == ("Vasilhame Cerveja 300ML que volta (48 garrafas)", "2 CX")


def test_item_fora_do_mapa_nao_cria_secao_repetida(tmp_path):
    # Caso real: LATAS termina em 32.502 no mapa e 200ML começa em 33. Uma lata que não
    # está no mapa ia para 33.002, depois do primeiro item de 200ML, e a folha saía com
    # dois títulos "LATAS".
    from openpyxl import load_workbook

    pasta = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ordem_mapa = core.le_ordem_mapa(os.path.join(pasta, "mapa_ordem.csv"))
    cadastro = {
        "1": "COCA COLA 350ML FD 12X1",       # LATAS, no mapa
        "2": "FANTA MORANGO 350ML FD 12X1",   # LATAS, fora do mapa
        "3": "COCA COLA 200ML FD 12X1",       # 200ML, no mapa
        "4": "SPRITE 200ML FD 12X1",          # 200ML, no mapa
        "5": "GUARANA NOVO 200ML FD 12X1",    # 200ML, fora do mapa
        "6": "AGUA MINALICE S/ GÁS 510ML FD 12X1",  # ÁGUAS, no mapa
    }
    itens = [{"ref": r, "nome": "", "qtd": 1.0} for r in cadastro]
    linhas, _, _, fora = linha_de(itens, cadastro, ordem="mapa", ordem_mapa=ordem_mapa)
    assert sorted(fora) == ["Fanta Morango 350ML", "Guarana Novo 200ML"]

    # Cada seção aparece num bloco só, na ordem do mapa, com o item novo no fim dela.
    assert [(l[0], l[5]) for l in linhas] == [
        ("Coca Cola 350ML", "LATAS"),
        ("Fanta Morango 350ML", "LATAS"),
        ("Coca Cola 200ML", "200ML"),
        ("Sprite 200ML", "200ML"),
        ("Guarana Novo 200ML", "200ML"),
        ("Agua Minalice S/ Gás 510ML", "AGUAS"),
    ]
    _, _, final = core.escreve_xlsx(linhas, str(tmp_path / "Carga.xlsx"), "26/11/2026", "A", 6.0)
    col_a = [c.value for c in load_workbook(final).active["A"]]
    for titulo in ("LATAS", "200ML", "ÁGUAS"):
        assert col_a.count(titulo) == 1, titulo
