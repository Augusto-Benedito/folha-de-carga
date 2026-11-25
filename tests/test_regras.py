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
