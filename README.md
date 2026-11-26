# folha-de-carga

Projeto com o objetivo de automatizar a conferência e projeção de cargas para logística de distribuição de mercadorias (foco em distribuição de bebidas).

O programa lê o PDF **"Relatório de produtos - Pré-Venda"** do sistema e gera a folha de carga do caminhão em Excel, com duas colunas (**Produto** e **Quantidade**), separada pelas seções do mapa atual e pronta para imprimir frente e verso.

## Como usar (no computador da empresa)

1. Abra o `FolhaDeCarga.exe`.
2. Escolha o PDF do dia (ou arraste o PDF em cima do ícone do programa).
3. Na primeira vez, escolha o cadastro de produtos (`relatorio_Lista_de_produtos...xlsx`). O programa lembra dele depois.
4. Clique em **GERAR FOLHA**.

A planilha `Carga ... .xlsx` é salva na mesma pasta do PDF. Na tela aparece **CONFERÊNCIA OK** quando a soma da folha bate com o total do rodapé do sistema, e os avisos (produto fora do cadastro, embalagem não reconhecida, item fora do mapa).

### Histórico de cargas

Cada folha gerada ganha uma linha no **`Histórico de cargas.xlsx`**, na mesma pasta das folhas. As linhas ficam em ordem cronológica: a mais nova sempre vai no fim. Cada linha tem a hora em que a folha foi gerada, a data e o nome do mapa, o número de linhas, os volumes, o total do sistema e a conferência (OK, NÃO BATEU ou SEM TOTAL). Clicando no nome do arquivo, a folha abre.

No programa, a aba **Histórico** mostra essa lista: a folha que não bateu aparece em vermelho, e a folha cujo arquivo foi apagado ou movido aparece em cinza. Dois cliques numa linha abrem a folha. A aba usa a pasta da última folha gerada. Para ver outra pasta, clique em **Trocar pasta...**.

O histórico só ganha linhas, nunca é apagado. Se ele estiver aberto no Excel na hora de gerar, a folha sai normalmente e a tela avisa que aquela folha não entrou na lista. Feche o histórico e gere de novo.

## Como pegar o .exe mais novo

A cada mudança neste repositório, o GitHub gera o `.exe` sozinho:

1. Abra a aba **Actions** do repositório.
2. Clique na execução mais recente de **Gerar FolhaDeCarga.exe** (com o ✓ verde).
3. Em **Artifacts**, baixe **FolhaDeCarga** (vem dentro de um .zip).

Também dá para gerar na mão com o `criar_exe.bat` (precisa do Python instalado).

## Arquivos

| Arquivo | Para que serve |
| --- | --- |
| `gerar_folha_carga.py` | O "motor": lê o PDF e o cadastro, monta e salva a folha. Também roda pelo terminal. |
| `folha_carga_app.py` | A janela do programa. |
| `mapa_ordem.csv` | A ordem e as seções do MAPA_ATUAL. Vai dentro do .exe. |
| `apelidos.csv` | (opcional) Nomes curtos para os produtos. |
| `tests/` | Testes das regras, para uma mudança nova não quebrar o que já funciona. |
| `criar_exe.bat` | Gera o .exe no seu computador. |

### Mudar a ordem do mapa sem gerar outro .exe

Coloque um `mapa_ordem.csv` editado **na mesma pasta do .exe**. O programa usa o de fora antes do que vai embutido.

### Apelidos (nomes curtos)

Crie um `apelidos.csv` na mesma pasta do .exe:

```
base,nome_curto
CERV ANTARTICA PILSEN 350ML,Antarctica Pilsen 350ML
```

A coluna `base` é o nome do cadastro **sem** a embalagem (FD 12X1, CX 24X1, UN...).

## Regras da folha

- Uma linha por produto, nome sem a embalagem; embalagens diferentes do mesmo produto ficam na mesma linha. Ex.: `Sprite 350ML — 12(FD); 2(C/6)`.
- **(FD)** = fardo de 12 (nos produtos de 1L ou mais, fardo de 6). Fardo de outro tamanho sai como (C/6), (C/18) etc. Ex.: Sprite Zero 350ML só tem FD 6X1, então sai `(C/6)`, nunca `(FD)`.
- **(CX)** e **(UN)**: caixa e unidade. Se o mesmo produto tem mais de um tamanho de caixa, a de 12 (ou, sem ela, a maior) fica (CX) e as outras viram (C/n).
- **Vasilhames que devem voltar:** no fim da seção dos produtos aparece uma linha em destaque para cada vasilhame da carga, na ordem do mapa. Em RETORNÁVEIS:
  - **Refri 2L**: todo refrigerante 2L RET (Coca, Coca Zero, Fanta Laranja, Fanta Uva...). O casco leva **12 garrafas**: o programa soma as garrafas (FD 6X1 = 6) e divide por 12; o que sobra sai como UN. Ex.: 7 FD = 42 garrafas = `3 CX + 6 UN`.
  - **KS**: todo produto com "KS" no nome (Coca KS, Coca Zero KS, Fanta KS, Sprite KS...).
  - **Guaraná 300ML**: Guaraná 300ML RET (a cerveja litrinho 300ML tem linha própria). O casco leva **24 garrafas**: soma as garrafas e divide por 24.

  Em CERVEJAS RETORNÁVEIS:
  - **Cerveja 300ML**: os litrinhos retornáveis (Skol, Brahma Chopp, Antarctica, Original Litrinho). A Original Litrinho entra mesmo sem RET no nome.
  - **Cerveja 600ML**: toda cerveja 600ML com RET no nome (Skol, Brahma, Antarctica, Original, Amstel, Sub-Zero, Spaten RET...). A Amstel volta junto com as da Ambev. A Spaten DESC (descartável) não conta.
  - **Heineken 600ML**: casco próprio, linha separada (conta toda Heineken 600ML, menos a DESC).
  - **Litrão 1L**: toda cerveja 1L com RET no nome (Skol, Brahma, Antarctica, Original, Glacial).
  - **Amstel Litrão 1L**: casco próprio, linha separada.

  Em DESTILADOS (uma linha só):
  - **Destilados**: Cachaça 51 965ML (normal, Ouro e Seleção), Cachaça Velho Barreiro 910ML, Aguardente Villa Velha 965ML e os Record de 900/960ML (Cana e Mel, Canelinha, Menta, Fernet, Amendoim, Coco, Cacau). O Corote Record e a Vodka Record de 500ML não contam. O casco leva **12 garrafas**: soma as garrafas e divide por 12, como no Refri 2L.

  Ex.: `Vasilhame KS que volta (245 garrafas) — 10 CX + 5 UN`. Nos outros vasilhames, cada CX/FD conta como um engradado e as garrafas são calculadas pelo tamanho da caixa no cadastro (CX 24X1 = 24). Essas linhas não entram no TOTAL. Para incluir outro vasilhame, acrescente uma regra em `TIPOS_VASILHAME` no `gerar_folha_carga.py`.
- Cada produto tem um quadradinho ☐ na coluna **OK**, para quem carrega ir marcando o que já subiu no caminhão.

## Pelo terminal

```
pip install -r requirements.txt
python gerar_folha_carga.py MAPA_A_2026_11_10.pdf --cadastro relatorio_Lista_de_produtos.xlsx
python -m pytest
```
