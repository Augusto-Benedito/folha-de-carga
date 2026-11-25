# Instruções do projeto (para o Claude)

- Dono: Augusto, auxiliar administrativo numa distribuidora de bebidas. Fale com ele em português, de forma simples.
- O programa roda num computador Windows da empresa **sem Python**: o que chega lá é o `FolhaDeCarga.exe`, gerado pelo GitHub Actions (`.github/workflows/gerar-exe.yml`) a cada push na `main`.
- `gerar_folha_carga.py` é o núcleo (lógica + linha de comando); `folha_carga_app.py` é só a janela (tkinter) e chama `core.processar(...)`. Mantenha essa separação.
- O código foi recuperado do .exe original em nov/2026, conferido instrução por instrução contra o bytecode. Os comentários foram reescritos; a lógica é a original.
- Regras de embalagem e formato da folha estão no README ("Regras da folha"). Não mude essas regras sem o Augusto pedir.
- Antes de commitar: `python -m pytest`. Ao mudar uma regra, adicione/ajuste um teste em `tests/test_regras.py`.
- Não versionar PDFs do sistema nem folhas geradas (dados do dia a dia da empresa).
