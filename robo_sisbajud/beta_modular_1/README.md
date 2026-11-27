# BETA modular 1 — Sisbajud TJSP

Beta separada da versão estável. Selecione as ações e clique em **Conferir planilha e ações** antes de iniciar. As três caixas começam marcadas e podem ser combinadas de sete maneiras; a ordem é Juntada → Sigilo → Liberação.

| Caixa | Comportamento |
|---|---|
| 1. Juntar PDF | Importa e salva o PDF, com código 938 para IRRISÓRIO e 939 para PARCIAL/INTEGRAL. |
| 2. Retirar sigilo da Decisão | Confere Peças sigilosas e remove somente o sigilo da Decisão encontrada. Não importa PDF nem libera peças por conta própria. |
| 3. Liberar peças | Seleciona e libera as peças aguardando liberação. Após Alt+L, aguarda 8 segundos pela gravação automática. |

Em todas as combinações, Limpar e a confirmação do campo vazio preparam o próximo processo. Uma falha interrompe as ações restantes daquela linha, registra as ações concluídas e recupera a tela antes de continuar o lote. Parar e a emergência pelo mouse permanecem.

## Como testar

- Use a fonte completa `saj_beta.py` junto de `imagens_saj/`, com as mesmas bibliotecas já instaladas para a versão estável.
- No pacote portátil existente, pode substituir **somente** o conteúdo de `saj_previa.py` pelo código desta beta e abrir pelo iniciador habitual. O ZIP antigo não foi reconstruído com esta beta.
- Para testar a fonte separadamente em Python Windows configurado: `py saj_beta.py`.
- Com a caixa 1 marcada, a planilha exige Processo e Status e a pasta deve conter exatamente um PDF correspondente por processo.
- Sem a caixa 1, basta a coluna Processo, como texto com 20 dígitos; status e PDFs não são exigidos.
- Ao mudar ações ou arquivos, confira a planilha novamente. Linhas excluídas da mesma planilha continuam excluídas após a reconferência.
- As pausas padrão continuam em 1,5 segundo; a espera após Alt+L permanece em 8 segundos.

## Histórico

A beta escreve em `historico_saj_beta.csv` e preserva o CSV anterior. Registra cada ação antes e depois da execução, permitindo revisar uma juntada concluída mesmo se o Sigilo falhar depois. Sigilo/Liberação não são bloqueados porque já houve uma juntada. O histórico anterior é lido como proteção conservadora contra repetir juntadas; em lote uma juntada já tentada é pulada, e a execução individual pede revisão antes de repetir.

## Verificação

52 testes de orquestração, planilha e interface aprovados, incluindo as sete combinações. Também foram aprovados quatro testes de GUI com CPython Windows 3.12.14 x64 sob Wine, em 850×540 e 1280×700. **Não houve execução das novas combinações no SAJ nem teste em Windows nativo.**

Testes: `python tests/testar_beta_planilha.py`, `python tests/testar_beta_interface.py` e `python tests/testar_beta_combinacoes.py`. Os dados de teste são artificiais, sem números de processos reais. A fonte estável permanece fora desta subpasta.
