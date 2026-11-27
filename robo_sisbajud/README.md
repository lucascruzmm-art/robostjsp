# Robô de Juntada e Liberação Juntos — Sisbajud TJSP

Automação auxiliar para o SAJ de Primeiro Grau no Windows. Lê uma planilha Excel, associa os PDFs aos números de processo e executa o fluxo de juntada, tratamento do sigilo da Decisão e liberação de peças pela interface do SAJ.

## Baixar e usar

Baixe a distribuição pronta na [Release Sisbajud v1.0.0 portátil](https://github.com/lucascruzmm-art/robo_resultadoAR/releases/tag/sisbajud-v1.0.0-portatil). O código deste projeto fica na subpasta `robo_sisbajud/`; o robô de conferência de AR permanece independente.

1. Baixe `Robo_Sisbajud_Portatil_Windows_x64.zip` na Release.
2. Extraia **o ZIP inteiro** para uma pasta local gravável.
3. Abra `Robo_Sisbajud_Portatil` e dê dois cliques em `Iniciar.bat`.
4. Selecione a planilha `.xlsx` e a pasta dos PDFs. Confira as linhas antes de iniciar.
5. Deixe o SAJ aberto na tela inicial e mantenha a opção correspondente marcada.

O ZIP contém CPython 3.12.14 Windows x64, Tcl/Tk, bibliotecas e as 25 imagens de calibração recebidas. Não é necessário instalar Python, executar pip ou baixar componentes na primeira abertura. Mova sempre a pasta inteira. Não execute de dentro do ZIP.

**O checkout deste repositório não inclui a pasta `python/`: os iniciadores da distribuição pronta precisam do runtime que acompanha a Release.** O ZIP de 113.366.149 bytes é distribuído como asset de Release, não como arquivo Git.

## Planilha e PDFs

Use os cabeçalhos `numero do processo` e `status`, na primeira aba. Os números devem ser Texto, preservando os 20 dígitos. Status aceitos: `IRRISÓRIO`, `PARCIAL` e `INTEGRAL`. Um PDF por processo, com o número correspondente no nome, diretamente na pasta selecionada.

| Status | Código |
|---|---|
| IRRISÓRIO | 938 |
| PARCIAL | 939 |
| INTEGRAL | 939 |

O histórico é criado localmente em `historico_saj.csv`; tentativas anteriores não são repetidas automaticamente pelo lote. Não distribua seu histórico nem as planilhas/PDFs junto do programa.

## Fluxo atual

A fonte atual é a **versão 1, revisão 4**. Preserva o fluxo já construído: abrir o digitalizador, carregar processo, fechar pendências quando houver, importar o PDF, tratar o sigilo da Decisão, salvar/liberar e limpar para o próximo processo.

Depois de `Alt+L`, a espera inicial é de 6 segundos e o prazo total para disponibilizar Salvar é de 8 segundos. O vencimento é registrado e o lote tenta a recuperação por Limpar. A prioridade de Parar é preservada, inclusive quando o prazo vence durante a busca. Consulte [VERSAO_ATUAL.json](VERSAO_ATUAL.json) e os snapshots em [versoes/](versoes/).

As imagens atuais estão em `imagens_saj/` e são preservadas byte a byte. Existem 18 referências obrigatórias. Mudanças na escala do Windows, resolução ou aparência do SAJ podem exigir nova captura em Calibrar imagens. As referências são recortes de controles da interface, sem PDFs ou planilhas.

## Requisitos externos

- Windows 10/11 x64, sessão interativa desbloqueada e pasta gravável.
- SAJ instalado e funcional, acesso autorizado e rede habitual do tribunal.
- Níveis de permissão compatíveis entre robô e SAJ.
- Recursos nativos Win32/COM/UI Automation, clipboard, Universal CRT e Media Foundation; Windows N pode precisar do Media Feature Pack.
- Planilha e PDFs correspondentes.

Excel, leitor de PDF separado e Python instalado não são exigidos pelo robô portátil. Durante a automação, evite usar mouse/teclado. Para interromper, use Parar ou mova o mouse ao canto superior esquerdo.

## Validação e limites

O pacote foi extraído em outro caminho e executado com os binários Windows reais **sob Wine 10.0/Linux**. `Iniciar.bat` abriu e fechou normalmente sem Python/py do sistema no PATH. Também foram verificados imports locais, Tcl/Tk/ImageTk, screenshot somente em memória, OpenCV/NumPy/Pillow, XLSX sintético, Win32/COM/pywinauto e restauração do clipboard.

**Não houve teste em Windows nativo nem execução no SAJ nesta validação do pacote.** A versão exata do Python originalmente usado no computador do usuário não foi informada. Um ensaio com `%`, `&` e `!` no caminho falhou no CMD do Wine; o ensaio com espaços e acentos passou. Consulte os [relatórios de validação](docs/distribuicao/VALIDACAO_PORTATIL.md).

Na máquina de destino, `Verificar_portabilidade.bat` gera `logs/relatorio_portabilidade.json` sem clicar no SAJ. O diagnóstico verifica dependências e referências, não comprova uma juntada real.

## Desenvolvimento e testes

Para executar a fonte com Python Windows x64 instalado:

```bat
py -m pip install -r requirements.txt
py saj_previa.py
```

O teste de regressão usa relógio virtual e imagens sintéticas em memória, sem desktop ou ações no SAJ:

```bat
py tests/testar_limite_alt_l.py
```

O projeto não inclui pipeline CI. O builder `tools/empacotar_portatil.py` recebe uma árvore Python Windows **já provisionada**, confere as dependências fixadas e monta o ZIP sem baixar componentes. Exemplo:

```bat
py tools/empacotar_portatil.py --runtime C:\RuntimeRobo\python
```

A árvore informada deve corresponder ao runtime do [manifesto da distribuição](docs/distribuicao/MANIFESTO.json), incluindo Tcl/Tk, DLLs e bibliotecas. O builder não altera a fonte nem as calibrações. Este helper de reconstrução não foi usado para gerar o ZIP já validado; qualquer reconstrução exige executar novamente a verificação de portabilidade.

## Conteúdo do repositório

- `saj_previa.py`: código atual do robô.
- `imagens_saj/`: 25 PNGs recebidos; 18 exigidos pelo código.
- `Iniciar.bat` e `iniciar_portatil.py`: abertura isolada usando o Python da distribuição.
- `Verificar_portabilidade.bat` e `verificar_portabilidade.py`: diagnóstico local.
- `requirements.txt`: versões fixadas das bibliotecas.
- `versoes/`: cinco snapshots legítimos da versão 1 e revisões 1–4.
- `docs/distribuicao/`: manifestação, validação e SHA256 do ZIP publicado.
- `tests/`: regressão do prazo após Alt+L.

Não foram incluídos no Git o runtime Python, o ZIP binário, dados de processos, PDFs, planilhas, histórico, prefixos Wine ou arquivos temporários de depuração.

## Licenças de terceiros

As licenças do Python, Tcl/Tk e das bibliotecas acompanham a árvore do ZIP portátil. A origem do runtime e as versões constam no manifesto. Nenhuma licença adicional para o código próprio foi presumida nesta preparação.
