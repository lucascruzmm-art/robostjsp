# Validação da distribuição portátil

Arquivo: `Robo_Sisbajud_Portatil_Windows_x64.zip` — 113,366,149 bytes.

SHA256: `711c56d42e376354385e88d837a2b5351eaaed17ef699e9fc58374e1bebe26a0`.

O ZIP foi extraído em uma pasta diferente, com espaços e acentos, e os executáveis Windows foram executados sob Wine 10.0 em Linux. **Não houve teste em Windows nativo nem execução de processos no SAJ.**

| Verificação | Resultado |
|---|---|
| CRC do ZIP e 5050 hashes antes e depois dos testes | Aprovado |
| Iniciar.bat real: abriu a interface e fechou normalmente | Aprovado, saída 0 |
| Diretório de trabalho diferente da pasta do pacote | Aprovado |
| Python e py do sistema fora do PATH | Ambos não encontrados |
| PYTHONHOME/PYTHONPATH externos inválidos | Ignorados/limpos pelo iniciador |
| CPython 3.12.14 x64 e bibliotecas vindos do ZIP | Aprovado |
| Tcl/Tk 8.6.15, ImageTk, screenshot em memória e interface completa | Aprovado |
| NumPy 2.2.6, OpenCV 4.11.0.86, Pillow e XLSX sintético | Aprovado |
| Win32, COM e pywinauto Desktop | Aprovado |
| Clipboard: copiar/colar e restauração | Aprovado |
| 18 referências obrigatórias e 25 PNGs preservados | Aprovado |
| pip check | Nenhum requisito quebrado |
| Ações no SAJ | Nenhuma |

O código completo está em `saj_previa.py` dentro do ZIP. Mantém o fluxo atual e o prazo total de 8 segundos após Alt+L para disponibilizar Salvar. A prioridade de Parar também foi conferida em 9 testes simulados.

A máquina de teste exigiu UCRT e UIAutomationCore originais da Microsoft somente no prefixo Wine. Essas DLLs e o Wine **não** estão no ZIP. A aplicação e as bibliotecas não foram alteradas para contornar o Wine.

Um ensaio adicional em pasta com `%`, `&` e `!` falhou na interpretação do caminho pelo CMD do Wine. O ensaio aprovado usou espaços e acentos. Esse caso extremo não foi verificado em Windows nativo.

## Uso

1. Extraia o ZIP inteiro para uma pasta local gravável.
2. Abra `Robo_Sisbajud_Portatil` e dê dois cliques em `Iniciar.bat`.
3. Selecione a planilha e os PDFs na interface.

Não é necessário instalar Python, usar pip ou abrir PowerShell. O Python incluído é 3.12.14 x64; não havia registro da versão exata de Python da máquina original do usuário.

Para verificar a máquina de destino, abra `Verificar_portabilidade.bat`. O diagnóstico fica em `logs/relatorio_portabilidade.json`. Esse teste não opera o SAJ.

## Recursos externos necessários

- Windows 10 ou 11 x64 com desktop interativo, sessão desbloqueada e recursos nativos Win32/COM/UI Automation, clipboard e Universal CRT.
- Media Foundation do Windows; em edições N, Media Feature Pack pode ser necessário para carregar OpenCV.
- SAJ instalado e funcional, conta autorizada e acesso habitual à rede/sistemas do tribunal.
- Permissões compatíveis entre robô e SAJ, e uma pasta portátil gravável.
- A planilha .xlsx e os PDFs correspondentes.
- Mesma escala/aparência de tela das calibrações; alterações podem exigir nova captura pela interface.

Excel, leitor PDF separado e Python instalado não são exigidos pelo código. As licenças dos componentes acompanham a árvore do pacote.
