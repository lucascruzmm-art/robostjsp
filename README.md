# Robô de conferência de citação no SAJ

O robô percorre a árvore de documentos da janela **Visualização de Autos**. Lê cada AR e procura registros de citação em mandados, certidões e editais. Na planilha de entrada, grava um resumo por processo e as ocorrências com as páginas correspondentes.

## Instalação no Windows

1. Instale Python 3.11 ou 3.12 e o [Tesseract OCR](https://tesseract-ocr.github.io/tessdoc/Installation.html), com o idioma português (`por`). Confira com `tesseract --list-langs`. Se o comando não estiver no `PATH`, informe o caminho em `leitura.executavel_tesseract` no `config.json`, por exemplo `C:/Program Files/Tesseract-OCR/tesseract.exe`.
2. Extraia o ZIP. Abra o PowerShell na pasta que contém `config.json` e rode:

   ```powershell
   py -3.12 -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

   Para Python 3.11, use `py -3.11` no primeiro comando.

3. Em `config.json`, informe a `aba` se precisar selecionar uma aba específica, a `linha_cabecalho` e a `coluna_processo`. `coluna_resultado: null` usa a coluna à direita. Se ela já tiver cabeçalho, ajuste `titulo_resultado` para o texto usado na planilha. Os números CNJ devem estar em células de texto, com ou sem pontuação.

## Primeira execução

Abra o SAJ em um processo, maximize a janela **Visualização de Autos**, deixe a aba **Documentos** aberta e mantenha a opção **Exibir em ordem cronológica**. Para a leitura das folhas, escolha um zoom que mantenha a folha inteira visível com margens do painel ao redor. Enquanto o robô roda, mantenha o SAJ à frente e deixe o mouse e o teclado livres.

1. Confira a navegação com uma cópia da planilha, que pode conter processos de teste. Para o primeiro lote, inclua um AR positivo, um AR negativo e um processo com AR negativo seguido de retorno positivo do oficial ou de edital.

   ```powershell
   .\.venv\Scripts\python.exe -m robo_ar executar "C:\Dados\Consulta_teste.xlsx" --somente-testar --limite 3
   ```

2. Compare as páginas e os resultados do terminal com o que foi registrado nos autos. Se houver falha de leitura ou de navegação, ajuste a configuração conforme a seção de calibração.
3. Feche o arquivo no Excel e rode com gravação.

   ```powershell
   .\.venv\Scripts\python.exe -m robo_ar executar "C:\Dados\Consulta.xlsx"
   ```

O robô cria uma cópia de segurança na pasta da planilha antes da primeira alteração e grava após cada linha. Por padrão, a execução seguinte pula células de resultado preenchidas. Use `--reprocessar` para recalcular essas células, inclusive as que foram preenchidas manualmente. O comando aceita `.xlsx` e `.xlsm`. Valide a cópia de teste também se sua planilha usar controles, desenhos ou recursos avançados do Excel, pois a gravação usa `openpyxl`.

## Resultados da consulta

| Evidência encontrada | Resultado |
| --- | --- |
| Entrega registrada no AR | `Positivo` |
| Devolução com marca inequívoca | `Negativo`, acompanhado do motivo |
| Certidão ou retorno de oficial com relato afirmativo da citação | `Citação registrada — oficial de justiça` |
| Mandado expedido, com instrução de citar | `Expedido` |
| Edital com texto de citação | `Edital de citação localizado — publicação a verificar` |
| Documento registra a publicação daquele edital de citação | `Edital de citação — publicação registrada` |
| Certidão registra o decurso do prazo daquele edital de citação | `Edital de citação — decurso do prazo certificado` |
| OCR ambíguo, informação conflitante ou página cortada | `Revisar` |

O resultado resume os registros identificados no processo. Em processos com várias partes, a conferência do destinatário de cada citação é feita pelas páginas indicadas. Os editais seguem os registros de publicação ou de decurso encontrados nos autos. O robô deixa o cálculo de prazos para a análise do processo.

## Calibração do SAJ

O pacote traz regiões iniciais estimadas a partir da captura enviada. A instalação do SAJ, a escala do Windows e o zoom podem mudar o encaixe dessas regiões. A primeira validação no Windows é necessária para confirmar o carregamento correto de cada página.

- `inspecionar` lista os controles da janela e os respectivos identificadores. Deixe o SAJ aberto no processo que servirá de teste.
- `capturar janela.png` salva uma captura para medir as regiões da janela. A captura solicitada explicitamente pode conter dados do processo.
- `recorte janela.png` abre a captura e devolve as quatro coordenadas proporcionais da seleção com o mouse. Ajuste `caption_region` para incluir `Página` e `Categoria`, `capture_region` para o painel de AR e `document_region` para o painel a ser usado com as outras peças. Essas regiões são relativas à janela. Mantenha margem suficiente para o documento aparecer por inteiro.
- Se o campo de consulta não for reconhecido, configure `process_fields`. Cada entrada recebe um `selector` e um `format`. Use `{cnj}` para campo único, `{prefix}` e `{suffix}` para campos separados, ou `{digits}` para 20 dígitos. `submit_control` aceita o seletor da lupa quando pressionar Enter no campo não abrir o processo.
- Se houver um controle acessível de ajuste de página inteira, informe o seletor em `fit_page_control`. Se o visualizador expuser o texto integral da página via UI Automation, informe esse controle em `document_text_control` e defina `backend` como `uia`.

Exemplo de inspeção e teste local de leitura:

```powershell
.\.venv\Scripts\python.exe -m robo_ar inspecionar
.\.venv\Scripts\python.exe -m robo_ar capturar janela.png
.\.venv\Scripts\python.exe -m robo_ar recorte janela.png
.\.venv\Scripts\python.exe -m robo_ar imagem ar.png
.\.venv\Scripts\python.exe -m robo_ar imagem retorno.png --tipo CERTIDAO
```

Para `imagem`, forneça a página inteira. O reconhecimento roda localmente. Quando um modelo de AR exigir ajuste, `leitura.regiao` seleciona um retângulo dentro da imagem recebida. `leitura.caixas` e `leitura.campos` aceitam regiões dentro do formulário já recortado. Opcionalmente, `leitura.molde_branco` aponta para uma imagem do mesmo formulário em branco, no mesmo enquadramento. O caminho é relativo ao `config.json`. Confira cada motivo de devolução com controles positivos e negativos antes de usar um leiaute novo num lote.
