from argparse import ArgumentParser, ArgumentTypeError
from collections import Counter, defaultdict
from pathlib import Path
import json
import sys

from PIL import Image
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from .atos import ResultadoAto, classificar_ato
from .leitura import ReceiptReader
from .planilha import PlanilhaErro, abrir_planilha, numero_cnj
from .saj import SajClient, SajError


def _limite(value):
    try:
        value = int(value)
    except ValueError as exc:
        raise ArgumentTypeError("use um número inteiro a partir de 1") from exc
    if value < 1:
        raise ArgumentTypeError("use um número inteiro a partir de 1")
    return value


def _argumentos():
    parser = ArgumentParser(description="Consulta de avisos de recebimento no SAJ")
    parser.add_argument("--config", type=Path, default=Path("config.json"))
    sub = parser.add_subparsers(dest="comando", required=True)
    sub.add_parser("inspecionar", help="lista os controles da janela do SAJ")
    captura = sub.add_parser("capturar", help="salva a janela do SAJ como PNG para calibração")
    captura.add_argument("arquivo", type=Path)
    imagem = sub.add_parser("imagem", help="classifica uma página completa a partir de uma imagem local")
    imagem.add_argument("arquivo", type=Path)
    imagem.add_argument("--tipo", choices=("AR", "MANDADO", "CERTIDAO", "EDITAL"), default="AR")
    recorte = sub.add_parser("recorte", help="mostra uma imagem para selecionar uma região com o mouse")
    recorte.add_argument("arquivo", type=Path)
    execucao = sub.add_parser("executar", help="lê os processos da planilha e consulta o SAJ")
    execucao.add_argument("planilha", type=Path)
    execucao.add_argument("--limite", type=_limite, help="quantidade máxima de linhas a processar")
    execucao.add_argument("--somente-testar", action="store_true", help="consulta sem editar a planilha")
    execucao.add_argument("--reprocessar", action="store_true", help="inclui linhas com resultado já preenchido")
    return parser.parse_args()


_ROTULOS = {"AR": "AR", "MANDADO": "Mandado", "CERTIDAO": "Certidão", "EDITAL": "Edital"}
_EVIDENCIAS = {"ar": "AR", "oficial": "oficial de justiça"}


def _resultado_documento(tipo, paginas):
    referencias = ", ".join(dict.fromkeys(pagina for pagina, _, _ in paginas))
    if tipo == "AR":
        resultados_unicos = {resultado for _, resultado, _ in paginas}
        if len(resultados_unicos) == 1:
            return referencias, paginas[0][1]
        return referencias, ResultadoAto(
            "Revisar", "resultados diferentes entre as páginas\n" + "\n".join(
                f"  {pagina}: {resultado.status} — {resultado.motivo}" for pagina, resultado, _ in paginas
            )
        )

    falhas = [f"{pagina} ({resultado.motivo})" for pagina, resultado, _ in paginas if resultado is not None]
    resultado = classificar_ato(tipo, "\n".join(texto for _, _, texto in paginas))
    if falhas:
        motivo = "leitura incompleta em " + ", ".join(falhas)
        if resultado.evidencia:
            motivo += f". Trecho disponível indica {resultado.motivo}"
        return referencias, ResultadoAto("Revisar", motivo)
    return referencias, resultado


def _conclusao(evidencias, pendencias):
    confirmacoes = [rotulo for valor, rotulo in _EVIDENCIAS.items() if valor in evidencias]
    edital = ""
    if "edital_prazo" in evidencias:
        edital = "Edital de citação — decurso do prazo certificado"
    elif "edital_publicado" in evidencias:
        edital = "Edital de citação — publicação registrada"
    elif "edital_localizado" in evidencias:
        edital = "Edital de citação localizado — publicação a verificar"
    if confirmacoes:
        resumo = "Citação registrada — " + " e ".join(confirmacoes)
        if edital:
            resumo += f"\n{edital}"
    else:
        resumo = edital or ("Citação a verificar" if pendencias else "Citação não localizada nos documentos consultados")
    if pendencias and (confirmacoes or edital):
        resumo += "\nHá documentos a revisar."
    return resumo


def _ler_pagina(captura, leitor):
    texto = captura.text or ""
    try:
        if captura.kind == "AR" and captura.image is not None and captura.complete:
            leitura = leitor.classify(captura.image)
            evidencia = "ar" if leitura.status == "Positivo" else ""
            return ResultadoAto(leitura.status, leitura.reason, evidencia), ""
        if captura.kind != "AR" and captura.image is not None and not texto:
            texto = leitor.extract_text(captura.image)
    except (OSError, RuntimeError) as exc:
        return ResultadoAto("Revisar", f"falha no OCR ({exc})"), ""

    if not captura.complete or captura.error or captura.kind == "AR" or (captura.image is None and not texto):
        return ResultadoAto("Revisar", captura.error or "página indisponível no SAJ"), texto
    return None, texto


def _consulta(cliente, leitor, cnj):
    documentos = defaultdict(list)
    detalhe = None
    try:
        cliente.open_case(cnj)
        for captura in cliente.iter_documents():
            resultado, texto = _ler_pagina(captura, leitor)
            pagina = f"pág. {captura.page}" if captura.page else "página não identificada"
            documentos[(captura.document_index, captura.kind)].append((pagina, resultado, texto))
    except SajError as exc:
        detalhe = str(exc) or "falha de navegação"

    linhas = []
    contagem = Counter()
    evidencias = set()
    pendencias = detalhe is not None
    for (_, tipo), paginas in documentos.items():
        referencias, resultado = _resultado_documento(tipo, paginas)
        if resultado.status == "Ignorar":
            continue
        contagem[tipo] += 1
        rotulo = _ROTULOS[tipo]
        linhas.append(f"{rotulo} {contagem[tipo]} ({referencias}): {resultado.status} — {resultado.motivo}")
        if resultado.evidencia:
            evidencias.add(resultado.evidencia)
        pendencias = pendencias or resultado.status == "Revisar"
    if not contagem["AR"] and detalhe is None:
        linhas.append("Sem AR")
    if detalhe is not None:
        linhas.append(f"Revisar — consulta incompleta no SAJ ({detalhe})")
    return "\n".join([_conclusao(evidencias, pendencias), *linhas])


def _executar(args, config, leitor):
    with abrir_planilha(args.planilha, config["planilha"], args.somente_testar) as planilha:
        cliente = None
        total = 0
        for linha, entrada in planilha.pendentes(args.reprocessar, args.limite):
            try:
                cnj = numero_cnj(entrada)
            except PlanilhaErro as exc:
                resultado = f"Revisar — {exc}"
            else:
                if cliente is None:
                    leitor.verify()
                    cliente = SajClient(config["saj"])
                    cliente.ready()
                resultado = _consulta(cliente, leitor, cnj)

            resultado = ILLEGAL_CHARACTERS_RE.sub("", resultado)
            if not args.somente_testar:
                planilha.gravar(linha, resultado)
                if total == 0:
                    print(f"Cópia de segurança: {planilha.backup}", flush=True)
            print(f"Linha {linha}: {resultado}", flush=True)
            total += 1
        print(f"Linhas processadas: {total}")


def main():
    args = _argumentos()
    try:
        if args.comando == "recorte":
            from .recorte import selecionar_regiao

            selecionar_regiao(args.arquivo)
            return 0

        with args.config.open(encoding="utf-8-sig") as arquivo:
            config = json.load(arquivo)
        base_dir = args.config.resolve().parent
        executavel = config["leitura"].get("executavel_tesseract", "tesseract")
        if "/" in executavel or "\\" in executavel:
            executavel = str(base_dir / executavel)
        config["saj"].setdefault("tesseract_cmd", executavel)
        config["saj"].setdefault("tesseract_lang", config["leitura"].get("idioma", "por"))
        if args.comando == "inspecionar":
            print(SajClient(config["saj"]).inspect())
        elif args.comando == "capturar":
            captura = SajClient(config["saj"]).snapshot_window()
            with args.arquivo.open("xb") as arquivo:
                captura.save(arquivo, format="PNG")
        else:
            leitor = ReceiptReader(config["leitura"], base_dir)
            if args.comando == "imagem":
                leitor.verify()
                with Image.open(args.arquivo) as imagem:
                    if args.tipo == "AR":
                        resultado = leitor.classify(imagem)
                        print(f"{resultado.status} — {resultado.reason}")
                    else:
                        resultado = classificar_ato(args.tipo, leitor.extract_text(imagem))
                        print(f"{resultado.status} — {resultado.motivo}")
            else:
                _executar(args, config, leitor)
    except KeyboardInterrupt:
        print("Execução interrompida.", file=sys.stderr)
        return 130
    except (SajError, PlanilhaErro, OSError, KeyError, ValueError, TypeError, RuntimeError) as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
