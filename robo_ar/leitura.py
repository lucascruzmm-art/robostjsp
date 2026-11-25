"""Lê imagens de AR com OCR local e análise de formulários calibráveis."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from io import BytesIO
from pathlib import Path
import csv
import os
import re
import subprocess
import unicodedata

import numpy as np
from PIL import Image, ImageFilter, ImageOps


Rect = tuple[float, float, float, float]

MOTIVOS = {
    "Mudou-se": ("mudou se",),
    "Endereço insuficiente": ("endereco insuficiente", "endereco incompleto"),
    "Não existe o número": ("nao existe o numero", "numero inexistente"),
    "Desconhecido": ("desconhecido", "destinatario desconhecido"),
    "Recusado": ("recusado", "recusa", "recebimento recusado"),
    "Não procurado": ("nao procurado", "objeto nao procurado"),
    "Ausente": ("ausente", "destinatario ausente"),
    "Falecido": ("falecido", "destinatario falecido"),
    "Outros": ("outros",),
}

CAIXAS_DIGITAL: dict[str, Rect] = {
    "Mudou-se": (0.422, 0.483, 0.437, 0.515),
    "Endereço insuficiente": (0.422, 0.529, 0.437, 0.561),
    "Não existe o número": (0.422, 0.575, 0.437, 0.607),
    "Desconhecido": (0.422, 0.621, 0.437, 0.653),
    "Recusado": (0.600, 0.483, 0.615, 0.515),
    "Não procurado": (0.600, 0.529, 0.615, 0.561),
    "Ausente": (0.600, 0.575, 0.615, 0.607),
    "Falecido": (0.600, 0.621, 0.615, 0.653),
    "Outros": (0.422, 0.667, 0.437, 0.699),
}

CAMPOS_DIGITAL: dict[str, Rect] = {
    "assinatura": (0.095, 0.836, 0.560, 0.875),
    "data": (0.600, 0.836, 0.740, 0.875),
    "nome": (0.100, 0.910, 0.550, 0.943),
    "outros": (0.495, 0.660, 0.705, 0.696),
}

DECLARACAO_POSITIVA = re.compile(
    r"(?:objeto entregue(?: ao destinatario)?|entrega realizada com sucesso|"
    r"entregue ao destinatario|entrega efetuada|entrega concluida|citacao efetuada|"
    r"(?:resultado|situacao|status) da entrega (?:entregue|recebido|entrega efetuada))"
)
DECLARACAO_NEGATIVA = re.compile(
    r"(?:objeto nao entregue|entrega nao realizada|entrega nao efetuada|"
    r"devolvido ao remetente|objeto devolvido ao remetente|citacao nao efetuada|"
    r"(?:resultado|situacao|status) da entrega (?:nao entregue|devolvido))"
)
MOTIVO_EXPLICITO = re.compile(r"(?:motivo(?: da devolucao)?|ocorrencia)\s*[:—–-]\s*(.+)")
DATA = re.compile(r"(?<!\d)(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{4}|\d{2})(?!\d)")


@dataclass(frozen=True)
class ReceiptResult:
    status: str
    reason: str


@dataclass(frozen=True)
class _Linha:
    texto: str
    confianca: float


def _normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto.casefold())
    sem_acentos = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(re.findall(r"[a-z0-9]+", sem_acentos))


def _ler_retangulo(valores: list[float] | Rect, nome: str) -> Rect:
    if not isinstance(valores, (list, tuple)) or len(valores) != 4:
        raise ValueError(f"{nome} deve conter quatro coordenadas proporcionais.")
    x1, y1, x2, y2 = (float(valor) for valor in valores)
    if not (0 <= x1 < x2 <= 1 and 0 <= y1 < y2 <= 1):
        raise ValueError(f"{nome} deve usar coordenadas entre 0 e 1, em ordem crescente.")
    return x1, y1, x2, y2


def _recortar(imagem: Image.Image, regiao: Rect) -> Image.Image:
    largura, altura = imagem.size
    x1, y1, x2, y2 = (
        round(valor * escala)
        for valor, escala in zip(regiao, (largura, altura, largura, altura))
    )
    x1, y1 = min(x1, largura - 1), min(y1, altura - 1)
    return imagem.crop((x1, y1, max(x1 + 1, x2), max(y1 + 1, y2)))


def _intervalos(indicadores: np.ndarray, lacuna_maxima: int) -> list[tuple[int, int]]:
    bordas = np.flatnonzero(np.diff(indicadores.astype(np.int8), prepend=0, append=0))
    intervalos: list[tuple[int, int]] = []
    for inicio, fim in zip(bordas[::2], bordas[1::2]):
        if intervalos and inicio - intervalos[-1][1] <= lacuna_maxima:
            intervalos[-1] = (intervalos[-1][0], int(fim))
        else:
            intervalos.append((int(inicio), int(fim)))
    return intervalos


def _recortar_folha(imagem: Image.Image) -> Image.Image:
    """Localiza a folha clara do AR quando a imagem inclui o painel cinza do SAJ."""
    matriz = np.asarray(imagem)
    altura, largura = matriz.shape
    contorno = np.concatenate((matriz[0], matriz[-1], matriz[:, 0], matriz[:, -1]))
    if np.median(contorno) >= 247:
        return imagem

    brancos = matriz >= 247
    candidatas = []
    for y1, y2 in _intervalos(brancos.mean(axis=1) >= 0.22, max(2, round(altura * 0.01))):
        if y2 - y1 < altura * 0.23:
            continue
        for x1, x2 in _intervalos(
            brancos[y1:y2].mean(axis=0) >= 0.45, max(2, round(largura * 0.01))
        ):
            largura_folha, altura_folha = x2 - x1, y2 - y1
            if largura_folha >= largura * 0.26 and 1.65 <= largura_folha / altura_folha <= 2.75:
                candidatas.append((x1, y1, x2, y2))
    if not candidatas:
        return imagem
    maior = max(candidatas, key=lambda caixa: (caixa[2] - caixa[0]) * (caixa[3] - caixa[1]))
    return imagem.crop(maior)


class ReceiptReader:
    """Classifica ARs individuais e devolve o motivo ou a causa da revisão."""

    def __init__(self, config: dict, base_dir: Path):
        self.executavel = str(config.get("executavel_tesseract", "tesseract"))
        if not Path(self.executavel).is_absolute() and (
            "/" in self.executavel or "\\" in self.executavel
        ):
            self.executavel = str(base_dir / self.executavel)
        self.idioma = str(config.get("idioma", "por"))
        self.tempo_ocr = float(config.get("tempo_ocr", 20))
        self.confianca_ocr = float(config.get("confianca_ocr", 65))
        if self.tempo_ocr <= 0 or not 0 <= self.confianca_ocr <= 100:
            raise ValueError("Configure tempo_ocr positivo e confianca_ocr entre 0 e 100.")
        self.regiao = (
            _ler_retangulo(config["regiao"], "regiao") if config.get("regiao") is not None else None
        )
        self.auto_recortar = bool(config.get("auto_recortar", True))
        self.modelo_digital = bool(config.get("modelo_digital", True))
        caixas = config.get("caixas", CAIXAS_DIGITAL)
        if not isinstance(caixas, dict) or not caixas:
            raise ValueError("caixas deve indicar os motivos de devolução e suas coordenadas.")
        self.caixas = {
            motivo: _ler_retangulo(regiao, f"caixas.{motivo}") for motivo, regiao in caixas.items()
        }
        campos = config.get("campos", {})
        if not isinstance(campos, dict) or campos.keys() - CAMPOS_DIGITAL.keys():
            raise ValueError("campos pode configurar assinatura, data, nome e outros.")
        self.campos = {
            nome: _ler_retangulo(regiao, f"campos.{nome}")
            for nome, regiao in (CAMPOS_DIGITAL | campos).items()
        }
        self.molde = None
        if config.get("molde_branco"):
            caminho = base_dir / config["molde_branco"]
            with Image.open(caminho) as arquivo:
                self.molde = ImageOps.exif_transpose(arquivo).convert("L")

    def verify(self) -> None:
        idiomas = self._tesseract(["--list-langs"]).splitlines()
        for idioma in self.idioma.split("+"):
            if idioma not in idiomas:
                raise RuntimeError(f"O idioma {idioma} não está instalado no Tesseract.")

    def _tesseract(self, argumentos: list[str], dados: bytes | None = None) -> str:
        ambiente = os.environ.copy()
        ambiente.setdefault("OMP_THREAD_LIMIT", "1")
        try:
            resultado = subprocess.run(
                [self.executavel, *argumentos],
                input=dados,
                capture_output=True,
                timeout=self.tempo_ocr,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                env=ambiente,
            )
        except FileNotFoundError as erro:
            raise RuntimeError(
                "Tesseract não encontrado. Configure executavel_tesseract ou o PATH."
            ) from erro
        except subprocess.TimeoutExpired as erro:
            raise RuntimeError("O Tesseract excedeu o tempo de leitura do AR.") from erro
        if resultado.returncode != 0:
            detalhe = resultado.stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"O Tesseract não conseguiu ler a imagem. {detalhe}")
        return resultado.stdout.decode("utf-8", errors="replace")

    def _ocr(self, imagem: Image.Image, psm: str = "11") -> list[_Linha]:
        largura, altura = imagem.size
        escala = 4 if psm == "7" else min(3, max(1, 1400 / largura))
        interpolacao = Image.Resampling.BICUBIC if psm == "6" else Image.Resampling.LANCZOS
        ampliada = imagem.resize(
            (round(largura * escala), round(altura * escala)), interpolacao
        )
        with BytesIO() as buffer:
            ImageOps.autocontrast(ampliada, cutoff=1).save(buffer, format="PNG")
            tsv = self._tesseract(
                [
                    "stdin", "stdout", "-l", self.idioma, "--psm", psm,
                    "--dpi", "300", "-c", "tessedit_create_tsv=1",
                ],
                buffer.getvalue(),
            )

        linhas: dict[tuple[str, ...], list[tuple[str, float]]] = {}
        leitor = csv.DictReader(tsv.splitlines(), delimiter="\t", quoting=csv.QUOTE_NONE)
        if not {"text", "conf", "line_num"}.issubset(leitor.fieldnames or []):
            raise RuntimeError("O Tesseract retornou dados de OCR em formato inesperado.")
        for item in leitor:
            palavra = (item.get("text") or "").strip()
            if palavra:
                chave = tuple(item[campo] for campo in ("page_num", "block_num", "par_num", "line_num"))
                linhas.setdefault(chave, []).append((palavra, float(item["conf"])))
        return [
            _Linha(" ".join(palavra for palavra, _ in palavras), min(conf for _, conf in palavras))
            for palavras in linhas.values()
        ]

    def extract_text(self, image: Image.Image) -> str:
        """Extrai o texto por linha de uma página para análise documental."""
        imagem = ImageOps.exif_transpose(image).convert("L")
        if imagem.width < 60 or imagem.height < 30:
            return ""
        return "\n".join(linha.texto for linha in self._ocr(imagem, psm="6"))

    def _declaracoes(self, linhas: list[_Linha]) -> tuple[bool, bool, set[str]]:
        positiva = negativa = False
        motivos = set()
        for linha in linhas:
            if linha.confianca < self.confianca_ocr:
                continue
            texto = _normalizar(linha.texto)
            positiva |= bool(DECLARACAO_POSITIVA.fullmatch(texto))
            negativa |= bool(DECLARACAO_NEGATIVA.fullmatch(texto))
            for motivo, aliases in MOTIVOS.items():
                for alias in aliases:
                    if motivo == "Outros" or not texto.endswith(f" {alias}"):
                        continue
                    declaracao = texto[: -len(alias)].strip().removesuffix(" por").strip()
                    if DECLARACAO_NEGATIVA.fullmatch(declaracao):
                        motivos.add(motivo)
            sem_acentos = unicodedata.normalize("NFKD", linha.texto.casefold())
            sem_acentos = "".join(c for c in sem_acentos if not unicodedata.combining(c))
            detalhe = MOTIVO_EXPLICITO.fullmatch(sem_acentos)
            if detalhe:
                motivo_lido = _normalizar(detalhe.group(1))
                for motivo, aliases in MOTIVOS.items():
                    if motivo != "Outros" and motivo_lido in aliases:
                        motivos.add(motivo)
        return positiva, negativa, motivos

    def _layout_digital(self, imagem: Image.Image, linhas: list[_Linha]) -> bool:
        largura, altura = imagem.size
        if largura < 350 or altura < 150 or not 1.75 <= largura / altura <= 2.65:
            return False
        texto = _normalizar(" ".join(linha.texto for linha in linhas))
        if "motivos de devolucao" not in texto:
            return False
        campos = (
            "data da entrega", "data de entrega", "assinatura do recebedor",
            "tentativas de entrega", "carimbo", "nome legivel do",
        )
        return sum(campo in texto for campo in campos) >= 2

    def _marcacoes(self, tinta: np.ndarray, tamanho: tuple[int, int]) -> tuple[set[str], bool]:
        imagem = Image.fromarray(tinta.astype(np.uint8))
        densidades = {
            motivo: float(np.asarray(_recortar(imagem, caixa)).mean())
            for motivo, caixa in self.caixas.items()
        }
        if self.molde is not None:
            return {motivo for motivo, densidade in densidades.items() if densidade >= 0.19}, any(
                0.09 <= densidade < 0.19 for densidade in densidades.values()
            )

        largura, altura = tamanho
        if any(
            (x2 - x1) * largura < 4 or (y2 - y1) * altura < 4
            for x1, y1, x2, y2 in self.caixas.values()
        ):
            return set(), True
        marcados = set()
        ambiguidade = False
        for motivo, densidade in densidades.items():
            x1 = self.caixas[motivo][0]
            vizinhas = [
                outra_densidade
                for outro, outra_densidade in densidades.items()
                if outro != motivo and abs(self.caixas[outro][0] - x1) < 0.025
            ]
            fundo = float(np.median(vizinhas)) if vizinhas else 0
            if densidade >= 0.235 and densidade - fundo >= 0.14:
                marcados.add(motivo)
            elif densidade >= 0.22 and (
                densidade - fundo >= 0.08 or (densidade >= 0.28 and fundo >= 0.24)
            ):
                ambiguidade = True
        return marcados, ambiguidade

    @staticmethod
    def _tem_preenchimento(tinta: Image.Image, regiao: Rect, extensao_minima: float) -> bool:
        recorte = np.asarray(_recortar(tinta, regiao), dtype=bool).copy()
        recorte[recorte.mean(axis=1) >= 0.55] = False
        return bool(
            recorte.mean() >= 0.035 and (recorte.sum(axis=0) >= 2).mean() >= extensao_minima
        )

    def _data_legivel(self, imagem: Image.Image) -> bool:
        for linha in self._ocr(_recortar(imagem, self.campos["data"]), psm="7"):
            if linha.confianca < self.confianca_ocr:
                continue
            captura = DATA.search(linha.texto)
            if captura:
                dia, mes, ano = map(int, captura.groups())
                try:
                    date(ano + 2000 if ano < 100 else ano, mes, dia)
                except ValueError:
                    continue
                return True
        return False

    def _campos_de_entrega(self, imagem: Image.Image, tinta: np.ndarray) -> bool:
        campos = Image.fromarray(tinta.astype(np.uint8))
        if not self._tem_preenchimento(campos, self.campos["assinatura"], 0.14):
            return False
        if not self._tem_preenchimento(campos, self.campos["data"], 0.23):
            return False
        if self.molde is not None and self._tem_preenchimento(campos, self.campos["nome"], 0.14):
            return True
        return self._data_legivel(imagem)

    def _detalhar_outros(self, imagem: Image.Image) -> ReceiptResult:
        for linha in self._ocr(_recortar(imagem, self.campos["outros"]), psm="7"):
            if linha.confianca >= self.confianca_ocr and len(_normalizar(linha.texto)) >= 5:
                return ReceiptResult("Negativo", f"Outros — {linha.texto}")
        return ReceiptResult("Revisar", "Opção Outros marcada. Detalhe ilegível.")

    @staticmethod
    def _tinta_acrescentada(imagem: Image.Image, molde: Image.Image) -> np.ndarray | None:
        proporcao_imagem = imagem.width / imagem.height
        proporcao_molde = molde.width / molde.height
        if abs(proporcao_imagem / proporcao_molde - 1) > 0.08:
            return None
        redimensionado = molde.resize(imagem.size, Image.Resampling.BILINEAR)
        impresso = np.asarray(redimensionado) <= 120
        captura_dilatada = np.asarray(imagem.filter(ImageFilter.MinFilter(5))) <= 150
        if impresso.mean() < 0.01 or captura_dilatada[impresso].mean() < 0.62:
            return None
        margem_impresso = np.asarray(redimensionado.filter(ImageFilter.MinFilter(3)))
        return (np.asarray(imagem) <= 135) & (margem_impresso >= 175)

    def classify(self, image: Image.Image) -> ReceiptResult:
        imagem = ImageOps.exif_transpose(image).convert("L")
        if self.regiao is not None:
            imagem = _recortar(imagem, self.regiao)
        elif self.auto_recortar:
            imagem = _recortar_folha(imagem)
        if imagem.width < 60 or imagem.height < 30:
            return ReceiptResult("Revisar", "Imagem pequena demais para leitura.")

        linhas = self._ocr(imagem)
        texto_positivo, texto_negativo, motivos = self._declaracoes(linhas)
        contem_marca_incerta = False
        campos_preenchidos = False
        if self.modelo_digital and self._layout_digital(imagem, linhas):
            if self.molde is not None:
                tinta = self._tinta_acrescentada(imagem, self.molde)
                if tinta is None:
                    return ReceiptResult("Revisar", "O AR não se alinhou ao molde em branco.")
            else:
                tinta = np.asarray(imagem) <= 135
            marcados, contem_marca_incerta = self._marcacoes(tinta, imagem.size)
            motivos.update(marcados)
            campos_preenchidos = self._campos_de_entrega(imagem, tinta)

        tem_entrega = texto_positivo or campos_preenchidos
        if tem_entrega and (motivos or texto_negativo):
            return ReceiptResult("Revisar", "Indícios de entrega e de devolução no mesmo AR.")
        if len(motivos) > 1:
            return ReceiptResult("Revisar", "Mais de um possível motivo de devolução.")
        if contem_marca_incerta:
            return ReceiptResult("Revisar", "Marcação no formulário ambígua.")
        if motivos:
            motivo = next(iter(motivos))
            if motivo == "Outros":
                return self._detalhar_outros(imagem)
            return ReceiptResult("Negativo", motivo)
        if texto_negativo:
            return ReceiptResult("Revisar", "Devolução indicada com motivo não identificado.")
        if tem_entrega:
            return ReceiptResult("Positivo", "Entrega registrada no AR.")
        return ReceiptResult("Revisar", "A imagem não permite confirmar a entrega ou a devolução.")
