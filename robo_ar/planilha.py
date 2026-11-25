from contextlib import contextmanager
from copy import copy
from pathlib import Path
import os
import re
import shutil
import tempfile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import column_index_from_string


class PlanilhaErro(Exception):
    pass


def numero_cnj(value):
    if not isinstance(value, str):
        raise PlanilhaErro("número do processo deve estar em uma célula de texto")
    value = value.strip()
    if re.fullmatch(r"[0-9]{7}-[0-9]{2}\.[0-9]{4}\.[0-9]\.[0-9]{2}\.[0-9]{4}", value):
        return value
    if re.fullmatch(r"[0-9]{20}", value):
        return f"{value[:7]}-{value[7:9]}.{value[9:13]}.{value[13]}.{value[14:16]}.{value[16:]}"
    raise PlanilhaErro("número do processo fora do formato CNJ")


def _coluna(value):
    try:
        coluna = column_index_from_string(value.strip())
        if coluna > 16384:
            raise ValueError
        return coluna
    except (AttributeError, ValueError) as exc:
        raise PlanilhaErro("use a letra da coluna do Excel, por exemplo A ou AB") from exc


def _vazia(value):
    return value is None or isinstance(value, str) and not value.strip()


def _estado(path):
    stat = path.stat()
    return stat.st_mtime_ns, stat.st_size, stat.st_ino


@contextmanager
def abrir_planilha(path, config, somente_testar):
    path = Path(path).resolve(strict=True)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise PlanilhaErro("use um arquivo .xlsx ou .xlsm")

    lock = path.with_name(f".{path.name}.robo-ar.lock")
    if not somente_testar:
        try:
            with lock.open("x", encoding="utf-8") as handle:
                handle.write(str(os.getpid()))
        except FileExistsError as exc:
            raise PlanilhaErro(f"planilha em uso pelo robô (arquivo de trava {lock.name})") from exc
    workbook = None
    try:
        estado = _estado(path)
        workbook = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
        if _estado(path) != estado:
            raise PlanilhaErro("a planilha foi alterada enquanto era aberta")
        yield Planilha(path, workbook, estado, config)
    finally:
        if workbook is not None:
            workbook.close()
            if workbook.vba_archive is not None:
                workbook.vba_archive.close()
        if not somente_testar:
            lock.unlink(missing_ok=True)


class Planilha:
    def __init__(self, path, workbook, estado, config):
        self.path = path
        self.workbook = workbook
        self.estado = estado
        self.backup = None
        aba = config.get("aba")
        if aba is not None and aba not in workbook.sheetnames:
            raise PlanilhaErro(f"aba não encontrada: {aba!r}")
        self.aba = workbook[aba] if aba is not None else workbook.active
        self.cabecalho = config.get("linha_cabecalho", 1)
        if not isinstance(self.cabecalho, int) or isinstance(self.cabecalho, bool) or self.cabecalho < 1:
            raise PlanilhaErro("linha_cabecalho deve ser um inteiro a partir de 1")
        self.coluna_processo = _coluna(config["coluna_processo"])
        coluna_resultado = config.get("coluna_resultado")
        self.coluna_resultado = (
            _coluna(coluna_resultado) if coluna_resultado else self.coluna_processo + 1
        )
        if self.coluna_resultado == self.coluna_processo or self.coluna_resultado > 16384:
            raise PlanilhaErro("configure uma coluna de resultado válida e diferente da coluna de processo")
        self.titulo = config.get("titulo_resultado", "Resultado AR")
        if not isinstance(self.titulo, str) or not self.titulo.strip():
            raise PlanilhaErro("titulo_resultado deve ser um texto não vazio")
        celula = self.aba.cell(self.cabecalho, self.coluna_resultado)
        if isinstance(celula, MergedCell) or celula.data_type == "f":
            raise PlanilhaErro(f"o cabeçalho de resultado {celula.coordinate} precisa ser uma célula de texto")
        if _vazia(celula.value):
            for (value,) in self.aba.iter_rows(
                min_row=self.cabecalho + 1,
                min_col=self.coluna_resultado,
                max_col=self.coluna_resultado,
                values_only=True,
            ):
                if not _vazia(value):
                    raise PlanilhaErro(f"a coluna de resultado tem dados sem o cabeçalho {self.titulo!r}")
        elif str(celula.value).strip().casefold() != self.titulo.strip().casefold():
            raise PlanilhaErro(f"o cabeçalho em {celula.coordinate} é {celula.value!r}. Configure titulo_resultado ou coluna_resultado")

    def pendentes(self, reprocessar, limite):
        selecionadas = 0
        for linha in range(self.cabecalho + 1, self.aba.max_row + 1):
            entrada = self.aba.cell(linha, self.coluna_processo)
            if _vazia(entrada.value):
                continue
            saida = self.aba.cell(linha, self.coluna_resultado)
            if isinstance(saida, MergedCell) or saida.data_type == "f":
                raise PlanilhaErro(f"a célula de resultado {saida.coordinate} está mesclada ou contém fórmula")
            if not reprocessar and not _vazia(saida.value):
                continue
            yield linha, entrada.value
            selecionadas += 1
            if limite is not None and selecionadas >= limite:
                break

    def gravar(self, linha, resultado):
        if len(resultado.encode("utf-16-le")) // 2 > 32767:
            raise PlanilhaErro(f"resultados da linha {linha} excedem o limite de uma célula do Excel")
        self._checar_estado()
        if self.backup is None:
            descriptor, name = tempfile.mkstemp(
                prefix=f"{self.path.stem}.backup-robo-ar-", suffix=self.path.suffix, dir=self.path.parent
            )
            os.close(descriptor)
            try:
                shutil.copy2(self.path, name)
                self._checar_estado()
            except Exception:
                Path(name).unlink(missing_ok=True)
                raise
            self.backup = Path(name)

        cabecalho = self.aba.cell(self.cabecalho, self.coluna_resultado)
        if _vazia(cabecalho.value):
            cabecalho.value = self.titulo
            cabecalho.data_type = "s"
        celula = self.aba.cell(linha, self.coluna_resultado)
        celula.value = resultado
        celula.data_type = "s"
        if not celula.alignment.wrap_text:
            alinhamento = copy(celula.alignment)
            alinhamento.wrap_text = True
            celula.alignment = alinhamento

        descriptor, name = tempfile.mkstemp(
            prefix=f".{self.path.stem}-", suffix=self.path.suffix, dir=self.path.parent
        )
        os.close(descriptor)
        temporario = Path(name)
        try:
            self.workbook.save(temporario)
            shutil.copymode(self.path, temporario)
            self._checar_estado()
            temporario.replace(self.path)
            self.estado = _estado(self.path)
        finally:
            temporario.unlink(missing_ok=True)

    def _checar_estado(self):
        if _estado(self.path) != self.estado:
            raise PlanilhaErro("a planilha foi alterada durante o processamento. Feche o Excel e reinicie o robô")
