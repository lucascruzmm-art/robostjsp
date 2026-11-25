from dataclasses import dataclass
import re
import unicodedata


@dataclass(frozen=True)
class ResultadoAto:
    status: str
    motivo: str
    evidencia: str = ""


def _normalizar(texto):
    texto = unicodedata.normalize("NFKD", texto.casefold())
    texto = "".join(letra for letra in texto if not unicodedata.combining(letra))
    texto = re.sub(r"(?<=\w)-[ \t]*\n(?=\w)", "", texto)
    return re.sub(r"\s+", " ", texto).strip()


_CITACAO = re.compile(r"\b(?:citacao|citar|citei|citad[oa]s?)\b")
_EDITAL_CITACAO = re.compile(r"\bedital\s+(?:de\s+|para\s+(?:a\s+)?)citacao\b|\bcitacao\s+por\s+edital\b")
_CONTEXTO_OFICIAL = re.compile(r"\boficial\s+d[ea]\s+justica\b|\bem\s+cumprimento\s+a[oa]\s+mandado\b|\bmandado\s+(?:foi\s+)?cumprido\b")
_RELATO = re.compile(r"\b(?:certifico|dou\s+fe|cumprimento\s+(?:a[oa]|d[eo])\s+mandado|diligenciei|dirigi-me)\b")
_CITEI = re.compile(
    r"\b(?:citei|dei\s+por\s+citad[oa]s?|dei\s+por\s+feita\s+a\s+citacao|procedi\s+a\s+citacao|(?:efetuei|efetivei)\s+a\s+citacao|"
    r"realizei\s+a\s+citacao|(?:foi|foram)\s+(?:devidamente\s+|pessoalmente\s+)?citad[oa]s?|"
    r"citacao\s+(?:foi\s+)?(?:efetuada|realizada|cumprida))\b"
)
_NEGATIVO = re.compile(
    r"\b(?:deixei\s+de\s+citar|nao\s+(?:(?:o|a|os|as|lhe|lhes)\s+)?"
    r"(?:citei|consegui\s+citar|(?:procedi|efetuei|efetivei|realizei)\s+a\s+citacao)|"
    r"nao\s+(?:(?:foi|foram)\s+)?(?:possivel\s+(?:realizar|efetuar)\s+a\s+citacao|"
    r"(?:(?:devidamente|pessoalmente)\s+)?citad[oa]s?)|"
    r"citacao\s+(?:restou\s+|foi\s+)?infrutifera|citacao\s+nao\s+foi\s+(?:efetuada|realizada|cumprida))\b"
)
_MENCAO_INCERTA = re.compile(
    r"\b(?:nao\s+consta|sem\s+que|caso|(?<!-)se|verificar\s+se|certificar\s+se|"
    r"(?!nao\s+so\b)nao|segundo|(?:informou|alegou)\s+que)\b[^.;:]{0,65}$"
)
_OUTRO_PROCESSO = re.compile(r"\b(?:em\s+outro\s+processo|em\s+autos\s+diversos|em\s+outros\s+autos)\b")
_PREPARATORIO = re.compile(
    r"\b(?:manda\s+(?:a[oa]|que)|proceda\s+a\s+citacao|devera\s+citar|"
    r"(?:expe[cç]a-se|expedido)\s+(?:o\s+)?mandado|ordem\s+de\s+citacao|cite-se)\b"
)
_EDITAL_APTO = re.compile(
    r"\b(?:fica[m]?\s+citad[oa]s?|faz\s+saber|que\s+pelo\s+presente\s+edital\s+cita|"
    r"pelo\s+presente\s+edital\s+(?:fica|ficam)\s+citad[oa]s?)\b"
)
_EDITAL_INDEFINIDO = re.compile(
    r"\b(?:minuta\s+(?:de|do)\s+edital|publique-se|publicar-se-a|sera\s+publicado|"
    r"expedicao\s+(?:de|do)\s+edital|expeca-se\s+(?:o\s+)?edital)\b"
)
_PUBLICADO = re.compile(
    r"\b(?:edital\s+de\s+citacao\s+(?:foi\s+|restou\s+)?publicado|"
    r"(?:publiquei|publicou-se|foi\s+publicado)\s+(?:no\s+(?:diario\s+da\s+justica(?:\s+(?:eletronico|nacional))?|djen?)\s+)?"
    r"o\s+edital\s+de\s+citacao|"
    r"(?:houve\s+(?:a\s+)?publicacao\s+do\s+edital\s+de\s+citacao|"
    r"publicacao\s+do\s+edital\s+de\s+citacao\s+(?:ocorreu|(?:foi\s+)?realizada))\s+"
    r"em\s+\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})\b"
)
_CERT_PUBLICACAO = re.compile(r"\b(?:certidao\s+de\s+publicacao|certifico|dou\s+fe|comprovante\s+de\s+publicacao)\b")
_DECURSO = re.compile(
    r"\b(?:decorreu|transcorreu|decorrido)\s+(?:integralmente\s+)?o\s+prazo\s+"
    r"(?:do\s+edital\s+de\s+citacao|da\s+citacao\s+por\s+edital)\b"
)
_INDEFERIDO = re.compile(
    r"\b(?:citacao|edital)\s+(?:foi\s+)?(?:anulad[oa]|cancelad[oa]|invalidada)|"
    r"\bnao\s+(?:foi\s+)?publicad[oa]\b|\bnulidade\s+da\s+citacao\b"
)
_MOTIVOS_DILIGENCIA = (
    (re.compile(r"\bmudou-se\b|\bmudou\s+d[eo]\s+endereco\b"), "mudança de endereço"),
    (re.compile(r"\bendereco\s+(?:insuficiente|incompleto|incorreto)\b"), "endereço insuficiente ou incorreto"),
    (re.compile(r"\b(?:nao\s+localizei|nao\s+(?:foi\s+)?localizad[oa])\b"), "pessoa ou endereço não localizado"),
    (re.compile(r"\b(?:faleceu|falecid[oa])\b"), "falecimento mencionado"),
    (re.compile(r"\b(?:estava|estar|encontrava-se)\s+ausente\b"), "ausência no endereço"),
)


def classificar_ato(categoria, texto):
    texto = _normalizar(texto)
    if sum(letra.isalpha() for letra in texto) < 45:
        return ResultadoAto("Revisar", "texto insuficiente para interpretar o documento")

    if _INDEFERIDO.search(texto):
        return ResultadoAto("Revisar", "há menção a anulação, cancelamento ou falta de publicação")

    edital_citacao = _EDITAL_CITACAO.search(texto)
    if edital_citacao or categoria == "EDITAL":
        if edital_citacao and _OUTRO_PROCESSO.search(texto):
            return ResultadoAto("Revisar", "o relato menciona edital em outro processo")
        if edital_citacao and re.search(r"\bminuta\s+(?:de|do)\s+edital\b", texto):
            return ResultadoAto("Preparatório", "minuta de edital de citação localizada")
        if edital_citacao and _CERT_PUBLICACAO.search(texto) and _DECURSO.search(texto):
            return ResultadoAto("Prazo certificado", "certidão registra o decurso do prazo do edital de citação", "edital_prazo")
        if edital_citacao and _CERT_PUBLICACAO.search(texto) and _PUBLICADO.search(texto):
            return ResultadoAto("Publicado", "publicação do edital de citação registrada", "edital_publicado")
        if categoria == "EDITAL" and edital_citacao and _EDITAL_APTO.search(texto):
            return ResultadoAto("Localizado", "edital de citação localizado", "edital_localizado")
        if edital_citacao and _EDITAL_INDEFINIDO.search(texto):
            return ResultadoAto("Preparatório", "documento menciona minuta ou ordem de expedição/publicação")
        if categoria == "EDITAL" or _CERT_PUBLICACAO.search(texto):
            if not _CITACAO.search(texto) and re.search(r"\bedital\s+de\s+(?:intimacao|leilao)\b", texto):
                return ResultadoAto("Ignorar", "edital para outra finalidade")
            return ResultadoAto("Revisar", "indicação de edital sem confirmação suficiente do ato")

    if _CERT_PUBLICACAO.search(texto) and re.search(r"\bpublicacao\s+(?:d[eo]\s+)?edital\b", texto):
        return ResultadoAto("Revisar", "publicação menciona edital com finalidade incerta")

    contexto_oficial = _CONTEXTO_OFICIAL.search(texto) or categoria == "MANDADO"
    if contexto_oficial and _RELATO.search(texto):
        negativo = _NEGATIVO.search(texto)
        texto_sem_negacao_direta = _NEGATIVO.sub(" ", texto)
        positivos = list(_CITEI.finditer(texto_sem_negacao_direta))
        if positivos and negativo:
            return ResultadoAto("Revisar", "o retorno contém citação e diligência negativa")
        if positivos and _OUTRO_PROCESSO.search(texto):
            return ResultadoAto("Revisar", "o relato menciona citação em outro processo")
        if any(_MENCAO_INCERTA.search(texto_sem_negacao_direta[:ocorrencia.start()]) for ocorrencia in positivos):
            return ResultadoAto("Revisar", "relato de citação com condição a verificar")
        if positivos:
            return ResultadoAto("Positivo", "citação registrada no retorno do oficial de justiça", "oficial")
        if negativo and _CITACAO.search(texto):
            causas = [motivo for pattern, motivo in _MOTIVOS_DILIGENCIA if pattern.search(texto)]
            motivo = "retorno do oficial indica citação não realizada"
            if len(causas) == 1:
                motivo += f" ({causas[0]})"
            return ResultadoAto("Negativo", motivo)

    if categoria == "MANDADO" and _CITACAO.search(texto) and _PREPARATORIO.search(texto):
        return ResultadoAto("Expedido", "mandado de citação localizado sem registro de cumprimento")
    if not _CITACAO.search(texto) and re.search(r"\b(?:mandado\s+de\s+intimacao|intimei)\b", texto):
        return ResultadoAto("Ignorar", "mandado de intimação")
    if categoria == "CERTIDAO" and not contexto_oficial and not _CITACAO.search(texto):
        return ResultadoAto("Ignorar", "certidão sem indicação de citação")
    return ResultadoAto("Revisar", "texto não permite determinar o resultado da citação")
