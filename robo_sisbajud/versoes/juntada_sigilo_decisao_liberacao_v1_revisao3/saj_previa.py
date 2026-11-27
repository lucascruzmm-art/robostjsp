import csv
import queue
import random
import re
import sys
import threading
import time
import unicodedata
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook

TITULO = 'Robô de Juntada e Liberação Juntos - Sisbajud TJSP'
BASE = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
IMAGENS = BASE / 'imagens_saj'
HISTORICO = BASE / 'historico_saj.csv'
INTERVALO = 3
ESPERA_SAJ = 6
LIMITE_APOS_ALT_L = 8
BG, PANEL, BLUE, TEXT, MUTED = '#e6e6e6', '#f3f3f3', '#005a9e', '#202020', '#606060'
ALVOS = {
    'fluxo': 'Fluxo de Trabalho',
    'digitalizacao': 'Digitalização de Peças — barra superior',
    'processo_rotulo': 'NOVO: somente a palavra Processo, sem ícones nem campo abaixo',
    'pendencias': 'Título Pendências e Prazos do Processo',
    'importar': 'Ícone de folhas sobrepostas',
    'codigo': 'Título Código + filtro vazio abaixo',
    'aguardando': 'Cabeçalho Peças aguardando liberação',
    'salvar': 'Salvar habilitado',
    'caixa_desmarcada': 'somente a caixa vazia de selecionar todas',
    'caixa_marcada': 'a mesma caixa marcada',
    'liberar': 'botão de liberação após marcar a caixa',
    'limpar': 'botão Limpar',
    'sigilo_lista': 'SIGILO 1: Exibir em ordem cronológica da lista superior',
    'sigilo_grupo': 'SIGILO 2: texto Peças sigilosas',
    'sigilo_decisao_texto': 'NOVO: somente a palavra Decisão abaixo do grupo, sem seleção',
    'sigilo_decisao_azul': 'NOVO: somente a palavra Decisão selecionada, com o fundo azul',
    'sigilo_cadeado': 'SIGILO: cadeado ativo da barra, incluindo a seta do menu',
    'sigilo_remover': 'SIGILO: opção Remover sigilo',
}


class ParadaSolicitada(RuntimeError):
    pass


def conferir_aparencia_botao(referencia, captura):
    import numpy as np
    if referencia.size != captura.size:
        return False
    modelo = np.asarray(referencia.convert('RGB'), dtype=np.uint8)
    atual = np.asarray(captura.convert('RGB'), dtype=np.float32)
    cores, contagens = np.unique(modelo.reshape(-1, 3), axis=0, return_counts=True)
    fundo = cores[int(contagens.argmax())].astype(np.float32)
    modelo = modelo.astype(np.float32)
    tracos = np.abs(modelo-fundo).mean(axis=2) >= 30
    if np.count_nonzero(tracos) < 6:
        raise ValueError('A imagem do botão precisa incluir seu texto ou ícone habilitado.')
    diferenca = np.abs(atual-modelo).mean(axis=2)[tracos]
    return float(diferenca.mean()) <= 35 and float(np.mean(diferenca <= 50)) >= .80



def campo_limpo_visual(captura, ponto):
    from PIL import Image
    import cv2
    import numpy as np
    if not isinstance(captura, Image.Image) or min(captura.size) <= 0:
        return None
    try:
        px, py = (int(v) for v in ponto)
    except (TypeError, ValueError, OverflowError):
        return None
    largura, altura = captura.size
    if not (0 <= px < largura and 0 <= py < altura):
        return None

    rgb = np.asarray(captura.convert('RGB'), dtype=np.uint8)
    minimo = rgb.min(axis=2).astype(np.int16)
    maximo = rgb.max(axis=2).astype(np.int16)
    luz = rgb.astype(np.float32).mean(axis=2)
    # Branco quase neutro: não confundir o azul de seleção com uma caixa vazia.
    branco = ((minimo >= 230) & (luz >= 235) & (maximo-minimo <= 18)).astype(np.uint8)
    total, _, stats, _ = cv2.connectedComponentsWithStats(branco, connectivity=8)
    candidatos = []
    for i in range(1, total):
        x, y, w, h, area = (int(v) for v in stats[i])
        if (w >= 30 and 8 <= h <= 40 and area >= 80
                and x <= px < x+w and y <= py < y+h):
            candidatos.append((w*h, x, y, w, h))
    if not candidatos:
        return None
    # Uma região branca externa pode ter uma bbox que engloba o campo;
    # prefere a caixa menor que contém o ponto esperado.
    _, x, y, w, h = min(candidatos)
    dentro = rgb[y+1:y+h-1, x+1:x+w-1].astype(np.int16)
    if min(dentro.shape[:2]) < 4:
        return None
    luz_dentro = dentro.mean(axis=2)
    saturacao = dentro.max(axis=2)-dentro.min(axis=2)

    # Uma seleção azul parcial pode deixar branco ao redor, cuja bbox ainda
    # contém o clique. Qualquer mancha cromática substancial é inconclusiva.
    colorido = ((saturacao > 35) & (luz_dentro < 235)).astype(np.uint8)
    total, _, stats, _ = cv2.connectedComponentsWithStats(colorido, connectivity=8)
    for i in range(1, total):
        _, _, cw, ch, area = (int(v) for v in stats[i])
        if cw >= 3 and ch >= 4 and area >= 12:
            return None

    escuro = ((luz_dentro <= 180) & (saturacao <= 35)).astype(np.uint8)
    # O caret pode tocar num ponto/traço da máscara, formando uma componente
    # mais larga. Remove primeiro hastes retas de 1–2px altas em relação à caixa.
    # Os dígitos das fontes de campo testadas são mais baixos que esse limiar.
    minimo_caret = max(8, int(np.ceil(escuro.shape[0] * .65)))
    hastes = []
    # Só remove hastes na margem inicial do campo. Remover hastes em toda a
    # largura apagaria dígitos '1' em campos compactos.
    for cx in range(min(3, escuro.shape[1])):
        coluna = escuro[:, cx].astype(np.int8)
        transicoes = np.diff(np.pad(coluna, (1, 1)))
        inicios = np.flatnonzero(transicoes == 1)
        finais = np.flatnonzero(transicoes == -1)
        for cy, fim in zip(inicios, finais):
            if fim-cy >= minimo_caret:
                hastes.append((cx, int(cy), int(fim)))
    grupos = []
    for haste in hastes:
        if (grupos and haste[0] == grupos[-1][-1][0] + 1
                and abs(haste[1]-grupos[-1][-1][1]) <= 2
                and abs(haste[2]-grupos[-1][-1][2]) <= 2):
            grupos[-1].append(haste)
        else:
            grupos.append([haste])
    for grupo in grupos:
        if len(grupo) <= 2:
            for cx, cy, fim in grupo:
                escuro[cy:fim, cx] = 0
    total, _, stats, _ = cv2.connectedComponentsWithStats(escuro, connectivity=8)
    estreitos_altos = 0
    for i in range(1, total):
        _, _, cw, ch, area = (int(v) for v in stats[i])
        # Pontos e traços da máscara são baixos; bordas já foram excluídas.
        if ch <= 3:
            continue
        if cw <= 2:
            if ch >= 5 and area >= 4:
                estreitos_altos += 1
            continue
        if cw >= 3 and ch >= 4 and area >= 5:
            return False
    # Várias hastes altas não são uma máscara vazia. Uma haste estreita
    # restante é ambígua (não foi confirmada como caret), então não conclui.
    if estreitos_altos >= 2:
        return False
    if estreitos_altos:
        return None
    return True


def normalizar(valor):
    return ''.join(c for c in unicodedata.normalize('NFD', str(valor or '').upper()) if unicodedata.category(c) != 'Mn').strip()


def digitos(valor):
    return re.sub(r'\D', '', str(valor))


def normalizar_regiao(regiao, tamanho=None):
    if regiao is None:
        return None
    if not isinstance(regiao, (tuple, list)) or len(regiao) != 4:
        raise ValueError('A área de busca precisa conter quatro valores: x, y, largura e altura.')
    try:
        x, y, w, h = (int(v) for v in regiao)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError('A área de busca contém coordenadas inválidas.') from exc
    if w <= 0 or h <= 0:
        raise ValueError('A área de busca está vazia: largura e altura devem ser positivas.')
    direita, baixo = x+w, y+h
    if tamanho is not None:
        largura_tela, altura_tela = (int(v) for v in tamanho)
        x, y = max(0, x), max(0, y)
        direita, baixo = min(largura_tela, direita), min(altura_tela, baixo)
    if direita <= x or baixo <= y:
        raise ValueError('A área de busca está fora da parte visível da tela.')
    return (int(x), int(y), int(direita-x), int(baixo-y))


def carregar_referencia(chave):
    from PIL import Image
    caminho = IMAGENS / (chave + '.png')
    if not caminho.is_file():
        raise ValueError('Imagem não encontrada: ' + caminho.name)
    with Image.open(caminho) as imagem:
        imagem.load()
        if imagem.format != 'PNG' or min(imagem.size) < 7:
            raise ValueError('Imagem inválida ou recorte pequeno demais: ' + caminho.name)
        referencia = imagem.convert('RGB').copy()
        if chave == 'sigilo_decisao_azul':
            recorte_selecao_azul(referencia)
        if chave == 'sigilo_cadeado':
            conferir_cores_cadeado(referencia, referencia)
        if chave in ('salvar', 'limpar'):
            conferir_aparencia_botao(referencia, referencia)
        return referencia


def recorte_selecao_azul(imagem):
    pixels = imagem.load()
    pontos = []
    for y in range(imagem.height):
        for x in range(imagem.width):
            r, g, b = pixels[x, y][:3]
            if b > 80 and b-r > 25 and b-g > 10:
                pontos.append((x, y))
    if len(pontos) < 25:
        raise ValueError('Capture somente a palavra Decisão selecionada, incluindo o fundo azul.')
    x1, y1 = min(x for x, y in pontos), min(y for x, y in pontos)
    x2, y2 = max(x for x, y in pontos)+1, max(y for x, y in pontos)+1
    if x2-x1 < 14 or y2-y1 < 6 or len(pontos)/((x2-x1)*(y2-y1)) < .45:
        raise ValueError('Recorte insuficiente. Inclua a palavra Decisão inteira e seu fundo azul.')
    return imagem.crop((x1, y1, x2, y2)), (x1, y1, x2-x1, y2-y1)


def conferir_cores_cadeado(referencia, captura):
    import numpy as np
    if referencia.size != captura.size:
        return False
    modelo = np.asarray(referencia.convert('RGB'), dtype=np.float32)
    atual = np.asarray(captura.convert('RGB'), dtype=np.float32)
    luz = modelo.mean(axis=2)
    fundo = float(np.percentile(luz, 85))
    tracos = luz <= min(125, fundo-55)
    if np.count_nonzero(tracos) < 6:
        raise ValueError('Refaça sigilo_cadeado.png: capture o cadeado ATIVO, com traços escuros, e a pequena seta. Não capture o cadeado cinza desativado.')
    diferenca = np.abs(atual-modelo).mean(axis=2)[tracos]
    return float(diferenca.mean()) <= 32 and float(np.mean(diferenca <= 45)) >= .80


def regiao_cadeado_sigilo(painel, tamanho, referencia):
    largura, altura = (int(v) for v in tamanho)
    direita_arvore = int(painel[2])
    escala = max(.75, min(3.0, referencia.width/20, referencia.height/12))
    esquerda = max(0, direita_arvore-2)
    direita = min(largura, direita_arvore+max(32, round(45*escala)))
    topo, baixo = max(0, int(painel[1])), min(altura, int(painel[3]))
    if direita <= esquerda or baixo <= topo:
        raise RuntimeError('A barra lateral superior está fora da área visível. O cadeado não será acionado.')
    regiao = normalizar_regiao((esquerda, topo, direita-esquerda, baixo-topo), tamanho)
    if referencia.width > regiao[2] or referencia.height > regiao[3]:
        raise RuntimeError('Refaça sigilo_cadeado.png com um recorte pequeno do cadeado ativo e sua seta. A imagem não cabe na barra lateral superior.')
    return regiao


def encontrar_contornos(referencia, captura, regiao):
    # O contorno das letras permite reencontrar o grupo quando seu fundo fica azul.
    import cv2
    import numpy as np
    from pyscreeze import Box
    margem = 2
    molde = referencia.crop((margem, margem, referencia.width-margem, referencia.height-margem))
    bordas = cv2.Canny(cv2.cvtColor(np.array(molde), cv2.COLOR_RGB2GRAY), 40, 100)
    tela = cv2.Canny(cv2.cvtColor(np.array(captura), cv2.COLOR_RGB2GRAY), 40, 100)
    if bordas.shape[0] > tela.shape[0] or bordas.shape[1] > tela.shape[1] or np.count_nonzero(bordas) < 20:
        return None
    resultado = cv2.matchTemplate(tela, bordas, cv2.TM_CCOEFF_NORMED)
    candidatos = []
    for _ in range(2):
        _, nota, _, (x, y) = cv2.minMaxLoc(resultado)
        if nota < .82:
            break
        candidatos.append((x, y))
        dx, dy = max(6, molde.width//2), max(6, molde.height//2)
        resultado[max(0, y-dy):y+dy+1, max(0, x-dx):x+dx+1] = -1
    if len(candidatos) > 1:
        raise RuntimeError('Mais de um grupo Peças sigilosas encontrado na lista superior.')
    if not candidatos:
        return None
    x, y = candidatos[0]
    return Box(int(regiao[0]+x-margem), int(regiao[1]+y-margem), int(referencia.width), int(referencia.height))


def referencia_valida(chave):
    try:
        carregar_referencia(chave)
        return True
    except Exception:
        return False


def validar_referencias(chaves):
    erros = []
    for chave in ALVOS:
        if chave in chaves:
            try:
                carregar_referencia(chave)
            except Exception as exc:
                erros.append(ALVOS[chave] + '\n' + str(exc))
    novas = {'sigilo_grupo', 'sigilo_decisao_texto', 'sigilo_decisao_azul'}
    if novas.issubset(chaves) and all(referencia_valida(k) for k in novas):
        grupo = carregar_referencia('sigilo_grupo')
        if any(carregar_referencia(k).height > grupo.height+10 for k in novas-{'sigilo_grupo'}):
            erros.append('Refaça as duas imagens novas da Decisão: capture somente a palavra, sem Peças sigilosas acima ou páginas abaixo.')
    return erros


def ler_planilha(caminho, pasta):
    pasta = Path(pasta)
    if not pasta.is_dir():
        raise ValueError('Selecione uma pasta de PDFs válida.')
    arquivos = {}
    for arquivo in pasta.iterdir():
        if arquivo.is_file() and arquivo.suffix.lower() == '.pdf':
            arquivos.setdefault(digitos(arquivo.stem), []).append(arquivo.resolve())
    livro = load_workbook(caminho, read_only=True, data_only=True)
    try:
        linhas = iter(livro.active.values)
        cabecalhos = [normalizar(v) for v in next(linhas, ())]
        coluna = next((i for i, v in enumerate(cabecalhos) if v in ('PROCESSO', 'NUMERO DO PROCESSO', 'NUMERO PROCESSO')), None)
        if coluna is None or 'STATUS' not in cabecalhos:
            raise ValueError('Cabeçalhos: numero do processo e status.')
        coluna_status = cabecalhos.index('STATUS')
        vistos, resultado = set(), []
        for numero_linha, linha in enumerate(linhas, 2):
            if not any(v is not None for v in linha):
                continue
            original = linha[coluna]
            processo = str(original or '').strip()
            chave, status = digitos(processo), normalizar(linha[coluna_status])
            pdfs = arquivos.get(chave, [])
            erro = ''
            if not isinstance(original, str) or len(chave) != 20:
                erro = 'Processo deve ser Texto com 20 dígitos'
            elif status not in ('IRRISORIO', 'PARCIAL', 'INTEGRAL'):
                erro = 'Status inválido'
            elif chave in vistos:
                erro = 'Processo repetido na planilha'
            elif len(pdfs) != 1:
                erro = 'PDF ausente ou mais de um correspondente'
            vistos.add(chave)
            resultado.append(dict(linha=numero_linha, processo=processo, status=status,
                codigo='938' if status == 'IRRISORIO' else '939', pdf=str(pdfs[0]) if len(pdfs) == 1 else '', erro=erro))
        return resultado
    finally:
        livro.close()


def ler_historico():
    if not HISTORICO.exists():
        return set()
    with HISTORICO.open(encoding='utf-8-sig', newline='') as arquivo:
        return {digitos(r.get('processo', '')) for r in csv.DictReader(arquivo, delimiter=';') if r.get('processo')}


def gravar_historico(registro, fase, resultado):
    novo = not HISTORICO.exists()
    with HISTORICO.open('a', encoding='utf-8-sig', newline='') as arquivo:
        escritor = csv.writer(arquivo, delimiter=';')
        if novo:
            escritor.writerow(['data', 'processo', 'status', 'codigo', 'fase', 'resultado'])
        escritor.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), registro['processo'], registro['status'], registro['codigo'], fase, resultado])


def importar_pdf(caminho, verificar, aguardar, registrar):
    from pywinauto import Desktop
    arquivo = Path(caminho)
    if not arquivo.is_file():
        raise RuntimeError('PDF não está mais disponível.')
    desktop, janela = Desktop(backend='win32'), None
    limite = time.monotonic() + 20
    while time.monotonic() < limite:
        verificar()
        janelas = [w for w in desktop.windows() if w.is_visible() and w.window_text().strip().casefold() == 'importar arquivo']
        if len(janelas) > 1:
            raise RuntimeError('Mais de uma janela Importar arquivo aberta.')
        if janelas:
            janela = janelas[0]
            break
        aguardar(.25)
    if janela is None or not janela.is_enabled():
        raise RuntimeError('Janela Importar arquivo ausente ou bloqueada.')
    candidatos = {}
    for controle in janela.descendants():
        if controle.class_name() != 'Edit':
            continue
        corresponde = controle.control_id() in (1148, 1152)
        ancestral = controle.parent()
        for _ in range(4):
            if ancestral is None or ancestral.handle == janela.handle:
                break
            if ancestral.control_id() == 1148:
                corresponde = True
                break
            ancestral = ancestral.parent()
        if corresponde and controle.is_visible() and controle.is_enabled():
            candidatos[controle.handle] = controle
    if len(candidatos) != 1:
        raise RuntimeError('Campo Nome não identificado com segurança. Enter não enviado.')
    campo = next(iter(candidatos.values()))
    verificar()
    janela.set_focus()
    campo.click_input()
    aguardar()
    verificar()
    campo.set_edit_text(arquivo.name)
    aguardar()
    if campo.window_text() != arquivo.name:
        raise RuntimeError('Nome do PDF não confere. Enter não enviado.')
    verificar()
    campo.set_focus()
    if not campo.has_focus() or campo.window_text() != arquivo.name:
        raise RuntimeError('Foco ou nome do PDF não conferem. Enter não enviado.')
    verificar()
    registrar('PDF conferido: ' + arquivo.name)
    campo.type_keys('{ENTER}', set_foreground=False)
    aguardar()
    limite = time.monotonic() + 15
    while time.monotonic() < limite:
        verificar()
        try:
            aberta = janela.is_visible()
        except Exception:
            aberta = False
        if not aberta:
            return
        aguardar(.25)
    raise RuntimeError('Importação permaneceu aberta. Confira a mensagem no SAJ.')


def ler_barra_vertical(hwnd):
    import ctypes
    from ctypes import wintypes

    class SCROLLINFO(ctypes.Structure):
        _fields_ = [('cbSize', wintypes.UINT), ('fMask', wintypes.UINT),
                    ('nMin', wintypes.INT), ('nMax', wintypes.INT),
                    ('nPage', wintypes.UINT), ('nPos', wintypes.INT),
                    ('nTrackPos', wintypes.INT)]

    class SCROLLBARINFO(ctypes.Structure):
        _fields_ = [('cbSize', wintypes.DWORD), ('rcScrollBar', wintypes.RECT),
                    ('dxyLineButton', wintypes.LONG), ('xyThumbTop', wintypes.LONG),
                    ('xyThumbBottom', wintypes.LONG), ('reserved', wintypes.LONG),
                    ('rgstate', wintypes.DWORD * 6)]

    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.GetScrollInfo.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.POINTER(SCROLLINFO)]
    user32.GetScrollInfo.restype = wintypes.BOOL
    user32.GetScrollBarInfo.argtypes = [wintypes.HWND, wintypes.LONG, ctypes.POINTER(SCROLLBARINFO)]
    user32.GetScrollBarInfo.restype = wintypes.BOOL
    info = SCROLLINFO()
    info.cbSize, info.fMask = ctypes.sizeof(info), 0x17
    if not user32.GetScrollInfo(hwnd, 1, ctypes.byref(info)):
        return None
    if info.nPage == 0 and info.nMax <= info.nMin:
        return None
    limite = max(info.nMin, info.nMax - max(int(info.nPage) - 1, 0))
    dados = {'posicao': info.nPos, 'minimo': int(info.nMin), 'limite': limite,
             'sem_rolagem': limite <= info.nMin and info.nPage > 0}
    barra = SCROLLBARINFO()
    barra.cbSize = ctypes.sizeof(barra)
    if user32.GetScrollBarInfo(hwnd, -5, ctypes.byref(barra)):
        r = barra.rcScrollBar
        dados.update(rect=(r.left, r.top, r.right, r.bottom),
                     seta=int(barra.dxyLineButton), thumb_top=int(barra.xyThumbTop),
                     thumb_bottom=int(barra.xyThumbBottom), estado=int(barra.rgstate[0]))
    return dados


def pontos_arraste(dados, para_topo=False):
    esquerda, topo, direita, baixo = dados['rect']
    seta = dados['seta']
    inicio_thumb, fim_thumb = dados['thumb_top'], dados['thumb_bottom']
    if (direita-esquerda < 4 or baixo-topo <= 2*seta+8
            or fim_thumb <= inicio_thumb or inicio_thumb < seta
            or fim_thumb > baixo-topo-seta+2
            or dados['estado'] & (0x1 | 0x8000 | 0x10000)):
        raise RuntimeError('A barra de documentos não está disponível para arrastar.')
    x = (esquerda+direita)//2
    destino = topo+seta+2 if para_topo else baixo-seta-2
    return (int(x), int(topo+(inicio_thumb+fim_thumb)//2)), (int(x), int(destino))


class Motor:
    def __init__(self, parar, evento):
        import pyautogui as pg
        self.pg, self.parar, self.evento = pg, parar, evento
        pg.FAILSAFE, pg.PAUSE = True, .25
        self.fase = 'preparação'

    def verificar(self):
        if self.parar.is_set():
            raise ParadaSolicitada('Interrompido pelo usuário')
        try:
            self.pg.failSafeCheck()
        except self.pg.FailSafeException as exc:
            raise ParadaSolicitada('Parada de emergência pelo mouse') from exc

    def aguardar(self, segundos=INTERVALO):
        limite = time.monotonic() + segundos
        while time.monotonic() < limite:
            self.verificar()
            time.sleep(min(.1, max(0, limite-time.monotonic())))

    def log(self, texto):
        self.evento('log', texto)

    def etapa(self, texto):
        self.fase = texto
        self.evento('fase', texto)
        self.log('Etapa: ' + texto)

    def localizar(self, chave, tempo=20, opcional=False, regiao=None):
        import pyscreeze
        if regiao is not None:
            regiao = normalizar_regiao(regiao, self.pg.size())
        referencia = carregar_referencia(chave)
        # Abrir pelo Pillow evita falhas do OpenCV com acentos no caminho.
        limite = time.monotonic() + tempo
        while time.monotonic() < limite:
            self.verificar()
            try:
                caixa = self.pg.locateOnScreen(referencia, confidence=.92, region=regiao)
            except (self.pg.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                caixa = None
            if caixa:
                return caixa
            self.aguardar(min(.25, max(0, limite-time.monotonic())))
        if not opcional:
            raise RuntimeError('Não localizado: ' + ALVOS[chave] + ' [' + chave + '.png]')
        return None

    def localizar_rotulo_processo(self, tempo=20, opcional=False):
        largura, altura = (int(v) for v in self.pg.size())
        topo = min(90, max(0, altura-1))
        regiao = (0, topo, min(450, largura), min(300, altura-topo))
        # Usa o cabeçalho já calibrado só para delimitar a faixa superior.
        # Se ele não aparecer, a procura continua na região superior esquerda.
        if referencia_valida('sigilo_lista'):
            cabecalho = self.localizar('sigilo_lista', tempo=.5, opcional=True,
                regiao=(0, topo, min(600, largura), altura-topo))
            if cabecalho:
                x = max(0, int(cabecalho.left)-10)
                baixo = int(cabecalho.top)
                y = max(topo, baixo-180)
                if baixo > y:
                    regiao = (x, y, min(450, largura-x), baixo-y)
        return self.localizar_unico('processo_rotulo', regiao=regiao,
                                    tempo=tempo, opcional=opcional, confianca=.96)

    def clicar_campo_processo(self):
        self.etapa('localizar a palavra Processo e clicar na caixa abaixo')
        rotulo = self.localizar_rotulo_processo()
        x = int(rotulo.left+rotulo.width//2)
        y = int(rotulo.top+rotulo.height+max(8, round(rotulo.height*.65)))
        largura, altura = self.pg.size()
        if not (0 <= x < largura and 0 <= y < altura):
            raise RuntimeError('A caixa abaixo de Processo está fora da área visível.')
        self.verificar()
        self.pg.click(x, y)
        self.aguardar()

    def clicar(self, chave, inferior=False, pausa=INTERVALO, regiao=None):
        caixa = self.localizar(chave, regiao=regiao)
        self.verificar()
        self.pg.click(int(caixa.left + caixa.width // 2), int(caixa.top + (int(caixa.height * .78) if inferior else caixa.height // 2)))
        self.aguardar(pausa)

    def tecla(self, tecla):
        self.verificar()
        self.pg.press(tecla)
        self.aguardar()

    def saj_disponivel(self, caixa):
        from pywinauto import Desktop
        import win32gui
        import ctypes
        from ctypes import wintypes
        self.verificar()
        try:
            x, y = int(caixa.left+caixa.width//2), int(caixa.top+caixa.height//2)
            controle = Desktop(backend='win32').from_point(x, y)
            raiz = controle.top_level_parent()
            titulo = normalizar(raiz.window_text())
            saj = 'SAJ' in titulo or 'SISTEMA DE AUTOMACAO DA JUSTICA' in titulo or 'DIGITALIZACAO' in titulo
            if (not saj or titulo == normalizar(TITULO) or not controle.is_visible()
                    or not controle.is_enabled() or not raiz.is_enabled()
                    or int(win32gui.GetForegroundWindow()) != int(raiz.handle)):
                return False
            user32 = ctypes.WinDLL('user32', use_last_error=True)
            user32.IsHungAppWindow.argtypes = [wintypes.HWND]
            user32.IsHungAppWindow.restype = wintypes.BOOL
            return not bool(user32.IsHungAppWindow(int(raiz.handle)))
        except Exception:
            return False

    def cancelar_importacao_aberta(self):
        from pywinauto import Desktop
        import win32gui
        self.verificar()
        janelas = [w for w in Desktop(backend='win32').windows()
                   if w.is_visible() and w.is_enabled()
                   and normalizar(w.window_text()) == 'IMPORTAR ARQUIVO']
        if len(janelas) != 1:
            return
        janela = janelas[0]
        if int(win32gui.GetForegroundWindow()) != int(janela.handle):
            return
        dono = win32gui.GetWindow(int(janela.handle), 4)
        pertence_saj, visitados = False, set()
        for _ in range(8):
            if not dono or dono in visitados:
                break
            visitados.add(dono)
            titulo_dono = normalizar(win32gui.GetWindowText(dono))
            if 'SAJ' in titulo_dono or 'SISTEMA DE AUTOMACAO DA JUSTICA' in titulo_dono or 'DIGITALIZACAO' in titulo_dono:
                pertence_saj = titulo_dono != normalizar(TITULO)
                break
            dono = win32gui.GetWindow(dono, 4) or win32gui.GetParent(dono)
        if not pertence_saj:
            return
        self.log('Cancelando a janela Importar arquivo do processo com erro.')
        janela.set_focus()
        self.verificar()
        janela.type_keys('{ESC}', set_foreground=False)
        self.aguardar()

    def clicar_quando_disponivel(self, chave, pausa=INTERVALO, recuperar=False, limite=None):
        referencia = carregar_referencia(chave)
        aviso_prazo = False

        def conferir_prazo():
            if limite is not None and time.monotonic() >= limite:
                raise TimeoutError(
                    f'{ALVOS[chave]} não ficou disponível em até '
                    f'{LIMITE_APOS_ALT_L} segundos após Alt + L.'
                )

        while True:
            self.verificar()
            conferir_prazo()
            if recuperar:
                self.cancelar_importacao_aberta()
            tempo_busca = .6 if limite is None else min(.3, max(0, limite-time.monotonic()))
            caixa = self.localizar(chave, tempo=tempo_busca, opcional=True)
            conferir_prazo()
            if caixa and self.saj_disponivel(caixa):
                area = normalizar_regiao((caixa.left, caixa.top, caixa.width, caixa.height), self.pg.size())
                captura = self.pg.screenshot(region=area)
                if conferir_aparencia_botao(referencia, captura):
                    self.verificar()
                    conferir_prazo()
                    self.evento('fase', self.fase)
                    self.pg.click(int(caixa.left+caixa.width//2), int(caixa.top+caixa.height//2))
                    self.aguardar(pausa)
                    return
            conferir_prazo()
            self.evento('fase', 'Aguardando o SAJ: '+ALVOS[chave])
            if limite is None:
                self.log(ALVOS[chave]+' ainda indisponível. Nova tentativa em 6 segundos.')
                self.aguardar(ESPERA_SAJ)
            else:
                if not aviso_prazo:
                    self.log('Aguardando '+ALVOS[chave]+f', respeitando o limite total de {LIMITE_APOS_ALT_L} segundos após Alt + L.')
                    aviso_prazo = True
                self.aguardar(min(.2, max(0, limite-time.monotonic())))

    def verificar_campo_limpo_visual(self, rotulo):
        largura, altura = (int(v) for v in self.pg.size())
        x = max(0, int(rotulo.left)-4)
        y = int(rotulo.top+rotulo.height)+1
        area = normalizar_regiao((x, y, min(320, max(180, int(rotulo.width)*6)),
                                 max(45, int(rotulo.height)*4)), (largura, altura))
        px = int(rotulo.left+rotulo.width//2)-area[0]
        py = int(rotulo.top+rotulo.height+max(8, round(rotulo.height*.65)))-area[1]
        self.verificar()
        captura = self.pg.screenshot(region=area)
        estado = campo_limpo_visual(captura, (px, py))
        if estado is None:
            # Uma nova captura evita depender da fase de piscar do cursor.
            self.aguardar(.25)
            captura = self.pg.screenshot(region=area)
            estado = campo_limpo_visual(captura, (px, py))
        return estado

    def limpar_para_proximo(self, recuperar=False):
        while True:
            self.etapa('limpar após erro para seguir ao próximo processo' if recuperar else 'limpar para o próximo processo')
            self.clicar_quando_disponivel('limpar', recuperar=recuperar)
            cursor_ajustado = False
            while True:
                self.verificar()
                rotulo = self.localizar_rotulo_processo(tempo=1, opcional=True)
                if rotulo and self.saj_disponivel(rotulo):
                    estado = self.verificar_campo_limpo_visual(rotulo)
                    if estado is True:
                        self.log('Tela limpa e campo Processo preparado para o próximo número.')
                        return
                    if estado is False:
                        self.log('O campo Processo ainda contém conteúdo. Aguardando 6 segundos para tentar Limpar novamente.')
                        self.aguardar(ESPERA_SAJ)
                        break
                    if estado is None and not cursor_ajustado:
                        self.log('Posicionando o cursor no início do campo Processo para conferir a limpeza.')
                        self.clicar_campo_processo()
                        self.tecla('home')
                        cursor_ajustado = True
                        continue
                self.evento('fase', 'Aguardando o SAJ concluir a limpeza')
                self.log('Aguardando a tela de digitalização ficar preparada. Nova verificação em 6 segundos.')
                self.aguardar(ESPERA_SAJ)

    def recuperar_apos_erro(self):
        while True:
            self.verificar()
            try:
                self.limpar_para_proximo(recuperar=True)
                return
            except ParadaSolicitada:
                raise
            except Exception as exc:
                if isinstance(exc, getattr(self.pg, 'FailSafeException', ParadaSolicitada)):
                    raise ParadaSolicitada('Parada de emergência pelo mouse') from exc
                self.evento('fase', 'Aguardando recuperação da tela do SAJ')
                self.log('A limpeza ainda não pôde ser concluída: '+str(exc)+'. Nova tentativa em 6 segundos.')
                self.aguardar(ESPERA_SAJ)

    def localizar_unico(self, chave, regiao=None, tempo=8, opcional=False, confianca=.94):
        import pyscreeze
        if regiao is not None:
            regiao = normalizar_regiao(regiao, self.pg.size())
        referencia = carregar_referencia(chave)
        if regiao and (referencia.width > regiao[2] or referencia.height > regiao[3]):
            if opcional:
                return None
            raise RuntimeError('A área visível é menor que a imagem: ' + chave + '.png')
        limite = time.monotonic() + tempo
        while time.monotonic() < limite:
            self.verificar()
            try:
                encontrados = list(self.pg.locateAllOnScreen(referencia, confidence=confianca, region=regiao, grayscale=False))
            except (self.pg.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                encontrados = []
            distintos = []
            for caixa in encontrados:
                if chave == 'sigilo_cadeado':
                    area = normalizar_regiao((caixa.left, caixa.top, caixa.width, caixa.height), self.pg.size())
                    captura = self.pg.screenshot(region=area)
                    if not conferir_cores_cadeado(referencia, captura):
                        continue
                if not any(abs(caixa.left-outra.left) < max(6, caixa.width//2)
                           and abs(caixa.top-outra.top) < max(6, caixa.height//2) for outra in distintos):
                    distintos.append(caixa)
            if len(distintos) > 1:
                if chave == 'sigilo_cadeado':
                    raise RuntimeError('Mais de um cadeado com aparência ativa encontrado na barra lateral superior. Refaça apenas sigilo_cadeado.png com o ícone ativo e a pequena seta, sem outros botões.')
                raise RuntimeError('Mais de um alvo encontrado: ' + ALVOS[chave])
            if distintos:
                return distintos[0]
            self.aguardar(.2)
        if not opcional:
            raise RuntimeError('Não localizado: ' + ALVOS[chave] + ' [' + chave + '.png]')
        return None

    def arrastar_documentos_ate_fim(self):
        from pywinauto import Desktop
        self.etapa('verificar se a lista de documentos precisa de rolagem')
        largura, altura = self.pg.size()
        cabecalho = self.localizar_unico('sigilo_lista', regiao=(0, 90, min(500, largura), altura-90))
        x = cabecalho.left+cabecalho.width//2
        y = cabecalho.top+cabecalho.height+22
        controle = Desktop(backend='win32').from_point(int(x), int(y))
        alvo = None
        for _ in range(8):
            self.verificar()
            if controle is None:
                break
            r = controle.rectangle()
            classe = controle.class_name().lower()
            eh_lista = 'tree' in classe or 'list' in classe or callable(getattr(controle, 'roots', None))
            if (eh_lista and controle.is_visible() and controle.is_enabled()
                    and r.left >= 0 and r.top >= 0 and r.right <= largura and r.bottom <= altura
                    and cabecalho.top+cabecalho.height-15 <= r.top <= cabecalho.top+cabecalho.height+20
                    and 90 <= r.right-r.left <= min(500, largura//2)
                    and r.bottom-r.top > 60 and r.left <= x < r.right and r.top <= y < r.bottom):
                # A lista é identificada mesmo quando não existe barra vertical.
                alvo = controle
                break
            controle = controle.parent()
        if alvo is None:
            raise RuntimeError('Não foi possível identificar a lista superior de documentos. A execução parou sem alterar sigilo.')

        for _ in range(3):
            self.verificar()
            if not alvo.is_visible() or not alvo.is_enabled():
                raise RuntimeError('A lista superior de documentos não está disponível.')
            r = alvo.rectangle()
            painel = (int(r.left), int(r.top), int(r.right), int(r.bottom))
            if not alvo.has_style(0x00200000):
                self.log('Lista superior sem barra vertical: todos os documentos já estão visíveis. Seguindo sem arrastar.')
                return painel
            dados = ler_barra_vertical(alvo.handle)
            if dados is None:
                if not alvo.has_style(0x00200000):
                    self.log('Lista superior sem barra vertical: todos os documentos já estão visíveis. Seguindo sem arrastar.')
                    return painel
                raise RuntimeError('Existe barra vertical na lista superior, mas não foi possível ler sua posição.')
            if dados['sem_rolagem']:
                self.log('Todos os documentos já cabem na lista superior. Seguindo sem arrastar.')
                return painel
            if dados['posicao'] >= dados['limite']:
                self.log('Fim da lista de documentos confirmado.')
                return painel
            if 'rect' not in dados:
                raise RuntimeError('A barra de documentos não forneceu uma posição válida para arrastar.')
            b = dados['rect']
            if not (r.left+60 <= b[0] < b[2] <= r.right+5
                    and r.top-5 <= b[1] < b[3] <= r.bottom+5):
                raise RuntimeError('A barra localizada não pertence à lista superior de documentos.')
            inicio, fim = pontos_arraste(dados)
            if not all(0 <= px < largura and 0 <= py < altura for px, py in (inicio, fim)):
                raise RuntimeError('A barra de documentos está fora da área visível da tela.')
            self.etapa('arrastar a barra de documentos até o fim')
            self.pg.moveTo(*inicio, duration=.15)
            self.verificar()
            self.pg.dragTo(*fim, duration=.7, button='left')
            self.aguardar()
        dados = ler_barra_vertical(alvo.handle)
        if dados and (dados['sem_rolagem'] or dados['posicao'] >= dados['limite']):
            r = alvo.rectangle()
            self.log('Fim da lista de documentos confirmado.')
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
        if dados is None and not alvo.has_style(0x00200000):
            r = alvo.rectangle()
            self.log('Lista superior sem barra vertical. Seguindo sem novo arraste.')
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
        raise RuntimeError('A lista não chegou ao fim após o arraste. O robô não vai presumir ausência de peças sigilosas.')

    def localizar_grupo_sigilo(self, regiao, tempo=5, opcional=False):
        regiao = normalizar_regiao(regiao, self.pg.size())
        referencia = carregar_referencia('sigilo_grupo')
        limite = time.monotonic()+tempo
        while time.monotonic() < limite:
            self.verificar()
            grupo = self.localizar_unico('sigilo_grupo', regiao=regiao, tempo=.3, opcional=True)
            if grupo:
                return grupo
            captura = self.pg.screenshot(region=regiao)
            grupo = encontrar_contornos(referencia, captura, regiao)
            if grupo:
                return grupo
            self.aguardar(.2)
        if not opcional:
            raise RuntimeError('Não foi possível reencontrar Peças sigilosas na lista superior.')
        return None

    def faixa_primeira_decisao(self, grupo, painel):
        normal = carregar_referencia('sigilo_decisao_texto')
        azul = carregar_referencia('sigilo_decisao_azul')
        altura = max(normal.height, azul.height)
        if altura > grupo.height+10:
            raise RuntimeError('Refaça as duas novas imagens da Decisão: capture somente a palavra, sem o grupo acima ou páginas abaixo.')
        esquerda = max(int(painel[0]), int(grupo.left)+max(8, int(grupo.height)//2))
        topo = max(int(painel[1]), int(grupo.top+grupo.height)-1)
        direita = int(painel[2])
        baixo = min(int(painel[3]), topo+altura+8)
        if direita <= esquerda or baixo <= topo:
            return None
        return normalizar_regiao((esquerda, topo, direita-esquerda, baixo-topo), self.pg.size())

    def contexto_decisao_sigilo(self, grupo, painel, tempo=5):
        regiao = self.faixa_primeira_decisao(grupo, painel)
        if regiao is None:
            return None
        limite = time.monotonic() + tempo
        while time.monotonic() < limite:
            self.verificar()
            selecionada = self.localizar_unico('sigilo_decisao_azul', regiao=regiao, tempo=.4, opcional=True, confianca=.96)
            if selecionada and self.conferir_azul_decisao(selecionada):
                return 'selecionada', selecionada
            normal = self.localizar_unico('sigilo_decisao_texto', regiao=regiao, tempo=.4, opcional=True, confianca=.94)
            if normal:
                return 'normal', normal
            self.aguardar(.2)
        return None

    def conferir_azul_decisao(self, caixa):
        import pyscreeze
        referencia = carregar_referencia('sigilo_decisao_azul')
        patch, (dx, dy, w, h) = recorte_selecao_azul(referencia)
        largura, altura = self.pg.size()
        x, y = max(0, caixa.left+dx-2), max(0, caixa.top+dy-2)
        regiao = (x, y, min(w+4, largura-x), min(h+4, altura-y))
        self.verificar()
        captura = self.pg.screenshot(region=normalizar_regiao(regiao, (largura, altura)))
        try:
            return self.pg.locate(patch, captura, confidence=.96, grayscale=False) is not None
        except (self.pg.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            return False

    def tratar_sigilo(self):
        self.etapa('verificar peças sigilosas por imagem')
        painel = self.arrastar_documentos_ate_fim()
        largura, altura = self.pg.size()
        esquerda, direita = max(0, painel[0]), min(largura, painel[2])
        topo = max(0, painel[1])
        regiao_grupo = (esquerda, topo, direita-esquerda, min(altura, painel[3])-topo)
        grupo = self.localizar_grupo_sigilo(regiao_grupo, tempo=5, opcional=True)
        if grupo is None:
            self.log('Grupo Peças sigilosas ausente. Seguindo para a liberação normal.')
            return

        contexto = self.contexto_decisao_sigilo(grupo, painel, tempo=1)
        if contexto is None:
            self.etapa('clicar uma vez no texto Peças sigilosas')
            grupo = self.localizar_grupo_sigilo(regiao_grupo)
            self.verificar()
            self.pg.click(int(grupo.left+grupo.width//2), int(grupo.top+grupo.height//2))
            self.aguardar()
            self.etapa('localizar a palavra Decisão na linha abaixo do grupo')
            grupo = self.localizar_grupo_sigilo(regiao_grupo)
            contexto = self.contexto_decisao_sigilo(grupo, painel, tempo=5)
            if contexto is None:
                # A expansão pode colocar o filho abaixo da área visível.
                painel = self.arrastar_documentos_ate_fim()
                grupo = self.localizar_grupo_sigilo(regiao_grupo)
                contexto = self.contexto_decisao_sigilo(grupo, painel, tempo=6)
            if contexto is None:
                raise RuntimeError('Decisão não encontrada na primeira linha abaixo de Peças sigilosas. Confira sigilo_decisao_texto.png: capture somente a palavra Decisão, sem o grupo ou páginas. Nenhum sigilo foi removido.')

        estado, caixa = contexto
        if estado != 'selecionada':
            self.etapa('clicar somente na Decisão abaixo de Peças sigilosas')
            self.verificar()
            self.pg.click(int(caixa.left+caixa.width//2), int(caixa.top+caixa.height//2))
            self.aguardar()
        else:
            self.log('A Decisão do grupo já está selecionada.')

        self.etapa('confirmar visualmente a Decisão selecionada')
        limite = time.monotonic()+8
        confirmada = False
        while time.monotonic() < limite:
            grupo = self.localizar_grupo_sigilo(regiao_grupo, tempo=1)
            faixa = self.faixa_primeira_decisao(grupo, painel)
            caixa = None if faixa is None else self.localizar_unico('sigilo_decisao_azul', regiao=faixa, tempo=.6, opcional=True, confianca=.96)
            if caixa and self.conferir_azul_decisao(caixa):
                confirmada = True
                break
            self.aguardar(.2)
        if not confirmada:
            raise RuntimeError('A seleção azul da Decisão dentro de Peças sigilosas não foi confirmada. O cadeado não será acionado.')

        self.etapa('clicar no cadeado ativo')
        regiao_barra = regiao_cadeado_sigilo(painel, self.pg.size(), carregar_referencia('sigilo_cadeado'))
        cadeado = self.localizar_unico('sigilo_cadeado', regiao=regiao_barra)
        self.verificar()
        self.pg.click(int(cadeado.left+cadeado.width//2), int(cadeado.top+cadeado.height//2))
        self.aguardar()
        self.etapa('clicar em Remover sigilo')
        remover = self.localizar_unico('sigilo_remover')
        self.verificar()
        self.pg.click(int(remover.left+remover.width//2), int(remover.top+remover.height//2))
        self.aguardar()

        if self.localizar_unico('sigilo_remover', tempo=1, opcional=True):
            raise RuntimeError('O menu Remover sigilo permaneceu aberto. Confira o SAJ.')
        self.log('Ciclo Peças sigilosas → Decisão → cadeado → Remover sigilo concluído. Seguindo para a liberação normal.')

    def subir_aguardando_ate_topo(self):
        from pywinauto import Desktop
        titulo = self.localizar('aguardando')
        largura, altura = (int(v) for v in self.pg.size())
        x = int(titulo.left+min(90, max(25, titulo.width//2)))
        y = int(titulo.top+titulo.height+12)
        if not (0 <= x < largura and 0 <= y < altura):
            raise RuntimeError('O painel Peças aguardando liberação está fora da área visível.')
        self.verificar()
        controle = Desktop(backend='win32').from_point(x, y)
        alvo = None
        for _ in range(8):
            self.verificar()
            if controle is None:
                break
            r = controle.rectangle()
            classe = controle.class_name().lower()
            eh_lista = 'tree' in classe or 'list' in classe or callable(getattr(controle, 'roots', None))
            if (eh_lista and controle.is_visible() and controle.is_enabled()
                    and r.left >= 0 and r.right <= largura and r.bottom <= altura
                    and titulo.top-5 <= r.top <= titulo.top+titulo.height+16
                    and 90 <= r.right-r.left <= min(500, largura//2)
                    and r.bottom-r.top >= 30
                    and r.left <= x < r.right and r.top <= y < r.bottom):
                alvo = controle
                break
            controle = controle.parent()
        if alvo is None:
            raise RuntimeError('Não foi possível identificar a lista inferior Peças aguardando liberação. Nenhuma caixa foi marcada.')

        for _ in range(3):
            self.verificar()
            if not alvo.is_visible() or not alvo.is_enabled():
                raise RuntimeError('A lista Peças aguardando liberação não está disponível.')
            r = alvo.rectangle()
            painel = (int(r.left), int(r.top), int(r.right), int(r.bottom))
            if not alvo.has_style(0x00200000):
                self.log('Peças aguardando liberação sem barra vertical: todos os documentos já estão visíveis. Seguindo para selecionar todas.')
                return painel
            dados = ler_barra_vertical(alvo.handle)
            if dados is None:
                if not alvo.has_style(0x00200000):
                    self.log('Peças aguardando liberação sem rolagem vertical. A primeira linha está visível.')
                    return painel
                raise RuntimeError('Não foi possível ler a posição da barra de Peças aguardando liberação.')
            if dados['sem_rolagem'] or dados['posicao'] <= dados['minimo']:
                self.log('Topo de Peças aguardando liberação confirmado.')
                return painel
            if 'rect' not in dados:
                raise RuntimeError('A barra de Peças aguardando liberação não forneceu uma posição válida.')
            b = dados['rect']
            if not (r.left+60 <= b[0] < b[2] <= r.right+5
                    and r.top-5 <= b[1] < b[3] <= r.bottom+5):
                raise RuntimeError('A barra localizada não pertence ao painel inferior Peças aguardando liberação.')
            inicio, fim = pontos_arraste(dados, para_topo=True)
            if not all(0 <= px < largura and 0 <= py < altura for px, py in (inicio, fim)):
                raise RuntimeError('A barra inferior está fora da área visível da tela.')
            self.etapa('subir a barra de Peças aguardando liberação até o topo')
            self.pg.moveTo(*inicio, duration=.15)
            self.verificar()
            self.pg.dragTo(*fim, duration=.7, button='left')
            self.aguardar()
        dados = ler_barra_vertical(alvo.handle)
        if dados and (dados['sem_rolagem'] or dados['posicao'] <= dados['minimo']):
            r = alvo.rectangle()
            self.log('Topo de Peças aguardando liberação confirmado.')
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
        raise RuntimeError('A lista Peças aguardando liberação não chegou ao topo. Nenhuma caixa foi marcada.')

    def selecionar_todas(self):
        self.etapa('conferir o topo de Peças aguardando liberação')
        painel = self.subir_aguardando_ate_topo()
        self.etapa('selecionar todas as peças aguardando liberação')
        titulo = self.localizar('aguardando')
        largura, altura = self.pg.size()
        x, y = max(0, titulo.left), titulo.top + titulo.height
        altura_caixa = max(carregar_referencia('caixa_desmarcada').height, carregar_referencia('caixa_marcada').height)
        regiao = normalizar_regiao((painel[0], painel[1], min(350, painel[2]-painel[0]),
                    min(painel[3]-painel[1], max(26, altura_caixa*2))), self.pg.size())
        vazia = self.localizar('caixa_desmarcada', tempo=1, opcional=True, regiao=regiao)
        marcada = self.localizar('caixa_marcada', tempo=1, opcional=True, regiao=regiao)
        caixas = [(c, estado) for c, estado in ((vazia, False), (marcada, True)) if c is not None]
        if not caixas:
            raise RuntimeError('Caixa de selecionar todas não localizada na primeira linha, mesmo com Peças aguardando liberação no topo. Nenhuma peça individual foi marcada.')
        caixa, selecionada = min(caixas, key=lambda item: (item[0].top, item[0].left))
        if not selecionada:
            self.verificar()
            self.pg.click(self.pg.center(caixa))
        self.aguardar()
        rx, ry = max(0, caixa.left - 4), max(0, caixa.top - 4)
        confer = (rx, ry, min(caixa.width + 8, largura-rx), min(caixa.height + 8, altura-ry))
        self.localizar('caixa_marcada', tempo=5, regiao=confer)
        return (x, max(0, titulo.top), min(450, largura-x), min(190, altura-max(0, titulo.top)))

    def processo(self, registro, completo):
        import pyperclip
        self.etapa('preencher número do processo')
        self.clicar_campo_processo()
        self.tecla('home')
        self.verificar()
        self.pg.write(digitos(registro['processo']), interval=.1)
        self.aguardar()
        self.etapa('carregar processo com Tab')
        self.tecla('tab')
        self.etapa('verificar pendências')
        if self.localizar('pendencias', tempo=6, opcional=True):
            self.tecla('enter')
            limite = time.monotonic() + 10
            while self.localizar('pendencias', tempo=.5, opcional=True):
                if time.monotonic() >= limite:
                    raise RuntimeError('Janela de pendências permaneceu aberta.')
                self.aguardar(.25)
        self.etapa('abrir importação de peça')
        self.clicar('importar')
        self.etapa('preencher código ' + registro['codigo'])
        self.clicar('codigo', inferior=True)
        self.verificar()
        self.pg.hotkey('ctrl', 'a')
        pyperclip.copy(registro['codigo'])
        self.pg.hotkey('ctrl', 'v')
        self.aguardar()
        self.tecla('enter')
        self.etapa('importar o PDF correspondente')
        importar_pdf(registro['pdf'], self.verificar, self.aguardar, self.log)
        self.localizar('aguardando')
        self.aguardar()
        if not completo:
            return 'PDF importado — conferir e salvar manualmente'
        self.etapa('salvar a peça importada')
        self.clicar('salvar')
        self.tratar_sigilo()
        regiao = self.selecionar_todas()
        self.etapa('clicar no botão de liberação')
        self.clicar('liberar', pausa=4, regiao=regiao)
        self.etapa('liberar com Alt + L')
        self.verificar()
        limite_alt_l = time.monotonic() + LIMITE_APOS_ALT_L
        self.pg.hotkey('alt', 'l')
        self.aguardar(min(ESPERA_SAJ, max(0, limite_alt_l-time.monotonic())))
        self.etapa('salvar após liberação')
        self.clicar_quando_disponivel('salvar', limite=limite_alt_l)
        self.limpar_para_proximo()
        return 'Fluxo de juntada e liberação executado'

    def executar(self, tarefas, completo, inicio):
        self.log('Início em 5 segundos. Coloque o SAJ em primeiro plano.')
        self.aguardar(5)
        concluidos, falhas = 0, 0
        for indice, (iid, registro) in enumerate(tarefas, 1):
            try:
                self.verificar()
            except ParadaSolicitada:
                self.evento('fase', 'Interrompido pelo usuário')
                self.log('Parada solicitada antes de iniciar o próximo processo.')
                return
            self.fase = 'início'
            resultado = ''
            erro_linha, precisa_recuperar, interrompido = False, False, False
            self.evento('linha', (iid, 'Em execução'))
            self.evento('atual', f'{indice}/{len(tarefas)} • {registro["processo"]}')
            try:
                self.verificar()
                if indice == 1 and inicio:
                    self.etapa('abrir Fluxo de Trabalho')
                    self.clicar('fluxo')
                    self.etapa('abrir Digitalização de Peças')
                    self.clicar('digitalizacao')
                resultado = self.processo(registro, completo)
            except ParadaSolicitada as exc:
                interrompido = True
                resultado = f'Interrompido na etapa {self.fase}: {exc}'
                self.log(resultado)
            except Exception as exc:
                if isinstance(exc, getattr(self.pg, 'FailSafeException', ParadaSolicitada)):
                    interrompido = True
                    resultado = f'Interrompido na etapa {self.fase}: parada de emergência pelo mouse'
                else:
                    erro_linha, precisa_recuperar = True, True
                    motivo = ' '.join(str(exc).split()) or type(exc).__name__
                    resultado = f'Erro na etapa {self.fase}: {motivo}'
                self.log(resultado)
            fase_registro = self.fase
            self.evento('linha', (iid, resultado))
            try:
                gravar_historico(registro, fase_registro, resultado)
            except Exception as exc:
                detalhe = ' '.join(str(exc).split()) or type(exc).__name__
                if not erro_linha and not interrompido:
                    resultado = 'Erro ao gravar histórico após execução: '+detalhe+' | '+resultado
                else:
                    resultado += ' | Falha no histórico: '+detalhe
                erro_linha = True
                self.log('Falha no histórico: '+detalhe)
                self.evento('linha', (iid, resultado))
            self.evento('progresso', indice/len(tarefas)*100)
            if interrompido:
                self.evento('fase', 'Interrompido pelo usuário')
                return
            if erro_linha:
                falhas += 1
            else:
                concluidos += 1
            if precisa_recuperar and indice < len(tarefas):
                self.log('O erro foi registrado em Resultado. Preparando a tela para o próximo processo.')
                try:
                    self.recuperar_apos_erro()
                except ParadaSolicitada:
                    self.evento('fase', 'Interrompido pelo usuário durante a recuperação')
                    self.log('Recuperação interrompida pelo usuário. O resultado com o motivo do erro foi mantido.')
                    return
        self.evento('fase', f'Execução encerrada: {concluidos} concluídos; {falhas} com erro')
        self.log(f'Execução encerrada. {concluidos} concluídos e {falhas} com erro. Confira a coluna Resultado.')


class Aplicativo:
    def __init__(self, root):
        self.root, self.eventos = root, queue.Queue()
        self.parar, self.executando, self.registros = threading.Event(), False, {}
        self.planilha, self.pasta = tk.StringVar(), tk.StringVar()
        self.inicio, self.completo = tk.BooleanVar(value=True), tk.BooleanVar(value=False)
        self.estado = tk.StringVar(value='Aguardando configuração')
        self.atual = tk.StringVar(value='Nenhum processo em execução')
        self.resumo = tk.StringVar(value='Nenhuma planilha carregada')
        root.title(TITULO)
        largura = min(1280, max(850, root.winfo_screenwidth()-80))
        altura = min(820, max(540, root.winfo_screenheight()-100))
        root.geometry(f'{largura}x{altura}+20+20')
        root.minsize(850, 540)
        root.configure(bg=BG)
        self.tema()

        cabecalho = tk.Frame(root, bg=BLUE)
        cabecalho.pack(side='top', fill='x')
        titulo = tk.Label(cabecalho, text=TITULO, bg=BLUE, fg='white', font=('Tahoma', 12, 'bold'), anchor='w', justify='left', padx=12, pady=10)
        titulo.pack(fill='x')
        cabecalho.bind('<Configure>', lambda e: titulo.configure(wraplength=max(300, e.width-30)))
        tk.Label(root, text='  Juntada e Liberação de Peças Processuais', bg='#353535', fg='white', font=('Tahoma', 9, 'bold'), anchor='w', pady=6).pack(side='top', fill='x')

        # Rodapé fixo: não pode ser empurrado para fora da tela.
        rodape = ttk.Frame(root, padding=(10, 7))
        rodape.pack(side='bottom', fill='x')
        ttk.Separator(rodape).pack(fill='x', pady=(0, 6))
        botoes = ttk.Frame(rodape)
        botoes.pack(fill='x')
        ttk.Checkbutton(botoes, text='Salvar, liberar e limpar', variable=self.completo).pack(side='left', padx=(0, 12))
        self.btn_individual = ttk.Button(botoes, text='Iniciar selecionado', command=lambda: self.iniciar(False), style='Primary.TButton')
        self.btn_individual.pack(side='left')
        self.btn_lote = ttk.Button(botoes, text='Iniciar lote', command=lambda: self.iniciar(True), style='Primary.TButton')
        self.btn_lote.pack(side='left', padx=8)
        self.btn_parar = ttk.Button(botoes, text='Parar', command=self.solicitar_parada, style='Stop.TButton')
        self.btn_parar.pack(side='left')
        ttk.Label(rodape, textvariable=self.estado, style='Status.TLabel').pack(anchor='w', pady=(6, 0))
        ttk.Label(rodape, textvariable=self.atual).pack(anchor='w', pady=2)
        self.progresso = ttk.Progressbar(rodape, maximum=100)
        self.progresso.pack(fill='x', pady=4)

        principal = ttk.Frame(root, padding=8)
        principal.pack(fill='both', expand=True)
        principal.columnconfigure(0, minsize=290)
        principal.columnconfigure(1, weight=1)
        principal.rowconfigure(0, weight=1)

        lateral = ttk.Frame(principal, padding=(0, 0, 8, 0))
        lateral.grid(row=0, column=0, sticky='nsew')
        canvas = tk.Canvas(lateral, bg=BG, highlightthickness=0, width=280)
        sy = ttk.Scrollbar(lateral, orient='vertical', command=canvas.yview)
        sy.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        canvas.configure(yscrollcommand=sy.set)
        config = ttk.Frame(canvas, padding=8)
        item = canvas.create_window((0, 0), window=config, anchor='nw')
        config.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfigure(item, width=e.width))
        ttk.Label(config, text='Configurações da execução', font=('Tahoma', 9, 'bold'), foreground=BLUE).pack(anchor='w', pady=(0, 12))
        for texto, var, diretorio in [('Planilha Excel', self.planilha, False), ('Pasta Respostas', self.pasta, True)]:
            ttk.Label(config, text=texto).pack(anchor='w', pady=(5, 3))
            ttk.Entry(config, textvariable=var).pack(fill='x')
            ttk.Button(config, text='Selecionar...', command=lambda v=var, d=diretorio: self.selecionar(v, d)).pack(anchor='e', pady=(4, 10))
        ttk.Separator(config).pack(fill='x', pady=8)
        ttk.Checkbutton(config, text='Começar na tela inicial do SAJ', variable=self.inicio).pack(anchor='w', pady=8)
        ttk.Button(config, text='Conferir planilha e PDFs', command=self.conferir).pack(fill='x', pady=4)
        ttk.Button(config, text='Calibrar imagens', command=self.calibrar).pack(fill='x', pady=4)
        ttk.Separator(config).pack(fill='x', pady=12)
        ttk.Label(config, text='Códigos dos documentos', font=('Tahoma', 9, 'bold')).pack(anchor='w')
        ttk.Label(config, text='IRRISÓRIO: 938\nPARCIAL / INTEGRAL: 939', justify='left').pack(anchor='w', pady=6)
        ttk.Label(config, text='O lote ignora processos com tentativas registradas no histórico.', wraplength=245, justify='left').pack(anchor='w', pady=8)
        def rolar_lateral(e):
            if canvas.yview() == (0.0, 1.0):
                return
            px, py = root.winfo_pointerxy()
            if not (canvas.winfo_rootx() <= px < canvas.winfo_rootx()+canvas.winfo_width() and canvas.winfo_rooty() <= py < canvas.winfo_rooty()+canvas.winfo_height()):
                return
            numero = getattr(e, 'num', None)
            delta = getattr(e, 'delta', 0)
            passo = -3 if numero == 4 or delta > 0 else 3
            canvas.yview_scroll(passo, 'units')
            return 'break'
        root.bind('<MouseWheel>', rolar_lateral)
        root.bind('<Button-4>', rolar_lateral)
        root.bind('<Button-5>', rolar_lateral)

        self.abas = ttk.Notebook(principal)
        self.abas.grid(row=0, column=1, sticky='nsew')
        processos, self.aba_historico = ttk.Frame(self.abas, padding=6), ttk.Frame(self.abas, padding=6)
        self.abas.add(processos, text='Processos')
        self.abas.add(self.aba_historico, text='Histórico da execução')
        processos.columnconfigure(0, weight=1)
        processos.rowconfigure(1, weight=1)
        barra_processos = ttk.Frame(processos)
        barra_processos.grid(row=0, column=0, sticky='ew', pady=(2, 8))
        ttk.Label(barra_processos, textvariable=self.resumo).pack(side='left')
        self.btn_excluir = ttk.Button(barra_processos, text='Excluir selecionadas', command=self.excluir_selecionados)
        self.btn_excluir.pack(side='right', padx=(8, 0))
        tabela_area = ttk.Frame(processos)
        tabela_area.grid(row=1, column=0, sticky='nsew')
        tabela_area.columnconfigure(0, weight=1)
        tabela_area.rowconfigure(0, weight=1)
        self.tabela = ttk.Treeview(tabela_area, columns=('processo', 'status', 'codigo', 'resultado'), show='headings', selectmode='extended')
        for chave, texto, tamanho in [('processo', 'Processo', 240), ('status', 'Status', 100), ('codigo', 'Código', 65), ('resultado', 'Resultado', 360)]:
            self.tabela.heading(chave, text=texto)
            self.tabela.column(chave, width=tamanho, minwidth=60, stretch=chave=='resultado')
        sy = ttk.Scrollbar(tabela_area, orient='vertical', command=self.tabela.yview)
        sx = ttk.Scrollbar(tabela_area, orient='horizontal', command=self.tabela.xview)
        self.tabela.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tabela.grid(row=0, column=0, sticky='nsew')
        sy.grid(row=0, column=1, sticky='ns')
        sx.grid(row=1, column=0, sticky='ew')
        self.tabela.tag_configure('erro', foreground='#b3261e')
        self.tabela.tag_configure('feito', foreground='#176332')
        self.tabela.bind('<Delete>', self.excluir_selecionados)
        self.menu_processos = tk.Menu(root, tearoff=0)
        self.menu_processos.add_command(label='Excluir linhas selecionadas', command=self.excluir_selecionados)
        self.tabela.bind('<Button-3>', self.abrir_menu_processos)
        self.aba_historico.columnconfigure(0, weight=1)
        self.aba_historico.rowconfigure(0, weight=1)
        self.log = tk.Text(self.aba_historico, bg='white', fg=TEXT, relief='sunken', bd=1, font=('Tahoma', 9), state='disabled', wrap='word', padx=8, pady=8)
        scroll_log = ttk.Scrollbar(self.aba_historico, orient='vertical', command=self.log.yview)
        self.log.configure(yscrollcommand=scroll_log.set)
        self.log.grid(row=0, column=0, sticky='nsew')
        scroll_log.grid(row=0, column=1, sticky='ns')
        root.protocol('WM_DELETE_WINDOW', self.fechar)
        root.after(100, self.poll)

    def mostrar_historico(self):
        self.abas.select(self.aba_historico)

    def tema(self):
        s = ttk.Style(self.root)
        s.theme_use('clam')
        s.configure('.', background=BG, foreground=TEXT, font=('Tahoma', 9))
        s.configure('TFrame', background=BG)
        s.configure('TLabel', background=BG, foreground=TEXT)
        s.configure('TEntry', fieldbackground='white', foreground=TEXT, insertcolor=TEXT, padding=4)
        s.configure('TButton', background='#efefef', foreground=TEXT, padding=(10, 6), bordercolor='#9b9b9b')
        s.map('TButton', background=[('active', '#dcecf7')], foreground=[('disabled', '#888888')])
        s.configure('Primary.TButton', background=BLUE, foreground='white', font=('Tahoma', 9, 'bold'))
        s.map('Primary.TButton', background=[('active', '#1475b8')])
        s.configure('Stop.TButton', background='#efefef', foreground='#9b2525')
        s.map('Stop.TButton', background=[('active', '#f5dddd')])
        s.configure('TCheckbutton', background=BG, foreground=TEXT)
        s.map('TCheckbutton', background=[('active', BG)], indicatorbackground=[('selected', BLUE), ('!selected', 'white')])
        s.configure('Treeview', background='white', fieldbackground='white', foreground=TEXT, rowheight=26, bordercolor='#b0b0b0')
        s.configure('Treeview.Heading', background='#efefef', foreground=TEXT, font=('Tahoma', 9), padding=5)
        s.map('Treeview', background=[('selected', '#c8dff0')], foreground=[('selected', '#123c5b')])
        s.configure('TNotebook', background=BG, bordercolor='#b0b0b0')
        s.configure('TNotebook.Tab', background='#e9e9e9', foreground=TEXT, padding=(14, 6))
        s.map('TNotebook.Tab', background=[('selected', '#f5f5f5')], foreground=[('selected', BLUE)])
        s.configure('Horizontal.TProgressbar', background=BLUE, troughcolor='#ececec', borderwidth=0)
        s.configure('Status.TLabel', background=BG, foreground=BLUE, font=('Tahoma', 9, 'bold'))

    def evento(self, tipo, dados):
        self.eventos.put((tipo, dados))

    def poll(self):
        while True:
            try:
                tipo, dados = self.eventos.get_nowait()
            except queue.Empty:
                break
            if tipo == 'log':
                self.log.configure(state='normal')
                self.log.insert('end', time.strftime('%H:%M:%S') + ' > ' + dados + '\n')
                self.log.see('end')
                self.log.configure(state='disabled')
            elif tipo == 'linha':
                iid, texto = dados
                self.tabela.set(iid, 'resultado', texto)
                tags = ('erro',) if texto.startswith(('Erro', 'Interrompido')) else ('feito',) if texto.startswith('Fluxo') else ()
                self.tabela.item(iid, tags=tags)
                self.tabela.see(iid)
            elif tipo == 'fase':
                self.estado.set(dados)
                if dados.startswith('Interrompido'):
                    self.mostrar_historico()
            elif tipo == 'atual':
                self.atual.set(dados)
            elif tipo == 'progresso':
                self.progresso['value'] = dados
            elif tipo == 'fim':
                self.executando = False
                self.btn_excluir.configure(state='normal')
                self.root.deiconify()
        self.root.after(100, self.poll)

    def solicitar_parada(self):
        self.parar.set()
        if self.executando:
            self.estado.set('Parada solicitada...')

    def fechar(self):
        if self.executando:
            self.solicitar_parada()
            messagebox.showinfo('Interrompendo', 'Aguarde o encerramento antes de fechar.')
        else:
            self.root.destroy()

    def selecionar(self, var, diretorio):
        if self.executando:
            return
        caminho = filedialog.askdirectory() if diretorio else filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')])
        if caminho:
            var.set(caminho)

    def conferir(self):
        if self.executando:
            return
        try:
            registros = ler_planilha(self.planilha.get(), self.pasta.get())
            self.tabela.delete(*self.tabela.get_children())
            self.registros.clear()
            for r in registros:
                iid = self.tabela.insert('', 'end', values=(r['processo'], r['status'], r['codigo'], r['erro'] or 'Pronto para teste'), tags=('erro',) if r['erro'] else ())
                self.registros[iid] = r
            self.estado.set('Planilha conferida')
            self.atualizar_resumo()
            self.progresso['value'] = 0
            self.evento('log', f'{len(registros)} linhas conferidas.')
        except Exception as exc:
            messagebox.showerror('Conferência', str(exc))

    def atualizar_resumo(self):
        registros = list(self.registros.values())
        validos = sum(not r['erro'] for r in registros)
        self.resumo.set(f'{len(registros)} processos | {validos} válidos | {len(registros)-validos} com pendências')

    def excluir_selecionados(self, event=None):
        if self.executando:
            return 'break'
        ids = tuple(self.tabela.selection())
        if not ids:
            return 'break'
        self.tabela.delete(*ids)
        for iid in ids:
            self.registros.pop(iid, None)
        self.atualizar_resumo()
        self.estado.set(f'{len(ids)} linha(s) removida(s) da lista')
        self.evento('log', f'{len(ids)} linha(s) excluída(s) apenas da lista do robô.')
        return 'break'

    def abrir_menu_processos(self, event):
        if self.executando:
            return 'break'
        iid = self.tabela.identify_row(event.y)
        if not iid:
            return 'break'
        if iid not in self.tabela.selection():
            self.tabela.selection_set(iid)
        self.tabela.focus(iid)
        try:
            self.menu_processos.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu_processos.grab_release()
        return 'break'

    def calibrar(self):
        if self.executando:
            return
        top = tk.Toplevel(self.root)
        top.title('Calibração e validação • SISBAJUD')
        top.configure(bg=BG)
        top.labels_referencias = {}
        largura = min(900, max(320, top.winfo_screenwidth() - 80))
        altura = min(700, max(300, top.winfo_screenheight() - 120))
        top.geometry(f'{largura}x{altura}+25+25')
        top.minsize(min(500, largura), min(300, altura))
        ttk.Label(top, text='Role a lista pela barra lateral, roda do mouse ou Page Down.', foreground=BLUE).pack(anchor='w', padx=15, pady=10)
        area = ttk.Frame(top)
        area.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        canvas = tk.Canvas(area, bg=BG, highlightthickness=0)
        barra = ttk.Scrollbar(area, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=barra.set)
        barra.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        conteudo = ttk.Frame(canvas)
        janela_conteudo = canvas.create_window((0, 0), window=conteudo, anchor='nw')
        conteudo.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        def redimensionar(event):
            canvas.itemconfigure(janela_conteudo, width=event.width)
            for label in descricoes:
                label.configure(wraplength=max(100, event.width - 170))
            instrucoes.configure(wraplength=max(100, event.width - 30))
        descricoes = []
        canvas.bind('<Configure>', redimensionar)
        instrucoes = ttk.Label(conteudo, justify='left', text='Clique em CAPTURAR e volte ao SAJ em até 4 segundos.\nArraste um recorte pequeno. Esc cancela.\nNOVO Processo: capture SOMENTE a palavra Processo.\nNão inclua setas, ícones, cursor, borda nem a caixa abaixo.\nO robô clicará na caixa abaixo do texto, com ou sem borda azul.\nCódigo: rótulo acima + filtro vazio abaixo.\nCaixas: capture somente o quadradinho, nos dois estados.\nLiberação: capture o ícone habilitado depois de marcar a caixa.\nCabeçalho: capture o texto completo Peças aguardando liberação.\nSigilo: capture o texto Exibir em ordem cronológica da lista superior.\nSigilo: o grupo será aberto com um clique no próprio texto Peças sigilosas.\nNOVO: capture somente a palavra Decisão na primeira linha abaixo do grupo.\nDeixe uma pequena margem à direita e não inclua o ícone, o grupo ou páginas.\nDepois clique na Decisão e capture somente a mesma palavra com o fundo azul.\nCadeado: capture somente o botão ATIVO escuro da barra à direita da lista superior, com sua pequena seta.\nNão capture o cadeado da linha Decisão ou o cadeado cinza desativado.\nCapture a opção Remover sigilo com o menu aberto.\nUse Validar todos os arquivos de imagem antes de iniciar.')
        instrucoes.pack(fill='x', padx=12, pady=15)
        for chave, descricao in ALVOS.items():
            bar = ttk.Frame(conteudo)
            bar.pack(fill='x', padx=12, pady=7)
            bar.columnconfigure(0, weight=1)
            label = ttk.Label(bar, justify='left', text=descricao + (' ✓' if referencia_valida(chave) else ' — capturar'), wraplength=500)
            label.grid(row=0, column=0, sticky='w', padx=(0, 12))
            descricoes.append(label)
            top.labels_referencias[chave] = label
            ttk.Button(bar, text='CAPTURAR', command=lambda k=chave: self.capturar(k, top)).grid(row=0, column=1, sticky='e')
        def validar_todas():
            erros = validar_referencias(set(ALVOS))
            for chave, label in top.labels_referencias.items():
                label.configure(text=ALVOS[chave]+(' ✓' if referencia_valida(chave) else ' — capturar'))
            if erros:
                messagebox.showerror('Imagens pendentes', '\n\n'.join(erros), parent=top)
            else:
                messagebox.showinfo('Validação', 'Todos os arquivos de referência estão válidos.', parent=top)
        ttk.Button(conteudo, text='Validar todos os arquivos de imagem', command=validar_todas).pack(anchor='w', padx=12, pady=15)
        ttk.Label(conteudo, text='Fim da lista de calibração.', foreground=BLUE).pack(anchor='w', padx=12, pady=18)
        def rolar(event):
            if canvas.yview() == (0.0, 1.0):
                return 'break'
            if getattr(event, 'num', None) == 4:
                passo = -3
            elif getattr(event, 'num', None) == 5:
                passo = 3
            else:
                delta = getattr(event, 'delta', 0)
                if not delta:
                    return
                passo = -max(1, round(abs(delta) / 120)) * (1 if delta > 0 else -1) * 3
            canvas.yview_scroll(passo, 'units')
            return 'break'
        # Vincula só à janela de calibração; não interfere no SAJ.
        top.bind('<MouseWheel>', rolar)
        top.bind('<Button-4>', rolar)
        top.bind('<Button-5>', rolar)
        def navegar(event):
            if event.keysym == 'Home':
                canvas.yview_moveto(0)
            elif event.keysym == 'End':
                canvas.yview_moveto(1)
            else:
                canvas.yview_scroll(-1 if event.keysym == 'Prior' else 1, 'pages')
            return 'break'
        for tecla in ('Prior', 'Next', 'Home', 'End'):
            top.bind('<' + tecla + '>', navegar)
        top.focus_set()

    def capturar(self, chave, top):
        try:
            import pyautogui as pg
            from PIL import ImageTk
        except ImportError as exc:
            messagebox.showerror('Dependências', str(exc))
            return
        top.withdraw()
        self.root.withdraw()
        def restaurar():
            self.root.deiconify()
            top.deiconify()
            label = getattr(top, 'labels_referencias', {}).get(chave)
            if label is not None:
                label.configure(text=ALVOS[chave]+(' ✓' if referencia_valida(chave) else ' — capturar'))
        def foto():
            try:
                imagem = pg.screenshot()
                overlay = tk.Toplevel(self.root)
                overlay.overrideredirect(True)
                overlay.geometry(f'{imagem.width}x{imagem.height}+0+0')
                canvas = tk.Canvas(overlay, cursor='crosshair', highlightthickness=0)
                canvas.pack(fill='both', expand=True)
                canvas.foto = ImageTk.PhotoImage(imagem)
                canvas.create_image(0, 0, image=canvas.foto, anchor='nw')
            except Exception as exc:
                restaurar()
                messagebox.showerror('Captura', str(exc))
                return
            inicio, retangulo = [], []
            def fechar(event=None):
                overlay.destroy()
                restaurar()
            def baixo(event):
                inicio[:] = [event.x, event.y]
                retangulo[:] = [canvas.create_rectangle(event.x, event.y, event.x, event.y, outline='red', width=2)]
            def mover(event):
                if retangulo:
                    canvas.coords(retangulo[0], *inicio, event.x, event.y)
            def soltar(event):
                if not inicio:
                    return
                x1, x2 = sorted([inicio[0], event.x])
                y1, y2 = sorted([inicio[1], event.y])
                try:
                    if x2-x1 > 6 and y2-y1 > 6:
                        IMAGENS.mkdir(exist_ok=True)
                        imagem.crop((x1, y1, x2, y2)).save(IMAGENS/(chave+'.png'))
                        self.evento('log', 'Imagem capturada: '+ALVOS[chave])
                except Exception as exc:
                    messagebox.showerror('Captura', str(exc))
                finally:
                    fechar()
            canvas.bind('<ButtonPress-1>', baixo)
            canvas.bind('<B1-Motion>', mover)
            canvas.bind('<ButtonRelease-1>', soltar)
            overlay.bind('<Escape>', fechar)
            overlay.focus_force()
        self.root.after(4000, foto)

    def iniciar(self, lote):
        if self.executando:
            return
        if lote and not self.completo.get():
            messagebox.showinfo('Lote', 'Marque Executar fluxo completo para executar o lote.')
            return
        try:
            import pyautogui, pyperclip, cv2, pywinauto
            anteriores = ler_historico()
            tarefas = []
            ids = self.tabela.get_children() if lote else self.tabela.selection()
            if not lote and len(ids) != 1:
                raise ValueError('Selecione exatamente um processo.')
            for iid in ids:
                r = self.registros[iid]
                if r['erro'] or self.tabela.set(iid, 'resultado') != 'Pronto para teste':
                    if not lote:
                        raise ValueError(r['erro'] or 'Linha já tentada. Confira o SAJ antes de reimportar a planilha.')
                    continue
                if digitos(r['processo']) in anteriores:
                    if lote:
                        self.tabela.set(iid, 'resultado', 'Já consta no histórico — revisar')
                        continue
                    if not messagebox.askyesno('Processo já tentado', 'Este processo já consta no histórico. Confira possíveis duplicidades no SAJ. Deseja tentar novamente?'):
                        return
                tarefas.append((iid, r))
            if not tarefas:
                raise ValueError('Não há processos pendentes elegíveis.')
            necessarios = {'processo_rotulo', 'pendencias', 'importar', 'codigo', 'aguardando'}
            if self.inicio.get():
                necessarios.update({'fluxo', 'digitalizacao'})
            if self.completo.get():
                necessarios.update({'salvar', 'caixa_desmarcada', 'caixa_marcada', 'liberar', 'limpar', 'sigilo_lista', 'sigilo_grupo', 'sigilo_decisao_texto', 'sigilo_decisao_azul', 'sigilo_cadeado', 'sigilo_remover'})
            erros_imagens = validar_referencias(necessarios)
            if erros_imagens:
                raise ValueError('Calibre ou substitua estas imagens antes de iniciar:\n\n'+'\n\n'.join(erros_imagens))
        except Exception as exc:
            messagebox.showerror('Preparação', str(exc))
            return
        self.executando = True
        self.btn_excluir.configure(state='disabled')
        self.parar.clear()
        self.progresso['value'] = 0
        self.estado.set('Preparando execução')
        completo, inicio = self.completo.get(), self.inicio.get()
        self.root.iconify()
        def trabalho():
            try:
                Motor(self.parar, self.evento).executar(tarefas, completo, inicio)
            except Exception as exc:
                self.evento('fase', 'Execução interrompida')
                self.evento('log', str(exc))
            finally:
                self.evento('fim', None)
        threading.Thread(target=trabalho, daemon=True).start()


if __name__ == '__main__':
    janela = tk.Tk()
    Aplicativo(janela)
    janela.mainloop()
