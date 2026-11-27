"""Autoteste local. Não clica no SAJ, não abre processos e não salva capturas."""
import importlib
import hashlib
import json
import os
import platform
import shutil
import struct
import sys
import tempfile
import traceback
import uuid
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ))
from iniciar_portatil import configurar_runtime, dentro

RELATORIO = RAIZ / "logs" / "relatorio_portabilidade.json"
RESULTADO = {
    "data": datetime.now().isoformat(timespec="seconds"),
    "ok": False,
    "escopo": "Autoteste local; sem cliques ou leitura de processos do SAJ; sem capturas persistidas.",
    "testes": [],
}


def rel(caminho):
    try:
        return str(Path(caminho).resolve().relative_to(RAIZ))
    except (ValueError, OSError):
        return "FORA_DO_PACOTE"


def conferir(condicao, mensagem):
    if not condicao:
        raise RuntimeError(mensagem)


def teste(nome, funcao):
    try:
        detalhe = funcao()
        RESULTADO["testes"].append({"nome": nome, "ok": True, "detalhe": detalhe})
        print("OK: " + nome)
        return detalhe
    except Exception:
        erro = traceback.format_exc().replace(str(RAIZ), "<PACOTE>")
        RESULTADO["testes"].append({"nome": nome, "ok": False, "erro": erro})
        print("FALHOU: " + nome)
        return None


def runtime():
    configurar_runtime()
    conferir(platform.python_implementation() == "CPython", "Runtime não é CPython.")
    conferir(sys.version_info[:3] == (3, 12, 14), "A versão esperada é CPython 3.12.14.")
    conferir(struct.calcsize("P") == 8, "O interpretador não é x64.")
    conferir(sys.flags.isolated == 1, "Execute com -I, usando Verificar_portabilidade.bat.")
    conferir(sys.flags.no_user_site == 1, "Site-packages do usuário não está isolado.")
    conferir(dentro(sys.executable, RAIZ / "python"), "Interpretador externo ao pacote.")
    python_no_path, py_no_path = shutil.which("python"), shutil.which("py")
    conferir(python_no_path is None and py_no_path is None, "Há python/py externo no PATH; execute pelo BAT de verificação.")
    for caminho in sys.path:
        conferir(bool(caminho) and dentro(caminho, RAIZ), "sys.path contém caminho externo ao pacote.")
    return {"python": platform.python_version(), "bits": struct.calcsize("P") * 8,
            "executavel": rel(sys.executable), "prefixo": rel(sys.prefix),
            "sys_path": [rel(p) for p in sys.path], "isolado": True,
            "python_no_path": python_no_path, "py_no_path": py_no_path}


def bibliotecas():
    nomes = ["tkinter", "_tkinter", "PIL", "PIL.Image", "PIL.ImageTk", "numpy", "cv2",
             "openpyxl", "et_xmlfile", "six", "pyperclip", "pyautogui", "pyscreeze",
             "pytweening", "pygetwindow", "pyrect", "mouseinfo", "pymsgbox", "pywintypes", "pythoncom",
             "win32api", "win32gui", "win32clipboard", "comtypes", "comtypes.client", "pywinauto"]
    origens = {}
    for nome in nomes:
        modulo = importlib.import_module(nome)
        origem = getattr(modulo, "__file__", None)
        conferir(origem is not None and dentro(origem, RAIZ), "Biblioteca externa ou sem origem local: " + nome)
        origens[nome] = rel(origem)
    return origens


def imagem_cv():
    import cv2
    import numpy as np
    from PIL import Image
    matriz = np.zeros((40, 40, 3), dtype=np.uint8)
    matriz[10:20, 12:23] = [45, 150, 230]
    imagem = Image.fromarray(matriz)
    conferir(imagem.size == (40, 40), "Pillow não criou a imagem de teste.")
    gray = cv2.cvtColor(np.asarray(imagem), cv2.COLOR_RGB2GRAY)
    resposta = cv2.matchTemplate(gray, gray[9:21, 11:24].copy(), cv2.TM_SQDIFF_NORMED)
    _, _, ponto, _ = cv2.minMaxLoc(resposta)
    conferir(ponto == (11, 9), "OpenCV não localizou o padrão sintético.")
    return {"opencv": cv2.__version__, "numpy": np.__version__, "pillow": Image.__version__}


def xlsx():
    from openpyxl import Workbook, load_workbook
    with tempfile.TemporaryDirectory(prefix="autoteste_xlsx_", dir=RELATORIO.parent) as pasta:
        caminho = Path(pasta) / "dados_sinteticos.xlsx"
        livro = Workbook()
        livro.active.append(["teste_local", 123])
        livro.save(caminho)
        livro.close()
        reaberto = load_workbook(caminho, read_only=True, data_only=True)
        try:
            conferir(next(reaberto.active.values) == ("teste_local", 123), "XLSX sintético não confere.")
        finally:
            reaberto.close()
    return "Leitura e escrita de XLSX sintético; arquivo temporário removido."


def win32():
    import ctypes
    import win32api
    import win32gui
    from pywinauto import Desktop
    import comtypes.client
    conferir(win32api.GetCurrentProcessId() > 0, "API Win32 indisponível.")
    conferir(win32gui.GetDesktopWindow() != 0, "Desktop Win32 indisponível.")
    quantidade = len(Desktop(backend="win32").windows())
    shell = comtypes.client.CreateObject("Scripting.Dictionary")
    shell.Add("teste", "local")
    conferir(shell.Exists("teste"), "Teste COM não respondeu.")
    del shell
    return {"win32": True, "comtypes": True, "pywinauto_desktop": True,
            "janelas_enumeradas": quantidade, "titulos_coletados": False}


def interface():
    import tkinter as tk
    from PIL import Image, ImageTk
    import pyautogui
    import saj_previa as app
    conferir(dentro(app.__file__, RAIZ), "Fonte do robô veio de fora do pacote.")
    root = tk.Tk()
    erros = []
    root.report_callback_exception = lambda tipo, valor, tb: erros.append("Erro de callback Tk: " + tipo.__name__)
    try:
        root.title("Teste local de portabilidade")
        img = ImageTk.PhotoImage(Image.new("RGB", (20, 20), "#005a9e"), master=root)
        tk.Label(root, image=img, text="Teste local", compound="left").pack()
        root.update_idletasks()
        root.update()
        captura = pyautogui.screenshot()
        tamanho = captura.size
        conferir(tamanho[0] > 0 and tamanho[1] > 0, "Captura de tela indisponível.")
        del captura
        for widget in root.winfo_children():
            widget.destroy()
        robo = app.Aplicativo(root)
        root.update_idletasks()
        root.update()
        conferir(not getattr(robo, "executando", False), "A UI iniciou execução automaticamente.")
        conferir(not erros, "Erro ao construir a interface do robô.")
        return {"tk": str(root.tk.call("info", "patchlevel")), "image_tk": True,
                "screenshot_em_memoria": True, "screenshot_salvo": False,
                "interface_app": True, "acoes_saj": 0}
    finally:
        root.destroy()


def calibracoes():
    from PIL import Image
    import saj_previa as app
    faltantes, invalidas = [], []
    chaves = list(app.ALVOS)
    pasta = Path(app.IMAGENS)
    arquivos = {}
    if pasta.is_dir():
        for caminho in sorted(pasta.iterdir(), key=lambda p: p.name.casefold()):
            if caminho.is_file() and caminho.suffix.lower() == ".png":
                arquivos[caminho.name] = {"sha256": hashlib.sha256(caminho.read_bytes()).hexdigest(),
                                         "bytes": caminho.stat().st_size}
    for chave in chaves:
        caminho = pasta / (chave + ".png")
        if not caminho.is_file():
            faltantes.append(chave + ".png")
            continue
        try:
            with Image.open(caminho) as imagem:
                imagem.verify()
        except Exception:
            invalidas.append(chave + ".png")
    validador = getattr(app, "validar_referencias", None)
    if callable(validador):
        erros_app = validador(set(chaves))
        conferir(isinstance(erros_app, list), "O validador de referências não retornou uma lista.")
        erros_app = [str(erro).replace(str(RAIZ), "<PACOTE>") for erro in erros_app]
    else:
        erros_app = ["A fonte não oferece validar_referencias; não foi possível validar as regras do robô."]
    # Ausência de referências é reportada como calibração pendente; nunca impede abrir a UI.
    return {"total_referencias": len(chaves), "ausentes": faltantes, "invalidas": invalidas,
            "validador_app_executado": callable(validador), "erros_app": erros_app,
            "arquivos_encontrados": arquivos,
            "fonte_sha256": hashlib.sha256(Path(app.__file__).read_bytes()).hexdigest(),
            "prontas": not faltantes and not invalidas and not erros_app,
            "escopo": "Valida PNGs e regras de referências da fonte; não comprova reconhecimento no SAJ."}


def clipboard():
    if "--clipboard" not in sys.argv[1:]:
        return {"testado": False, "motivo": "Opcional: execute com --clipboard para testar texto e restaurar.", "alterado": False}
    import win32clipboard as clip
    import pyperclip
    clip.OpenClipboard()
    try:
        formatos, formato = [], 0
        while True:
            formato = clip.EnumClipboardFormats(formato)
            if not formato:
                break
            formatos.append(formato)
    finally:
        clip.CloseClipboard()
    if any(f not in (1, 7, 13, 16) for f in formatos):
        return {"testado": False, "motivo": "Clipboard com formatos não textuais preservado sem alterações.", "alterado": False}
    original = pyperclip.paste() if formatos else None
    marcador = "autoteste_portatil_" + uuid.uuid4().hex
    try:
        pyperclip.copy(marcador)
        conferir(pyperclip.paste() == marcador, "Clipboard de texto não respondeu.")
    finally:
        if original is None:
            clip.OpenClipboard()
            try:
                clip.EmptyClipboard()
            finally:
                clip.CloseClipboard()
        else:
            pyperclip.copy(original)
    conferir(original is None or pyperclip.paste() == original, "Não foi possível conferir restauração do clipboard.")
    return {"testado": True, "restaurado": True, "conteudo_coletado": False}


def main():
    RELATORIO.parent.mkdir(parents=True, exist_ok=True)
    if teste("Runtime isolado do ZIP", runtime) is not None:
        teste("Bibliotecas com origem no ZIP", bibliotecas)
        teste("Pillow, NumPy e OpenCV", imagem_cv)
        teste("XLSX sintético", xlsx)
        teste("Win32, COM e Desktop", win32)
        teste("Tk, ImageTk, screenshot e interface do robô", interface)
        teste("Referências de calibração", calibracoes)
        teste("Clipboard opcional", clipboard)
    RESULTADO["ok"] = all(t["ok"] for t in RESULTADO["testes"])
    RESULTADO["calibracoes_prontas"] = next((t.get("detalhe", {}).get("prontas") for t in RESULTADO["testes"] if t["nome"] == "Referências de calibração" and t["ok"]), False)
    RESULTADO["pronto_para_saj"] = RESULTADO["ok"] and RESULTADO["calibracoes_prontas"]
    RELATORIO.write_text(json.dumps(RESULTADO, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\n" + ("AUTOTESTE CONCLUIDO." if RESULTADO["ok"] else "AUTOTESTE COM FALHAS."))
    print("Relatorio: logs/relatorio_portabilidade.json")
    if not RESULTADO["calibracoes_prontas"]:
        print("Calibracoes pendentes: o robo pode abrir, mas sera necessario capturar as referencias ausentes/invalidas.")
    return 0 if RESULTADO["ok"] else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:
        RESULTADO["erro_fatal"] = traceback.format_exc().replace(str(RAIZ), "<PACOTE>")
        try:
            RELATORIO.parent.mkdir(parents=True, exist_ok=True)
            RELATORIO.write_text(json.dumps(RESULTADO, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass
        print("Falha no autoteste. Consulte a pasta logs.")
        raise SystemExit(1)
