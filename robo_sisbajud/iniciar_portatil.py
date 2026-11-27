"""Inicia a fonte do robô sem alterar seu código ou suas calibrações."""
import ctypes
import os
import struct
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
PYTHON = RAIZ / "python"
DLL_HANDLES = []


def dentro(caminho, pasta):
    try:
        Path(caminho).resolve().relative_to(Path(pasta).resolve())
        return True
    except (ValueError, OSError):
        return False


def configurar_runtime():
    if os.name != "nt":
        raise RuntimeError("Este pacote foi preparado para Windows x64.")
    if struct.calcsize("P") != 8 or not dentro(sys.executable, PYTHON):
        raise RuntimeError("Use o Python x64 incluído na subpasta python.")
    # -I ignora PYTHONPATH e o diretório do script: insere apenas recursos locais.
    if str(RAIZ) not in sys.path:
        sys.path.insert(0, str(RAIZ))
    site = PYTHON / "Lib" / "site-packages"
    if not site.is_dir():
        raise RuntimeError("Bibliotecas portáteis ausentes em python/Lib/site-packages.")
    for pasta in (site, site / "win32", site / "win32" / "lib", site / "pywin32_system32"):
        if pasta.is_dir() and str(pasta) not in sys.path:
            sys.path.append(str(pasta))
    for pasta in (PYTHON, PYTHON / "DLLs", site / "pywin32_system32", site / "win32"):
        if pasta.is_dir():
            DLL_HANDLES.append(os.add_dll_directory(str(pasta)))
    tcl = PYTHON / "tcl" / "tcl8.6"
    tk = PYTHON / "tcl" / "tk8.6"
    if not (tcl / "init.tcl").is_file() or not (tk / "tk.tcl").is_file():
        raise RuntimeError("Tcl/Tk portátil incompleto. Extraia novamente o pacote inteiro.")
    os.environ["TCL_LIBRARY"] = str(tcl)
    os.environ["TK_LIBRARY"] = str(tk)
    # Os arquivos locais ficam ao lado da fonte, independentemente de quem abriu o .bat.
    os.chdir(RAIZ)
    return site


def registrar_erro(texto):
    registro = datetime.now().isoformat(timespec="seconds") + "\n" + texto + "\n\n"
    for pasta in (RAIZ / "logs", Path(tempfile.gettempdir()) / "RoboSisbajud"):
        try:
            pasta.mkdir(parents=True, exist_ok=True)
            destino = pasta / "erro_inicializacao.txt"
            with destino.open("a", encoding="utf-8") as arquivo:
                arquivo.write(registro)
            return destino
        except OSError:
            pass
    return None


def avisar_erro(texto):
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        user32.MessageBoxW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint]
        user32.MessageBoxW.restype = ctypes.c_int
        user32.MessageBoxW(None, texto, "Robô Sisbajud — falha na inicialização", 0x10)
    except Exception:
        if sys.stderr is not None:
            print(texto, file=sys.stderr)


def main():
    teste = "--teste" in sys.argv[1:]
    try:
        configurar_runtime()
        import tkinter as tk
        import saj_previa as app
        root = tk.Tk()

        def erro_callback(tipo, valor, traco):
            caminho = registrar_erro("".join(traceback.format_exception(tipo, valor, traco)))
            if not teste:
                avisar_erro("Ocorreu um erro na interface.\n" + ("Detalhes em:\n" + str(caminho) if caminho else "Não foi possível gravar o diagnóstico."))

        root.report_callback_exception = erro_callback
        # Não exige as imagens para abrir: a calibração pode ser feita pela própria UI.
        app.Aplicativo(root)
        if teste:
            root.after(400, root.destroy)
        root.mainloop()
        return 0
    except Exception:
        caminho = registrar_erro(traceback.format_exc())
        if not teste:
            avisar_erro("Não foi possível iniciar o robô.\n" + ("Detalhes em:\n" + str(caminho) if caminho else "Não foi possível gravar o diagnóstico."))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
