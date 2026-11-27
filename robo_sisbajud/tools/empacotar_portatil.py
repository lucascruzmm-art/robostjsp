"""Monta ZIP com runtime Windows x64 já provisionado. Não baixa componentes."""
import argparse
import hashlib
import json
import shutil
import struct
import tempfile
import zipfile
from email.parser import Parser
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
NOME = 'Robo_Sisbajud_Portatil'


def sha(caminho):
    with Path(caminho).open('rb') as arquivo:
        return hashlib.file_digest(arquivo, 'sha256').hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--runtime', required=True, type=Path,
                        help='Árvore Python Windows completa, já com Tcl/Tk e bibliotecas.')
    parser.add_argument('--saida', type=Path, default=RAIZ / 'dist')
    args = parser.parse_args()
    runtime = args.runtime.resolve()
    obrigatorios = ['python.exe', 'pythonw.exe', 'python312.dll', 'DLLs/_tkinter.pyd',
                    'DLLs/tcl86t.dll', 'DLLs/tk86t.dll', 'tcl/tcl8.6/init.tcl',
                    'tcl/tk8.6/tk.tcl', 'Lib/site-packages/pywin32_system32/pythoncom312.dll',
                    'Lib/site-packages/pywin32_system32/pywintypes312.dll', 'msvcp140.dll']
    for nome in obrigatorios:
        if not (runtime / nome).is_file():
            raise RuntimeError('Runtime incompleto: ' + nome)
    with (runtime / 'python.exe').open('rb') as arquivo:
        if arquivo.read(2) != b'MZ':
            raise RuntimeError('python.exe não é um executável Windows.')
        arquivo.seek(0x3c); offset = struct.unpack('<I', arquivo.read(4))[0]
        arquivo.seek(offset)
        if arquivo.read(4) != b'PE\0\0' or struct.unpack('<H', arquivo.read(2))[0] != 0x8664:
            raise RuntimeError('O runtime precisa ser Windows x64.')
    instalado = {}
    for meta in (runtime / 'Lib/site-packages').glob('*.dist-info/METADATA'):
        dados = Parser().parsestr(meta.read_text(encoding='utf-8'))
        instalado[dados['Name'].lower().replace('_', '-')] = dados['Version']
    for linha in (RAIZ / 'requirements.txt').read_text().splitlines():
        if not linha.strip() or linha.lstrip().startswith('#'):
            continue
        nome, versao = linha.split('==', 1)
        if instalado.get(nome.lower().replace('_', '-')) != versao:
            raise RuntimeError('Dependência ausente ou diferente da versão fixada: ' + linha)
    manifesto_original = json.loads((RAIZ / 'docs/distribuicao/MANIFESTO.json').read_text(encoding='utf-8'))
    for nome, esperado in manifesto_original['imagens_recebidas'].items():
        if sha(RAIZ / 'imagens_saj' / nome) != esperado:
            raise RuntimeError('Calibração difere do snapshot validado: ' + nome)
    args.saida.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='empacotar_', dir=args.saida) as temporario:
        pacote = Path(temporario) / NOME
        pacote.mkdir()
        for nome in ['saj_previa.py', 'Iniciar.bat', 'iniciar_portatil.py',
                     'Verificar_portabilidade.bat', 'verificar_portabilidade.py']:
            shutil.copy2(RAIZ / nome, pacote / nome)
        shutil.copytree(RAIZ / 'imagens_saj', pacote / 'imagens_saj')
        shutil.copytree(runtime, pacote / 'python', ignore=shutil.ignore_patterns('__pycache__', '*.pyc', '*.pyo'))
        shutil.copy2(RAIZ / 'requirements.txt', pacote / 'VERSOES_BIBLIOTECAS.txt')
        shutil.copy2(RAIZ / 'README.md', pacote / 'LEIA-ME.md')
        manifesto = {'tipo': 'Reconstrução local; executar Verificar_portabilidade.bat antes de compartilhar.',
                     'fonte_sha256': sha(pacote / 'saj_previa.py'),
                     'imagens': {p.name: sha(p) for p in sorted((pacote / 'imagens_saj').glob('*.png'))}}
        (pacote / 'MANIFESTO_RECONSTRUCAO.json').write_text(json.dumps(manifesto, ensure_ascii=False, indent=2), encoding='utf-8')
        somas = [sha(p) + '  ' + p.relative_to(pacote).as_posix() for p in sorted(pacote.rglob('*')) if p.is_file()]
        (pacote / 'SHA256SUMS.txt').write_text('\n'.join(somas) + '\n', encoding='utf-8')
        destino = args.saida / (NOME + '_Windows_x64.zip')
        parcial = destino.with_suffix('.zip.tmp')
        with zipfile.ZipFile(parcial, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            for p in sorted(pacote.rglob('*')):
                if p.is_file() and '__pycache__' not in p.parts and p.suffix not in ('.pyc', '.pyo'):
                    z.write(p, Path(NOME) / p.relative_to(pacote))
            z.writestr(NOME + '/logs/', b'')
        with zipfile.ZipFile(parcial) as z:
            if z.testzip():
                raise RuntimeError('Falha CRC no ZIP reconstruído.')
        parcial.replace(destino)
        destino.with_suffix('.zip.sha256').write_text(sha(destino) + '  ' + destino.name + '\n', encoding='utf-8')
    print('ZIP reconstruído:', destino)
    print('Esta reconstrução exige novo teste: não herda a validação do ZIP da Release.')


if __name__ == '__main__':
    main()
