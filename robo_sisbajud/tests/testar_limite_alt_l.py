import importlib.util
import sys
import tempfile
import threading
from pathlib import Path
from types import SimpleNamespace

raiz = Path(__file__).resolve().parent
fonte = next(p for p in (raiz/'saj_matrix.py', raiz.parent/'saj_previa.py') if p.is_file())
spec = importlib.util.spec_from_file_location('saj_teste', fonte)
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)
sys.modules['pyperclip'] = SimpleNamespace(copy=lambda texto: None)


class Fake(app.Motor):
    def __init__(self, parar_apos_alt=False):
        self.parar = threading.Event()
        self.eventos = []
        self.acoes = []
        self.fase = ''
        self.parar_apos_alt = parar_apos_alt
        self.evento = lambda *args: self.eventos.append(args)
        self.pg = SimpleNamespace(
            failSafeCheck=lambda: None,
            FailSafeException=type('Emergency', (Exception,), {}),
            hotkey=lambda *keys: self.acoes.append(('hotkey', keys)),
            write=lambda *a, **kw: self.acoes.append(('digitar', a[0])))

    def aguardar(self, segundos=3):
        self.verificar()
        self.acoes.append(('aguardar', segundos))
        if self.parar_apos_alt and segundos == 8:
            self.parar.set()
        self.verificar()

    def clicar_campo_processo(self):
        self.acoes.append(('campo_processo',))

    def tecla(self, tecla):
        self.acoes.append(('tecla', tecla))

    def localizar(self, chave, **kwargs):
        return None if chave == 'pendencias' else True

    def clicar(self, chave, **kwargs):
        self.acoes.append(('clicar', chave))

    def tratar_sigilo(self):
        self.acoes.append(('sigilo',))

    def selecionar_todas(self):
        self.acoes.append(('selecionar_todas',))
        return (0, 0, 300, 100)

    def limpar_para_proximo(self, recuperar=False):
        self.verificar()
        self.acoes.append(('limpar', recuperar))


app.importar_pdf = lambda caminho, verificar, aguardar, registrar: verificar()
# Identificadores artificiais inválidos: nenhum número de processo real é usado.
registro = dict(processo='00000000000000000000', status='IRRISORIO', codigo='938', pdf='sintetico.pdf')
motor = Fake()
assert motor.processo(registro, True).startswith('Fluxo')
pos = motor.acoes.index(('hotkey', ('alt', 'l')))
assert motor.acoes[pos+1:] == [('aguardar', 8), ('limpar', False)], motor.acoes[pos:]
assert motor.acoes.count(('clicar', 'salvar')) == 1
assert motor.acoes.index(('clicar', 'salvar')) < pos
print('OK: após Alt+L, somente espera de 8s e Limpar; Salvar permanece antes da liberação.')

motor = Fake(True)
try:
    motor.processo(registro, True)
    raise AssertionError('Parar não interrompeu')
except app.ParadaSolicitada:
    pass
assert not any(a[0] == 'limpar' for a in motor.acoes)
print('OK: Parar durante os 8s interrompe antes de Limpar.')

motor = Fake()
assert motor.processo(registro, False).startswith('PDF importado')
assert ('hotkey', ('alt', 'l')) not in motor.acoes
assert not any(a[0] == 'limpar' for a in motor.acoes)
print('OK: importação sem fluxo completo preservada.')

with tempfile.TemporaryDirectory() as pasta:
    app.HISTORICO = Path(pasta)/'historico.csv'
    motor = Fake()
    motor.executar([('1', registro), ('2', {**registro, 'processo':'00000000000000000001'})], True, False)
    assert motor.acoes.count(('hotkey', ('alt', 'l'))) == 2
    assert motor.acoes.count(('aguardar', 8)) == 2
    assert motor.acoes.count(('limpar', False)) == 2
    assert motor.acoes.count(('clicar', 'salvar')) == 2
    assert not any(k == 'linha' and 'Erro' in v[1] for k, v in motor.eventos)
print('OK: lote de dois processos segue após gravação automática e limpeza.')
