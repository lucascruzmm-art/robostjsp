"""Relógio virtual: não usa desktop nem envia ações ao SAJ."""
import importlib.util
import threading
from collections import namedtuple
from pathlib import Path
from types import SimpleNamespace
from PIL import Image, ImageDraw

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('saj_prazo',ROOT/'saj_previa.py')
app=importlib.util.module_from_spec(spec);spec.loader.exec_module(app)
Box=namedtuple('Box','left top width height')
class Clock:
    def __init__(self):self.now=0.0
    def monotonic(self):return self.now
    def sleep(self,n):self.now+=max(0,float(n))
class FakePG:
    class FailSafeException(Exception):pass
    def __init__(self,clock,ativo):self.clock=clock;self.ativo=ativo;self.cliques=[]
    def failSafeCheck(self):pass
    def size(self):return (1366,768)
    def screenshot(self,region):
        assert all(type(v) is int for v in region)
        return ATIVO.copy() if self.clock.now>=self.ativo else INATIVO.copy()
    def click(self,x,y):self.cliques.append((self.clock.now,x,y))
ATIVO=Image.new('RGB',(36,16),'#dddddd')
ImageDraw.Draw(ATIVO).rectangle((5,4,30,11),fill='black')
INATIVO=Image.new('RGB',(36,16),'#dddddd')
ImageDraw.Draw(INATIVO).rectangle((5,4,30,11),fill='#aaaaaa')
app.carregar_referencia=lambda chave:ATIVO.copy()
original_time=app.time
resultados=[]
def motor_teste(ativo=float('inf')):
    clock=Clock();app.time=clock
    motor=app.Motor.__new__(app.Motor)
    motor.parar=threading.Event();motor.pg=FakePG(clock,ativo)
    motor.fase='salvar após liberação';motor.eventos=[]
    motor.evento=lambda *a:motor.eventos.append(a)
    motor.localizar=lambda *a,**k:Box(10,700,36,16)
    motor.saj_disponivel=lambda caixa:True
    return motor,clock

def caso(nome,executar):
    try:
        executar();resultados.append((nome,True,''));print('OK:',nome,flush=True)
    except Exception as exc:
        resultados.append((nome,False,repr(exc)));print('FALHOU:',nome,repr(exc),flush=True)

def pronto6():
    motor,clock=motor_teste(6)
    motor.aguardar(6)
    motor.clicar_quando_disponivel('salvar',limite=8)
    assert motor.pg.cliques[0][0]==6,motor.pg.cliques

def pronto75():
    motor,clock=motor_teste(7.5)
    motor.aguardar(6)
    motor.clicar_quando_disponivel('salvar',limite=8)
    assert 7.5<=motor.pg.cliques[0][0]<8,motor.pg.cliques

def indisponivel8():
    motor,clock=motor_teste()
    motor.aguardar(6)
    try:motor.clicar_quando_disponivel('salvar',limite=8);raise AssertionError('Faltou timeout')
    except TimeoutError:pass
    assert not motor.pg.cliques
    assert abs(clock.now-8)<1e-8,clock.now

caso('Salvar ativo em 6s: clica em 6s',pronto6)
caso('Salvar ativo em 7.5s: clica antes de 8s',pronto75)
caso('Sempre desabilitado: timeout em 8s sem clique',indisponivel8)

if __name__=='__main__':
    # Casos adicionais são definidos abaixo; nenhum relógio real é alterado.
    pass

def imagem_tardia():
    motor,clock=motor_teste(8.1)
    motor.aguardar(6)
    def localizar(*a,**k):
        clock.now=8.1
        return Box(10,700,36,16)
    motor.localizar=localizar
    try:motor.clicar_quando_disponivel('salvar',limite=8);raise AssertionError('Faltou timeout')
    except TimeoutError:pass
    assert not motor.pg.cliques


def parar_antes():
    motor,clock=motor_teste(6)
    motor.aguardar(6)
    motor.parar.set()
    try:motor.clicar_quando_disponivel('salvar',limite=8);raise AssertionError('Faltou parada')
    except app.ParadaSolicitada:pass
    assert not motor.pg.cliques


def parar_no_prazo():
    motor,clock=motor_teste()
    motor.aguardar(8)
    motor.parar.set()
    try:motor.clicar_quando_disponivel('salvar',limite=8);raise AssertionError('Faltou parada')
    except app.ParadaSolicitada:pass
    assert not motor.pg.cliques


def parar_durante_busca():
    motor,clock=motor_teste(8.1)
    motor.aguardar(6)
    def localizar(*a,**k):
        clock.now=8.1
        motor.parar.set()
        return Box(10,700,36,16)
    motor.localizar=localizar
    try:motor.clicar_quando_disponivel('salvar',limite=8);raise AssertionError('Faltou parada')
    except app.ParadaSolicitada:pass
    assert not motor.pg.cliques


def validacao_tardia():
    motor,clock=motor_teste(7.6)
    motor.aguardar(7.6)
    original=app.conferir_aparencia_botao
    def conferir(*args):clock.sleep(.5);return True
    app.conferir_aparencia_botao=conferir
    try:
        try:motor.clicar_quando_disponivel('salvar',limite=8);raise AssertionError('Faltou timeout')
        except TimeoutError:pass
        assert not motor.pg.cliques
    finally:app.conferir_aparencia_botao=original


def lote_continua():
    motor,clock=motor_teste()
    historico,sequencia=[],[]
    original=app.gravar_historico
    app.gravar_historico=lambda r,f,res:historico.append((r['processo'],f,res))
    def processo(registro,completo):
        sequencia.append(('processo',registro['processo']))
        if registro['processo']=='1':
            prazo=clock.now+8
            motor.aguardar(6)
            motor.etapa('salvar após liberação')
            motor.clicar_quando_disponivel('salvar',limite=prazo)
        return 'Fluxo de juntada e liberação executado'
    def recuperar():
        sequencia.append(('limpar',))
        motor.etapa('limpeza após erro')
    motor.processo=processo
    motor.recuperar_apos_erro=recuperar
    try:
        motor.executar([('linha1',dict(processo='1')),('linha2',dict(processo='2'))],True,False)
        assert sequencia==[('processo','1'),('limpar',),('processo','2')],sequencia
        assert historico[0][1]=='salvar após liberação',historico
        assert historico[0][2].startswith('Erro na etapa salvar após liberação:'),historico
        assert '8 segundos' in historico[0][2]
        assert historico[1][2]=='Fluxo de juntada e liberação executado'
        linhas=[d for t,d in motor.eventos if t=='linha']
        assert ('linha1',historico[0][2]) in linhas
    finally:app.gravar_historico=original

caso('Imagem só retorna após prazo: nenhum clique',imagem_tardia)
caso('Parada manual anterior prevalece',parar_antes)
caso('Parada e deadline simultâneos: parada prevalece',parar_no_prazo)
caso('Parada durante busca que excede prazo prevalece',parar_durante_busca)
caso('Validação de aparência excede prazo: nenhum clique',validacao_tardia)
caso('Timeout registra motivo, limpa e continua segundo processo',lote_continua)
app.time=original_time
print('\nResultado:',sum(ok for _,ok,_ in resultados),'/',len(resultados))
if not all(ok for _,ok,_ in resultados):
    raise SystemExit(1)
