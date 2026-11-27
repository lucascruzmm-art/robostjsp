"""Testes sem desktop/SAJ; não altera a fonte da beta."""
import csv
import importlib.util
import itertools
from pathlib import Path
import sys
import tempfile
import threading
import types
import unittest
from unittest.mock import patch
from openpyxl import Workbook

ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('beta_em_teste',next(x for x in (ROOT/'saj_beta.py',ROOT.parent/'saj_beta.py') if x.is_file()))
beta=importlib.util.module_from_spec(spec);spec.loader.exec_module(beta)
COMBINACOES=[tuple(a for a,b in zip(beta.ACOES,bits) if b) for bits in itertools.product((False,True),repeat=3) if any(bits)]
P1='0000001-00.0000.0.00.0000'
P2='0000002-00.0000.0.00.0000'

class PGFake:
    class FailSafeException(Exception):pass
    def __init__(self,m):self.m=m
    def failSafeCheck(self):pass
    def write(self,texto,interval):self.m.passos.append(('digitar',texto,interval))
    def press(self,tecla):self.m.passos.append(('tecla',tecla))
    def hotkey(self,*teclas):self.m.passos.append(('atalho',tuple(teclas)))

class MotorFake(beta.Motor):
    def __init__(self):
        self.parar=threading.Event();self.fase='preparação';self.pg=PGFake(self)
        self.eventos=[];self.passos=[];self.acao_atual=''
        self.tentadas=[];self.concluidas=[];self.resultados_acoes=[]
        self.corrente=None;self.falha_sigilo=None;self.sem_sigilo=False;self.sem_pecas=False
    def evento(self,tipo,dados):self.eventos.append((tipo,dados))
    def aguardar(self,segundos=beta.INTERVALO):
        self.verificar();self.passos.append(('aguardar',segundos))
    def carregar_processo(self,r):
        self.verificar();self.corrente=r['processo'];self.passos.append(('carregar',self.corrente))
    def clicar(self,chave,inferior=False,pausa=beta.INTERVALO,regiao=None):
        self.verificar();self.passos.append(('clicar',chave,self.corrente));self.aguardar(pausa)
    def localizar(self,chave,**kwargs):
        self.verificar();self.passos.append(('localizar',chave));return object()
    def tratar_sigilo(self):
        self.verificar();self.passos.append(('sigilo',self.corrente));self.etapa('retirar sigilo da Decisão')
        if self.falha_sigilo:self.falha_sigilo(self)
        return not self.sem_sigilo
    def selecionar_todas(self):
        self.verificar();self.passos.append(('selecionar_todas',self.corrente))
        return None if self.sem_pecas else (0,0,250,120)
    def limpar_para_proximo(self,recuperar=False):
        self.verificar();self.passos.append(('limpar',self.corrente,recuperar))
    def recuperar_apos_erro(self):
        self.verificar();self.passos.append(('recuperar',self.corrente));self.etapa('recuperação da tela')

class Var:
    def __init__(self,v):self.valor=v
    def get(self):return self.valor
    def set(self,v):self.valor=v

class TesteBeta(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup);self.pasta=Path(self.temp.name)
        for nome,valor in [('HISTORICO',self.pasta/'historico_saj_beta.csv'),('HISTORICO_ANTERIOR',self.pasta/'historico_saj.csv')]:
            p=patch.object(beta,nome,valor);p.start();self.addCleanup(p.stop)
        self.motor=MotorFake()
        mod=types.ModuleType('pyperclip');mod.copy=lambda t:self.motor.passos.append(('clipboard',t))
        p=patch.dict(sys.modules,{'pyperclip':mod});p.start();self.addCleanup(p.stop)
        p=patch.object(beta,'importar_pdf',self.importar_fake);p.start();self.addCleanup(p.stop)
    def importar_fake(self,caminho,verificar,aguardar,registrar):
        verificar();self.assertTrue(Path(caminho).is_file())
        rows=self.historico();self.assertEqual((rows[-1]['acao'],rows[-1]['estado']),('juntada','iniciada'))
        self.motor.passos.append(('pdf',str(caminho),self.motor.corrente))
    def registro(self,processo=P1):
        pdf=self.pasta/(processo+'.pdf');pdf.write_bytes(b'%PDF-1.4\nfixture\n')
        return dict(processo=processo,status='IRRISORIO',codigo='938',pdf=str(pdf),erro='',linha=2)
    def historico(self):
        if not beta.HISTORICO.exists():return []
        with beta.HISTORICO.open(encoding='utf-8-sig',newline='') as f:return list(csv.DictReader(f,delimiter=';'))
    def workbook(self,headers,linhas,nome='base.xlsx'):
        p=self.pasta/nome;w=Workbook();s=w.active;s.append(headers)
        for r in linhas:s.append(r)
        w.save(p);w.close();return p

    def test_sete_combinacoes_ordem_e_modulos_desmarcados(self):
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                if beta.HISTORICO.exists():beta.HISTORICO.unlink()
                self.motor=MotorFake();res=self.motor.processo(self.registro(),acoes);p=self.motor.passos
                self.assertTrue(res.startswith('Concluído:'));self.assertEqual(tuple(self.motor.concluidas),acoes)
                pdf=[i for i,a in enumerate(p) if a[0]=='pdf']
                sig=[i for i,a in enumerate(p) if a[0]=='sigilo']
                lib=[i for i,a in enumerate(p) if a[0]=='selecionar_todas']
                sal=[i for i,a in enumerate(p) if a[:2]==('clicar','salvar')]
                self.assertEqual(len(pdf),int('juntada' in acoes));self.assertEqual(len(sal),len(pdf))
                self.assertEqual(len(sig),int('sigilo' in acoes));self.assertEqual(len(lib),int('liberacao' in acoes))
                self.assertEqual(sum(a==('atalho',('alt','l')) for a in p),int('liberacao' in acoes))
                self.assertEqual(sum(a==('aguardar',8) for a in p),int('liberacao' in acoes))
                if pdf:
                    self.assertLess(pdf[0],sal[0])
                    for i in sig+lib:self.assertLess(sal[0],i)
                if sig and lib:self.assertLess(sig[0],lib[0])
                self.assertEqual(sum(a[0]=='limpar' for a in p),1);self.assertEqual(p[-1][0],'limpar')
                self.assertEqual([(r['acao'],r['estado']) for r in self.historico()],[(a,e) for a in acoes for e in ('iniciada','concluida')])

    def test_zero_acoes_sem_efeitos(self):
        with self.assertRaises(ValueError):self.motor.processo(self.registro(),())
        self.assertEqual(self.motor.passos,[]);self.assertEqual(self.historico(),[])
    def test_plano_imutavel(self):
        escolhas=['liberacao','juntada'];plano=beta.normalizar_acoes(escolhas);escolhas.clear()
        self.assertEqual(plano,('juntada','liberacao'));self.assertIsInstance(plano,tuple)
        for invalido in ((),None,('outra',)):
            with self.assertRaises(ValueError):beta.normalizar_acoes(invalido)
    def test_sem_juntada_nao_exige_status_pasta_pdf(self):
        fonte=self.workbook(['Processo'],[[P1]])
        for acoes in [('sigilo',),('liberacao',),('sigilo','liberacao')]:
            r=beta.ler_planilha(fonte,self.pasta/'nao_existe',acoes)[0]
            self.assertEqual((r['erro'],r['pdf'],r['codigo']),('','',''))
            self.motor.processo(r,acoes)
        self.assertFalse(any(p[0]=='pdf' for p in self.motor.passos))
    def test_sem_juntada_ignora_status_invalido(self):
        fonte=self.workbook(['Processo','status'],[[P1,'inválido']])
        self.assertEqual(beta.ler_planilha(fonte,'',('sigilo',))[0]['erro'],'')
    def test_juntada_exige_pasta_status_pdf(self):
        fonte=self.workbook(['Processo'],[[P1]])
        with self.assertRaises(ValueError):beta.ler_planilha(fonte,'',('juntada',))
        with self.assertRaises(ValueError):beta.ler_planilha(fonte,self.pasta,('juntada',))
        fonte=self.workbook(['Processo','status'],[[P1,'IRRISÓRIO']],'com_status.xlsx')
        self.assertIn('PDF ausente',beta.ler_planilha(fonte,self.pasta,('juntada',))[0]['erro'])
        self.registro();self.assertEqual(beta.ler_planilha(fonte,self.pasta,('juntada',))[0]['erro'],'')
    def test_codigos(self):
        dados=[(P1,'IRRISÓRIO'),(P2,'PARCIAL'),('0000003-00.0000.0.00.0000','INTEGRAL')]
        for numero,_ in dados:self.registro(numero)
        fonte=self.workbook(['Processo','status'],dados)
        self.assertEqual([r['codigo'] for r in beta.ler_planilha(fonte,self.pasta,('juntada',))],['938','939','939'])
    def test_processo_precisa_texto_sem_duplicidade(self):
        fonte=self.workbook(['Processo'],[[P1],[P1],[15000341420268260090]])
        rs=beta.ler_planilha(fonte,'',('liberacao',))
        self.assertEqual(rs[0]['erro'],'');self.assertIn('repetido',rs[1]['erro']);self.assertIn('Texto',rs[2]['erro'])
    def test_imagens_modulares(self):
        for acoes in COMBINACOES:
            refs=beta.imagens_necessarias(acoes)
            self.assertTrue({'processo_rotulo','pendencias','limpar'}<=refs)
            self.assertEqual('importar' in refs,'juntada' in acoes);self.assertEqual('salvar' in refs,'juntada' in acoes)
            self.assertEqual('sigilo_cadeado' in refs,'sigilo' in acoes);self.assertEqual('liberar' in refs,'liberacao' in acoes)
            self.assertFalse({'fluxo','digitalizacao'}&refs)
            self.assertTrue({'fluxo','digitalizacao'}<=beta.imagens_necessarias(acoes,True))
    def test_juntada_journalizada_antes_sigilo(self):
        def conferir(m):self.assertEqual([r['estado'] for r in self.historico() if r['acao']=='juntada'],['iniciada','concluida'])
        self.motor.falha_sigilo=conferir;self.motor.processo(self.registro(),beta.ACOES)
    def test_sem_alvo_sigilo_nao_libera(self):
        self.motor.sem_sigilo=True;res=self.motor.processo(self.registro(),('sigilo',))
        self.assertIn('Sem Decisão sigilosa',res);self.assertFalse(any(p[0]=='selecionar_todas' for p in self.motor.passos))
        self.assertEqual(beta.ler_historico(),set())
    def test_sem_pecas_nao_envia_alt_l_nem_espera8(self):
        self.motor.sem_pecas=True;res=self.motor.processo(self.registro(),('liberacao',))
        self.assertIn('Sem peças',res);self.assertNotIn(('atalho',('alt','l')),self.motor.passos)
        self.assertNotIn(('aguardar',8),self.motor.passos);self.assertEqual(self.motor.passos[-1][0],'limpar')
    def test_historico_protege_apenas_juntada(self):
        r=self.registro()
        for a in ('sigilo','liberacao'):beta.gravar_historico(r,'fase','feito',(a,),a,'concluida')
        self.assertEqual(beta.ler_historico(),set())
        beta.gravar_historico(r,'fase','tentativa',('juntada',),'juntada','iniciada')
        self.assertEqual(beta.ler_historico(),{beta.digitos(P1)})
    def test_legado_preservado(self):
        beta.HISTORICO_ANTERIOR.write_text('data;processo;resultado\n2026;'+P2+';Interrompido\n',encoding='utf-8-sig')
        antes=beta.HISTORICO_ANTERIOR.read_bytes();self.assertEqual(beta.ler_historico(),{beta.digitos(P2)})
        self.motor.processo(self.registro(),('liberacao',));self.assertEqual(beta.HISTORICO_ANTERIOR.read_bytes(),antes)
        self.assertNotEqual(beta.HISTORICO,beta.HISTORICO_ANTERIOR)
    def test_falha_sigilo_preserva_juntada_recupera_segue(self):
        def falhar(m):
            if m.corrente==P1:raise RuntimeError('erro simulado no sigilo')
        self.motor.falha_sigilo=falhar
        self.motor.executar([('1',self.registro()),('2',self.registro(P2))],beta.ACOES,False)
        primeira=[r for r in self.historico() if r['processo']==P1]
        self.assertTrue(any(r['acao']=='juntada' and r['estado']=='concluida' for r in primeira))
        self.assertFalse(any(r['acao']=='liberacao' for r in primeira))
        self.assertIn('Ações concluídas: Juntar PDF',primeira[-1]['resultado'])
        self.assertIn('Não concluídas: Retirar sigilo da Decisão, Liberar peças',primeira[-1]['resultado'])
        p=self.motor.passos;rec=next(i for i,a in enumerate(p) if a[0]=='recuperar')
        prox=next(i for i,a in enumerate(p) if a==('carregar',P2));self.assertLess(rec,prox)
        self.assertEqual(sum(a[0]=='pdf' and a[2]==P1 for a in p),1)
    def verificar_parada(self,bruta):
        def parar(m):raise (m.pg.FailSafeException('mouse') if bruta else beta.ParadaSolicitada('usuário'))
        self.motor.falha_sigilo=parar
        self.motor.executar([('1',self.registro()),('2',self.registro(P2))],beta.ACOES,False)
        self.assertFalse(any(p[0] in ('limpar','recuperar') for p in self.motor.passos))
        self.assertNotIn(('carregar',P2),self.motor.passos);self.assertEqual(self.historico()[-1]['estado'],'interrompida')
    def test_parada_manual(self):self.verificar_parada(False)
    def test_failsafe_bruto(self):self.verificar_parada(True)
    def test_journal_falha_antes_impede_pdf(self):
        original=beta.gravar_historico
        def falhar(r,f,res,acoes=beta.ACOES,acao='',estado='resumo'):
            if acao=='juntada' and estado=='iniciada':raise OSError('disco sem escrita')
            return original(r,f,res,acoes,acao,estado)
        with patch.object(beta,'gravar_historico',falhar):self.motor.executar([('1',self.registro())],('juntada',),False)
        self.assertFalse(any(p[0]=='pdf' for p in self.motor.passos))
    def test_journal_falha_depois_nao_refaz_pdf(self):
        original=beta.gravar_historico
        def falhar(r,f,res,acoes=beta.ACOES,acao='',estado='resumo'):
            if acao=='juntada' and estado=='concluida':raise OSError('falha no journal final')
            return original(r,f,res,acoes,acao,estado)
        with patch.object(beta,'gravar_historico',falhar):self.motor.executar([('1',self.registro())],beta.ACOES,False)
        self.assertEqual(sum(p[0]=='pdf' for p in self.motor.passos),1)
        self.assertFalse(any(p[0] in ('sigilo','selecionar_todas') for p in self.motor.passos))
        self.assertIn(beta.digitos(P1),beta.ler_historico())
    def test_modo_alterado_bloqueia_inicio(self):
        fonte=self.workbook(['Processo'],[[P1]]);ui=beta.Aplicativo.__new__(beta.Aplicativo)
        ui.executando=False;ui.planilha=Var(str(fonte));ui.pasta=Var('')
        ui.acoes_vars={a:Var(a=='liberacao') for a in beta.ACOES};ui.conferencia=ui.assinatura_conferencia()
        ui.acoes_vars['juntada'].set(True);erros=[]
        with patch.object(beta.messagebox,'showerror',lambda t,s:erros.append(s)):ui.iniciar(False)
        self.assertTrue(erros);self.assertIn('mudaram',erros[0])
    def test_pasta_irrelevante_sem_juntada(self):
        fonte=self.workbook(['Processo'],[[P1]]);ui=beta.Aplicativo.__new__(beta.Aplicativo)
        ui.planilha=Var(str(fonte));ui.pasta=Var('pasta1');ui.acoes_vars={a:Var(a=='liberacao') for a in beta.ACOES}
        antes=ui.assinatura_conferencia();ui.pasta.set('pasta2');self.assertEqual(ui.assinatura_conferencia(),antes)

if __name__=='__main__':unittest.main(verbosity=2)
