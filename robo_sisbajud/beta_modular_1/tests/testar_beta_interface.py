"""Testes da interface beta sem desktop ou interação com o SAJ."""
import ast
import hashlib
import importlib.util
import itertools
import queue
import sys
import tempfile
import threading
import unittest
from collections import OrderedDict
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location('saj_beta_interface_test', next(x for x in (ROOT/'saj_beta.py',ROOT.parent/'saj_beta.py') if x.is_file()))
beta = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(beta)
COMBINACOES = [tuple(a for a, marcado in zip(beta.ACOES, bits) if marcado)
               for bits in itertools.product((False, True), repeat=3) if any(bits)]
P1, P2 = '0000001-00.0000.0.00.0000', '0000002-00.0000.0.00.0000'


class Var:
    def __init__(self, value=''): self.value = value
    def get(self): return self.value
    def set(self, value): self.value = value


class Controle:
    def __init__(self): self.state = 'normal'; self.configuracoes = []
    def configure(self, **kw):
        self.configuracoes.append(kw)
        if 'state' in kw: self.state = kw['state']


class Tabela:
    def __init__(self): self.rows = OrderedDict(); self.sel = (); self.counter = 0
    def insert(self, parent, position, values, tags=()):
        self.counter += 1
        iid = 'i'+str(self.counter)
        self.rows[iid] = dict(zip(('processo','status','codigo','resultado'), values))
        self.rows[iid]['tags'] = tags
        return iid
    def get_children(self): return tuple(self.rows)
    def selection(self): return self.sel
    def selection_set(self, ids): self.sel = (ids,) if isinstance(ids, str) else tuple(ids)
    def delete(self, *ids):
        for iid in ids: self.rows.pop(iid)
        self.sel = tuple(i for i in self.sel if i in self.rows)
    def set(self, iid, column, value=None):
        if value is not None: self.rows[iid][column] = value
        return self.rows[iid][column]
    def item(self, iid, **kw): self.rows[iid].update(kw)
    def see(self, iid): pass
    def identify_row(self, y):
        ids = self.get_children(); index = y//10-1
        return ids[index] if 0 <= index < len(ids) else ''
    def focus(self, iid): self.focused = iid


class Menu:
    def __init__(self): self.popups = []; self.releases = 0
    def tk_popup(self, *xy): self.popups.append(xy)
    def grab_release(self): self.releases += 1


class Root:
    def __init__(self): self.iconified = False; self.afters = []
    def iconify(self): self.iconified = True
    def deiconify(self): self.iconified = False
    def after(self, *args): self.afters.append(args)


class Log(Controle):
    def insert(self, *args): pass
    def see(self, *args): pass


class ThreadAdiada:
    criadas = []
    def __init__(self, target, daemon): self.target = target; self.daemon = daemon; self.started = False; self.criadas.append(self)
    def start(self): self.started = True


class BetaInterfaceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.pasta = Path(self.tmp.name)
        self.excel = self.pasta/'entrada.xlsx'
        self.escrever_excel(self.excel, [('Processo','Status'), (P1,'IRRISÓRIO'), (P2,'PARCIAL')])
        for p in (P1,P2): (self.pasta/(p+'.pdf')).write_bytes(b'%PDF original de teste')
        self.originais = {p:hashlib.sha256(p.read_bytes()).hexdigest() for p in self.pasta.iterdir()}
        self.app = beta.Aplicativo.__new__(beta.Aplicativo)
        self.app.root, self.app.eventos = Root(), queue.Queue()
        self.app.parar, self.app.executando = threading.Event(), False
        self.app.planilha, self.app.pasta = Var(str(self.excel)), Var(str(self.pasta))
        self.app.inicio = Var(False)
        self.app.acoes_vars = {a:Var(True) for a in beta.ACOES}
        self.app.controles = [Controle() for _ in range(10)]
        self.app.entrada_pasta, self.app.botao_pasta = self.app.controles[:2]
        self.app.btn_individual, self.app.btn_lote, self.app.btn_excluir = Controle(), Controle(), Controle()
        self.app.conferencia, self.app.planilha_carregada = None, ''
        self.app.excluidas, self.app.registros = set(), {}
        self.app.tabela, self.app.menu_processos = Tabela(), Menu()
        self.app.estado, self.app.atual, self.app.resumo = Var(), Var(), Var()
        self.app.progresso, self.app.log = {}, Log()
        self.app.mostrar_historico = lambda: None
        self.erros, self.perguntas, self.infos = [], [], []
        self.patches = [
            patch.object(beta.messagebox,'showerror',side_effect=lambda *a:self.erros.append(a)),
            patch.object(beta.messagebox,'showinfo',side_effect=lambda *a:self.infos.append(a)),
            patch.object(beta.messagebox,'askyesno',side_effect=lambda *a:self.perguntas.append(a) or False),
        ]
        for p in self.patches: p.start()
        ThreadAdiada.criadas.clear()

    def tearDown(self):
        for p in reversed(self.patches): p.stop()
        for p, h in self.originais.items():
            self.assertEqual(hashlib.sha256(p.read_bytes()).hexdigest(), h, p)
        self.tmp.cleanup()

    @staticmethod
    def escrever_excel(path, rows):
        wb = Workbook()
        for row in rows: wb.active.append(row)
        wb.save(path); wb.close()

    def modo(self, acoes):
        for a, v in self.app.acoes_vars.items(): v.set(a in acoes)

    def conferir(self, acoes=beta.ACOES):
        self.modo(acoes); self.app.conferir()
        self.assertFalse(self.erros, self.erros)
        self.assertIsNotNone(self.app.conferencia)

    def dependencias(self, historico=()):
        deps = {n:SimpleNamespace() for n in ('pyautogui','pyperclip','cv2','pywinauto')}
        return (patch.dict(sys.modules,deps),
                patch.object(beta,'ler_historico',return_value=set(historico)),
                patch.object(beta,'validar_referencias',return_value=[]),
                patch.object(beta.threading,'Thread',ThreadAdiada))

    def iniciar_fake(self, lote=True, historico=()):
        a,b,c,d = self.dependencias(historico)
        with a,b,c,d: self.app.iniciar(lote)

    def test_sete_combinacoes_e_nenhuma_acao(self):
        self.assertEqual(len(COMBINACOES),7)
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                self.modo(acoes); self.assertEqual(self.app.opcoes_execucao(),acoes)
        self.modo(())
        with self.assertRaisesRegex(ValueError,'pelo menos uma'): self.app.opcoes_execucao()
        self.app.conferir()
        self.assertIsNone(self.app.conferencia)
        self.assertEqual(self.app.btn_lote.state,'disabled')
        self.assertIn('pelo menos uma',self.erros[-1][1])
        self.iniciar_fake()
        self.assertFalse(ThreadAdiada.criadas)

    def test_imagens_exatamente_das_acoes_escolhidas(self):
        base = {'processo_rotulo','pendencias','limpar'}
        extras = {
            'juntada':{'importar','codigo','aguardando','salvar'},
            'sigilo':{'sigilo_lista','sigilo_grupo','sigilo_decisao_texto','sigilo_decisao_azul','sigilo_cadeado','sigilo_remover'},
            'liberacao':{'aguardando','caixa_desmarcada','caixa_marcada','liberar'},
        }
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                esperado = base | set().union(*(extras[a] for a in acoes))
                self.assertEqual(beta.imagens_necessarias(acoes),esperado)
                self.assertEqual(beta.imagens_necessarias(acoes,True),esperado|{'fluxo','digitalizacao'})

    def test_pdf_e_status_so_obrigatorios_com_juntada(self):
        simples = self.pasta/'somente processos.xlsx'
        self.escrever_excel(simples,[('Processo',),(P1,)])
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                if 'juntada' in acoes:
                    with self.assertRaisesRegex(ValueError,'pasta de PDFs'):
                        beta.ler_planilha(simples,'',acoes)
                    with self.assertRaisesRegex(ValueError,'status'):
                        beta.ler_planilha(simples,self.pasta,acoes)
                else:
                    registros = beta.ler_planilha(simples,'pasta inexistente',acoes)
                    self.assertEqual([(r['erro'],r['pdf'],r['codigo']) for r in registros],[('','','')])
        invalido = self.pasta/'status invalido.xlsx'
        self.escrever_excel(invalido,[('Processo','Status'),(P1,'INVENTADO')])
        self.assertEqual(beta.ler_planilha(invalido,'',('sigilo',))[0]['erro'],'')
        self.assertIn('Status inválido',beta.ler_planilha(invalido,self.pasta,('juntada',))[0]['erro'])

    def test_assinatura_detecta_planilha_modo_e_pasta_apenas_juntada(self):
        self.modo(('sigilo',)); uma = self.app.assinatura_conferencia()
        self.app.pasta.set('outra pasta'); self.assertEqual(self.app.assinatura_conferencia(),uma)
        self.modo(('juntada',)); duas = self.app.assinatura_conferencia()
        self.app.pasta.set(str(self.pasta)); self.assertNotEqual(self.app.assinatura_conferencia(),duas)
        outra = self.pasta/'outro.xlsx'; self.escrever_excel(outra,[('Processo',),(P1,)])
        self.app.planilha.set(str(outra)); self.assertNotEqual(self.app.assinatura_conferencia()[0],uma[0])

    def test_troca_acao_invalida_conferencia_sem_alterar_lista_resultados_exclusoes(self):
        self.conferir()
        iid = self.app.tabela.get_children()[0]
        self.app.registros[iid].update(tentado=True,resultado='Erro anterior que deve permanecer')
        self.app.tabela.set(iid,'resultado','Erro anterior que deve permanecer')
        self.app.excluidas.add((99,'9'*20))
        antes = (dict(self.app.registros),dict(self.app.tabela.rows),set(self.app.excluidas))
        self.modo(('sigilo',))
        with patch.object(beta,'ler_planilha',side_effect=AssertionError('não reler ao alternar')):
            self.app.atualizar_opcoes()
        self.assertIsNone(self.app.conferencia)
        self.assertEqual((self.app.registros,self.app.tabela.rows,self.app.excluidas),antes)
        self.assertEqual(self.app.btn_lote.state,'disabled')
        self.assertEqual(self.app.entrada_pasta.state,'disabled')

    def test_controles_congelados_e_restaurados_com_pasta_condicional(self):
        self.conferir(('sigilo',))
        self.assertEqual(self.app.entrada_pasta.state,'disabled')
        self.app.definir_executando(True)
        self.assertTrue(all(c.state=='disabled' for c in self.app.controles))
        self.assertTrue(all(c.state=='disabled' for c in (self.app.btn_lote,self.app.btn_individual,self.app.btn_excluir)))
        assinatura, ids = self.app.conferencia, self.app.tabela.get_children()
        self.app.tabela.selection_set(ids)
        self.app.atualizar_opcoes(); self.app.excluir_selecionados(); self.app.conferir()
        self.assertEqual(self.app.conferencia,assinatura)
        self.assertEqual(self.app.tabela.get_children(),ids)
        self.app.definir_executando(False)
        self.assertEqual(self.app.entrada_pasta.state,'disabled')
        self.assertEqual(self.app.botao_pasta.state,'disabled')
        self.assertTrue(all(c.state=='normal' for c in self.app.controles[2:]))
        self.assertEqual(self.app.btn_lote.state,'normal')
        self.modo(('juntada',)); self.app.atualizar_opcoes()
        self.assertEqual(self.app.entrada_pasta.state,'normal')
        self.assertEqual(self.app.botao_pasta.state,'normal')
        self.assertEqual(self.app.btn_lote.state,'disabled')

    def test_exclusao_persistente_nos_sete_modos_sem_modificar_originais(self):
        self.conferir()
        primeiro, segundo = self.app.tabela.get_children()
        self.app.registros[segundo].update(tentado=True,resultado='Concluído: resultado preservado')
        self.app.tabela.selection_set(primeiro)
        with patch('builtins.open',side_effect=AssertionError('excluir não pode gravar arquivos')):
            self.app.excluir_selecionados()
        excluidas = {(2,beta.digitos(P1))}
        self.assertEqual(self.app.excluidas,excluidas)
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                self.conferir(acoes)
                self.assertEqual(len(self.app.registros),1)
                r = next(iter(self.app.registros.values()))
                self.assertEqual(r['processo'],P2)
                self.assertEqual(r['resultado'],'Concluído: resultado preservado')
                self.assertTrue(r['tentado'])
                self.assertEqual(self.app.excluidas,excluidas)

    def test_outro_excel_limpa_exclusoes_sem_transferir_resultados(self):
        self.conferir()
        primeiro = self.app.tabela.get_children()[0]
        self.app.tabela.selection_set(primeiro); self.app.excluir_selecionados()
        for r in self.app.registros.values(): r.update(tentado=True,resultado='Resultado da planilha antiga')
        outra = self.pasta/'nova.xlsx'; self.escrever_excel(outra,[('Processo',),(P1,),(P2,)])
        self.app.planilha.set(str(outra)); self.conferir(('sigilo',))
        self.assertEqual(self.app.excluidas,set())
        self.assertEqual(len(self.app.registros),2)
        self.assertTrue(all(not r['tentado'] and not r['resultado'] for r in self.app.registros.values()))

    def test_reconferencia_atualiza_validacao_sem_pdf_ou_status_antigo(self):
        simples = self.pasta/'status ruim.xlsx'
        self.escrever_excel(simples,[('Processo','Status'),(P1,'ERRADO')])
        self.app.planilha.set(str(simples)); self.conferir(('juntada',))
        self.assertIn('Status inválido',next(iter(self.app.registros.values()))['erro'])
        self.conferir(('sigilo','liberacao'))
        self.assertEqual(next(iter(self.app.registros.values()))['erro'],'')
        self.assertIn('1 válidos',self.app.resumo.get())
        self.conferir(('juntada',))
        self.assertIn('Status inválido',next(iter(self.app.registros.values()))['erro'])

    def test_todos_sete_modos_chegam_ao_motor_e_snapshot_nao_muda(self):
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                self.app.definir_executando(False); self.conferir(acoes)
                self.iniciar_fake()
                self.assertTrue(self.app.executando)
                self.assertTrue(ThreadAdiada.criadas[-1].started)
                esperados = [dict(r) for r in self.app.registros.values()]
                self.modo(())
                for r in self.app.registros.values(): r['processo'] = 'alteração posterior apenas no fake'
                chamadas = []
                motor = SimpleNamespace(executar=lambda *a:chamadas.append(a))
                with patch.object(beta,'Motor',return_value=motor): ThreadAdiada.criadas[-1].target()
                tarefas, escolhidas, inicio = chamadas[0]
                self.assertEqual(escolhidas,acoes)
                self.assertEqual([r for iid,r in tarefas],esperados)
                self.assertFalse(inicio)
                self.assertFalse(self.erros,self.erros)

    def test_sem_juntada_aceita_historico_e_resultado_anterior_lote_e_individual(self):
        sem_juntada = [a for a in COMBINACOES if 'juntada' not in a]
        for acoes,lote in itertools.product(sem_juntada,(False,True)):
            with self.subTest(acoes=acoes,lote=lote):
                self.app.definir_executando(False); self.conferir(acoes)
                for iid,r in self.app.registros.items():
                    r.update(tentado=True,resultado='Erro anterior: outra ação')
                    self.app.tabela.set(iid,'resultado','Erro anterior: outra ação')
                self.app.tabela.selection_set(self.app.tabela.get_children()[0])
                antes = len(ThreadAdiada.criadas)
                self.iniciar_fake(lote,historico=(beta.digitos(P1),beta.digitos(P2)))
                self.assertEqual(len(ThreadAdiada.criadas),antes+1)
                self.assertFalse(self.perguntas)
                self.assertFalse(self.erros,self.erros)

    def test_juntada_historico_pula_no_lote_e_pede_confirmacao_individual(self):
        self.conferir(('juntada',))
        self.iniciar_fake(True,historico=(beta.digitos(P1),beta.digitos(P2)))
        self.assertFalse(ThreadAdiada.criadas)
        self.assertIn('Não há processos',self.erros[-1][1])
        self.assertTrue(all('Juntada já tentada' in r['resultado'] for r in self.app.tabela.rows.values()))
        self.erros.clear()
        self.app.tabela.selection_set(self.app.tabela.get_children()[0])
        self.iniciar_fake(False,historico=(beta.digitos(P1),))
        self.assertEqual(len(self.perguntas),1)
        self.assertFalse(ThreadAdiada.criadas)

    def test_inicio_exige_uma_selecao_e_conferencia_atual(self):
        self.conferir(('sigilo',))
        for ids in ((),self.app.tabela.get_children()):
            self.app.tabela.selection_set(ids); self.iniciar_fake(False)
            self.assertIn('exatamente um',self.erros[-1][1])
        self.modo(('liberacao',)); self.iniciar_fake(True)
        self.assertIn('mudaram',self.erros[-1][1])
        self.assertFalse(ThreadAdiada.criadas)

    def test_erros_de_linha_impedem_individual_e_lote_sem_elegiveis(self):
        self.conferir(('sigilo',))
        for r in self.app.registros.values(): r['erro'] = 'Número inválido'
        self.app.tabela.selection_set(self.app.tabela.get_children()[0])
        self.iniciar_fake(False); self.assertIn('Número inválido',self.erros[-1][1])
        self.iniciar_fake(True); self.assertIn('Não há processos',self.erros[-1][1])
        self.assertFalse(ThreadAdiada.criadas)

    def test_imagens_faltantes_impedem_inicio_somente_das_acoes_escolhidas(self):
        self.conferir(('sigilo',))
        a,b,c,d = self.dependencias()
        with a,b,d,patch.object(beta,'validar_referencias',return_value=['cadeado inválido']) as validar:
            self.app.iniciar(True)
        self.assertEqual(validar.call_args.args[0],beta.imagens_necessarias(('sigilo',)))
        self.assertIn('cadeado inválido',self.erros[-1][1])
        self.assertFalse(ThreadAdiada.criadas)

    def test_poll_registra_resultado_e_restaura_controles_no_fim(self):
        self.conferir(('sigilo',))
        iid = self.app.tabela.get_children()[0]
        self.app.definir_executando(True)
        self.app.eventos.put(('linha',(iid,'Erro na etapa X: motivo concreto')))
        self.app.eventos.put(('fim',None)); self.app.poll()
        self.assertTrue(self.app.registros[iid]['tentado'])
        self.assertEqual(self.app.registros[iid]['resultado'],'Erro na etapa X: motivo concreto')
        self.assertEqual(self.app.tabela.rows[iid]['tags'],('erro',))
        self.assertFalse(self.app.executando)
        self.assertEqual(self.app.btn_excluir.state,'normal')
        self.assertEqual(self.app.entrada_pasta.state,'disabled')

    def test_contexto_preserva_multiselecao_e_delete_bloqueado_ao_rodar(self):
        self.conferir(('sigilo',)); ids = self.app.tabela.get_children()
        self.app.tabela.selection_set(ids)
        self.app.abrir_menu_processos(SimpleNamespace(y=10,x_root=1,y_root=2))
        self.assertEqual(self.app.tabela.selection(),ids)
        self.app.tabela.selection_set(ids[:1])
        self.app.abrir_menu_processos(SimpleNamespace(y=20,x_root=1,y_root=2))
        self.assertEqual(self.app.tabela.selection(),ids[1:])
        popups = len(self.app.menu_processos.popups)
        self.app.abrir_menu_processos(SimpleNamespace(y=999,x_root=1,y_root=2))
        self.assertEqual(len(self.app.menu_processos.popups),popups)
        self.app.definir_executando(True); self.app.excluir_selecionados()
        self.assertEqual(self.app.tabela.get_children(),ids)

    def test_layout_850_e_acoes_na_lateral_sem_reintroduzir_atalhos(self):
        source = (next(x for x in (ROOT/'saj_beta.py',ROOT.parent/'saj_beta.py') if x.is_file())).read_text()
        tree = ast.parse(source)
        cls = next(n for n in tree.body if isinstance(n,ast.ClassDef) and n.name=='Aplicativo')
        init = next(n for n in cls.body if isinstance(n,ast.FunctionDef) and n.name=='__init__')
        trecho = ast.get_source_segment(source,init)
        self.assertIn('root.minsize(850, 540)',trecho)
        self.assertIn("selectmode='extended'",trecho)
        self.assertIn("self.tabela.bind('<Delete>', self.excluir_selecionados)",trecho)
        self.assertIn("self.tabela.bind('<Button-3>', self.abrir_menu_processos)",trecho)
        self.assertIn('ttk.Checkbutton(config, text=texto, variable=self.acoes_vars[acao]',trecho)
        self.assertIn('self.controles.append(caixa)',trecho)
        self.assertNotIn('ttk.Checkbutton(botoes',trecho)
        self.assertNotIn('Salvar, liberar e limpar',trecho)
        self.assertNotIn('Intervalo: 3 segundos',trecho)


if __name__ == '__main__': unittest.main(verbosity=2)
