"""Testes da BETA modular com XLSX sintéticos. Não abre desktop nem opera o SAJ."""
import csv
import hashlib
import importlib.util
import itertools
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

PASTA = Path(__file__).resolve().parent
FONTE = next((p for p in (PASTA / 'saj_beta.py', PASTA.parent / 'saj_beta.py') if p.is_file()), None)
if FONTE is None:
    raise RuntimeError('Coloque este teste ao lado de saj_beta.py ou na subpasta tests.')
spec = importlib.util.spec_from_file_location('saj_beta_em_teste', FONTE)
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

# Ano/segmentos zero: identificadores artificiais, sem dados de casos reais.
PROCESSOS = ('0000001-00.0000.0.00.0000', '0000002-00.0000.0.00.0000', '0000003-00.0000.0.00.0000')
COMBINACOES = [tuple(a for a, ativo in zip(app.ACOES, bits) if ativo)
               for bits in itertools.product((False, True), repeat=3) if any(bits)]


class Variavel:
    def __init__(self, valor): self.valor = valor
    def get(self): return self.valor
    def set(self, valor): self.valor = valor


class Controle:
    def __init__(self): self.estado = None
    def configure(self, **kw): self.estado = kw.get('state', self.estado)


class BetaPlanilha(unittest.TestCase):
    def setUp(self):
        self.temporario = tempfile.TemporaryDirectory(prefix='beta_planilha_sintetica_')
        self.raiz = Path(self.temporario.name)
        self.pdfs = self.raiz / 'pdfs'
        self.pdfs.mkdir()
        self.numero_arquivo = 0

    def tearDown(self):
        self.temporario.cleanup()

    def planilha(self, cabecalhos, linhas):
        self.numero_arquivo += 1
        caminho = self.raiz / ('sintetica_' + str(self.numero_arquivo) + '.xlsx')
        livro = Workbook()
        livro.active.append(cabecalhos)
        for linha in linhas:
            livro.active.append(linha)
        livro.save(caminho)
        livro.close()
        return caminho

    def pdf(self, processo, nome=None):
        caminho = self.pdfs / (nome or processo + '.pdf')
        # Fixture de existência: o robô associa nomes, não interpreta o conteúdo do PDF.
        caminho.write_bytes(b'%PDF-1.4\n% fixture sintetica de existencia\n%%EOF\n')
        return caminho

    def test_sete_combinacoes_e_ordem_canonica(self):
        self.assertEqual(len(COMBINACOES), 7)
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                self.assertEqual(app.normalizar_acoes(reversed(acoes)), acoes)
                self.assertEqual(app.normalizar_acoes(acoes + acoes), acoes)
        for invalida in ((), [], None, ('inexistente',), ('juntada', 'inexistente')):
            with self.subTest(invalida=invalida), self.assertRaises(ValueError):
                app.normalizar_acoes(invalida)

    def test_sete_combinacoes_com_xlsx_reais(self):
        for processo in PROCESSOS:
            self.pdf(processo)
        caminho = self.planilha(['numero do processo', 'status'],
                                list(zip(PROCESSOS, ('IRRISÓRIO', 'PARCIAL', 'INTEGRAL'))))
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                linhas = app.ler_planilha(caminho, self.pdfs, acoes)
                self.assertEqual(len(linhas), 3)
                self.assertFalse(any(r['erro'] for r in linhas))
                self.assertEqual([r['processo'] for r in linhas], list(PROCESSOS))
                self.assertTrue(all(app.digitos(r['processo']).startswith('000000') for r in linhas))
                if 'juntada' in acoes:
                    self.assertEqual([r['codigo'] for r in linhas], ['938', '939', '939'])
                    self.assertTrue(all(Path(r['pdf']).is_file() for r in linhas))
                else:
                    self.assertEqual([r['codigo'] for r in linhas], ['', '', ''])
                    self.assertEqual([r['pdf'] for r in linhas], ['', '', ''])

    def test_sem_juntada_so_processo_e_sem_varrer_pasta(self):
        for cabecalho in ('Processo', 'numero do processo', 'Número do Processo'):
            caminho = self.planilha([cabecalho], [[PROCESSOS[0]]])
            for acoes in COMBINACOES:
                if 'juntada' in acoes:
                    continue
                for pasta in ('', self.raiz / 'inexistente', self.pdfs):
                    with self.subTest(cabecalho=cabecalho, acoes=acoes, pasta=str(pasta)):
                        with patch.object(Path, 'iterdir', side_effect=AssertionError('Não deve varrer pasta sem juntada')):
                            linhas = app.ler_planilha(caminho, pasta, acoes)
                        self.assertEqual(linhas[0]['erro'], '')
                        self.assertEqual(linhas[0]['codigo'], '')
                        self.assertEqual(linhas[0]['pdf'], '')

    def test_status_invalido_ignorado_sem_juntada(self):
        caminho = self.planilha(['Processo', 'status'], [[PROCESSOS[0], 'IGNORAR ESTE TEXTO']])
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                self.pdf(PROCESSOS[0])
                registro = app.ler_planilha(caminho, self.pdfs, acoes)[0]
                if 'juntada' in acoes:
                    self.assertIn('Status inválido', registro['erro'])
                else:
                    self.assertFalse(registro['erro'])
                    self.assertFalse(registro['codigo'])

    def test_status_e_pasta_exigidos_somente_na_juntada(self):
        caminho = self.planilha(['Processo'], [[PROCESSOS[0]]])
        for acoes in COMBINACOES:
            if 'juntada' not in acoes:
                continue
            with self.subTest(acoes=acoes):
                with self.assertRaisesRegex(ValueError, 'status'):
                    app.ler_planilha(caminho, self.pdfs, acoes)
                for pasta in ('', self.raiz / 'nao_existe', caminho):
                    with self.assertRaisesRegex(ValueError, 'pasta'):
                        app.ler_planilha(caminho, pasta, acoes)

    def test_pdf_ausente_duplicado_e_mascara(self):
        caminho = self.planilha(['Processo', 'status'], [[PROCESSOS[0], 'INTEGRAL']])
        self.assertIn('PDF ausente', app.ler_planilha(caminho, self.pdfs, ('juntada',))[0]['erro'])
        pdf = self.pdf(PROCESSOS[0], app.digitos(PROCESSOS[0]) + '.PDF')
        registro = app.ler_planilha(caminho, self.pdfs, ('juntada',))[0]
        self.assertFalse(registro['erro'])
        self.assertEqual(Path(registro['pdf']), pdf.resolve())
        self.pdf(PROCESSOS[0])
        self.assertIn('mais de um', app.ler_planilha(caminho, self.pdfs, ('juntada',))[0]['erro'])
        for acoes in (('sigilo',), ('liberacao',), ('sigilo', 'liberacao')):
            self.assertFalse(app.ler_planilha(caminho, self.pdfs, acoes)[0]['erro'])

    def test_processo_texto_len20_sempre(self):
        # Não se inventa preenchimento de zeros para números que o Excel já converteu.
        entradas = [0, 12345678901234567890, '00001', '0' * 21]
        for valor in entradas:
            caminho = self.planilha(['Processo', 'status'], [[valor, 'INTEGRAL']])
            for acoes in COMBINACOES:
                with self.subTest(valor=str(valor), acoes=acoes):
                    registro = app.ler_planilha(caminho, self.pdfs, acoes)[0]
                    self.assertIn('Texto com 20 dígitos', registro['erro'])
        caminho = self.planilha(['Processo'], [['0' * 19 + '1']])
        self.assertEqual(app.ler_planilha(caminho, '', ('sigilo',))[0]['processo'], '0' * 19 + '1')

    def test_duplicados_por_digitos_e_linhas_vazias(self):
        self.pdf(PROCESSOS[0])
        caminho = self.planilha(['Processo', 'status'], [
            [PROCESSOS[0], 'INTEGRAL'],
            [None, None],
            [app.digitos(PROCESSOS[0]), 'PARCIAL'],
        ])
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes):
                linhas = app.ler_planilha(caminho, self.pdfs, acoes)
                self.assertEqual(len(linhas), 2)
                self.assertFalse(linhas[0]['erro'])
                self.assertIn('repetido', linhas[1]['erro'])
                self.assertEqual([r['linha'] for r in linhas], [2, 4])

    def test_cabecalho_processo_ausente_e_zero_acoes(self):
        caminho = self.planilha(['outra coluna'], [['teste']])
        for acoes in COMBINACOES:
            with self.subTest(acoes=acoes), self.assertRaisesRegex(ValueError, 'cabeçalho'):
                app.ler_planilha(caminho, self.pdfs, acoes)
        with self.assertRaises(ValueError):
            app.ler_planilha(caminho, self.pdfs, ())

    def test_imagens_por_cada_uma_das_sete_combinacoes(self):
        base = {'processo_rotulo', 'pendencias', 'limpar'}
        por_acao = {
            'juntada': {'importar', 'codigo', 'aguardando', 'salvar'},
            'sigilo': {'sigilo_lista', 'sigilo_grupo', 'sigilo_decisao_texto',
                       'sigilo_decisao_azul', 'sigilo_cadeado', 'sigilo_remover'},
            'liberacao': {'aguardando', 'caixa_desmarcada', 'caixa_marcada', 'liberar'},
        }
        for acoes in COMBINACOES:
            esperado = base.union(*(por_acao[a] for a in acoes))
            with self.subTest(acoes=acoes):
                self.assertEqual(app.imagens_necessarias(acoes), esperado)
                self.assertEqual(app.imagens_necessarias(acoes, True), esperado | {'fluxo', 'digitalizacao'})
        with self.assertRaises(ValueError):
            app.imagens_necessarias(())

    def test_assinatura_muda_com_acoes_e_nao_exige_pasta_sem_juntada(self):
        caminho = self.planilha(['Processo'], [[PROCESSOS[0]]])
        ui = app.Aplicativo.__new__(app.Aplicativo)
        ui.planilha, ui.pasta = Variavel(str(caminho)), Variavel('pasta_qualquer')
        assinaturas = {acoes: ui.assinatura_conferencia(acoes) for acoes in COMBINACOES}
        self.assertEqual(len(set(assinaturas.values())), 7)
        for acoes in COMBINACOES:
            self.assertEqual(assinaturas[acoes][-1], acoes)
            self.assertEqual(assinaturas[acoes][3], 'pasta_qualquer' if 'juntada' in acoes else '')
        ui.pasta.set('outra_pasta')
        self.assertEqual(ui.assinatura_conferencia(('sigilo',)), assinaturas[('sigilo',)])
        self.assertNotEqual(ui.assinatura_conferencia(('juntada',)), assinaturas[('juntada',)])

    def test_mudar_caixas_invalida_conferencia_e_bloqueia_inicio(self):
        ui = app.Aplicativo.__new__(app.Aplicativo)
        ui.executando = False
        ui.conferencia = ('conferida',)
        ui.estado = Variavel('Pronto')
        ui.acoes_vars = {a: Variavel(a == 'sigilo') for a in app.ACOES}
        ui.controles = []
        for nome in ('entrada_pasta', 'botao_pasta', 'btn_individual', 'btn_lote', 'btn_excluir'):
            setattr(ui, nome, Controle())
        ui.atualizar_opcoes()
        self.assertIsNone(ui.conferencia)
        self.assertEqual(ui.btn_individual.estado, 'disabled')
        self.assertEqual(ui.btn_lote.estado, 'disabled')
        self.assertEqual(ui.entrada_pasta.estado, 'disabled')
        ui.executando, ui.conferencia = True, ('em execução',)
        ui.atualizar_opcoes()
        self.assertEqual(ui.conferencia, ('em execução',))

    def test_historico_beta_separado_e_legado_somente_leitura(self):
        legado = self.raiz / 'historico_saj.csv'
        beta = self.raiz / 'historico_saj_beta.csv'
        with legado.open('w', encoding='utf-8-sig', newline='') as arquivo:
            escritor = csv.writer(arquivo, delimiter=';')
            escritor.writerow(['data', 'processo', 'status', 'codigo', 'fase', 'resultado'])
            escritor.writerow(['data sintética', PROCESSOS[0], '', '', 'teste', 'tentativa sintética'])
        antes = hashlib.sha256(legado.read_bytes()).digest()
        registro = dict(processo=PROCESSOS[1], status='', codigo='', pdf='')
        with patch.object(app, 'HISTORICO', beta), patch.object(app, 'HISTORICO_ANTERIOR', legado):
            self.assertEqual(app.ler_historico(), {app.digitos(PROCESSOS[0])})
            app.gravar_historico(registro, 'retirar sigilo', 'ok', ('sigilo', 'liberacao'), 'sigilo', 'concluida')
            app.gravar_historico(registro, 'liberar', 'ok', ('liberacao',), 'liberacao', 'concluida')
            self.assertEqual(app.ler_historico(), {app.digitos(PROCESSOS[0])})
            app.gravar_historico(registro, 'importar PDF', 'tentativa', ('juntada',), 'juntada', 'iniciada')
            self.assertEqual(app.ler_historico(), {app.digitos(PROCESSOS[0]), app.digitos(PROCESSOS[1])})
            with beta.open(encoding='utf-8-sig', newline='') as arquivo:
                linhas = list(csv.DictReader(arquivo, delimiter=';'))
            self.assertEqual(len(linhas), 3)
            self.assertEqual(linhas[0]['acoes'], 'sigilo+liberacao')
            self.assertEqual(linhas[0]['acao'], 'sigilo')
            self.assertEqual(linhas[2]['estado'], 'iniciada')
        self.assertEqual(hashlib.sha256(legado.read_bytes()).digest(), antes)


if __name__ == '__main__':
    unittest.main(verbosity=2)
