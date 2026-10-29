"""Gera as bases fictícias do portfólio (ASO + custos de exames).

Tudo é inventado: nomes, CPFs, CNPJs, unidades, preços e volumes.
Semente fixa: rodar de novo produz exatamente os mesmos dados.
"""
import json
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

random.seed(20261026)
REF = date(2026, 10, 26)  # data de referência da base fictícia
OUT = Path(__file__).parent / "dados"
OUT.mkdir(exist_ok=True)

# ----------------------------------------------------------------------------
# Cadastros fictícios
# ----------------------------------------------------------------------------
PRIMEIROS = """ANA BRUNO CARLA DANIEL EDUARDA FELIPE GABRIELA HEITOR ISABELA JOAO
JULIANA KAIO LARISSA LUCAS MARIANA MATEUS NATALIA OTAVIO PAULA RAFAEL RENATA
SAMUEL TATIANE THIAGO VANESSA VINICIUS YASMIN ALINE ANDRE BEATRIZ CAIO CAMILA
DIEGO ELISA FABIO FERNANDA GUSTAVO HELENA IGOR JESSICA LEONARDO LIVIA MARCOS
MELISSA NICOLAS PRISCILA RODRIGO SABRINA TIAGO VITORIA WAGNER ARTHUR BIANCA
CESAR DEBORA ENZO FLAVIA GIOVANA HUGO IARA""".split()
SOBRENOMES = """SILVA SOUZA COSTA SANTOS OLIVEIRA PEREIRA RODRIGUES ALMEIDA NASCIMENTO
LIMA ARAUJO FERNANDES CARVALHO GOMES MARTINS ROCHA RIBEIRO ALVES MONTEIRO
MENDES BARROS FREITAS BARBOSA PINTO MOURA CAVALCANTI DIAS CASTRO CAMPOS
CARDOSO TEIXEIRA CORREIA MACHADO VIEIRA NUNES MOREIRA FONSECA AZEVEDO BATISTA
REIS DUARTE LOPES MARQUES ANDRADE MIRANDA PIRES""".split()

UNIDADES = {  # unidade: (peso, [(função, GHE)])
    "PRODUÇÃO": (30, [("OPERADOR DE PRODUÇÃO", "GHE 01 - PRODUÇÃO"),
                      ("LÍDER DE PRODUÇÃO", "GHE 01 - PRODUÇÃO"),
                      ("AUXILIAR DE PRODUÇÃO", "GHE 01 - PRODUÇÃO")]),
    "ADMINISTRATIVO": (22, [("ANALISTA ADMINISTRATIVO", "GHE 05 - ADMINISTRATIVO"),
                            ("ASSISTENTE ADMINISTRATIVO", "GHE 05 - ADMINISTRATIVO"),
                            ("COORDENADOR", "GHE 05 - ADMINISTRATIVO")]),
    "MANUTENÇÃO": (14, [("MECÂNICO", "GHE 02 - MANUTENÇÃO"),
                        ("ELETRICISTA", "GHE 02 - MANUTENÇÃO"),
                        ("TÉCNICO DE MANUTENÇÃO", "GHE 02 - MANUTENÇÃO")]),
    "LOGÍSTICA": (11, [("MOTORISTA", "GHE 03 - LOGÍSTICA"),
                       ("OPERADOR DE EMPILHADEIRA", "GHE 03 - LOGÍSTICA"),
                       ("CONFERENTE", "GHE 03 - LOGÍSTICA")]),
    "CAMPO": (9, [("TÉCNICO DE CAMPO", "GHE 04 - CAMPO"),
                  ("SUPERVISOR DE CAMPO", "GHE 04 - CAMPO")]),
    "ENGENHARIA": (6, [("ENGENHEIRO", "GHE 05 - ADMINISTRATIVO"),
                       ("PROJETISTA", "GHE 05 - ADMINISTRATIVO")]),
    "QUALIDADE": (5, [("INSPETOR DE QUALIDADE", "GHE 02 - MANUTENÇÃO"),
                      ("ANALISTA DE QUALIDADE", "GHE 05 - ADMINISTRATIVO")]),
    "ESTÁGIO": (3, [("ESTAGIÁRIO", "GHE 05 - ADMINISTRATIVO")]),
}
# funções expostas a risco fazem periódico semestral
SEMESTRAL = {"ELETRICISTA", "OPERADOR DE EMPILHADEIRA", "TÉCNICO DE CAMPO"}
EMPRESAS = {"EMPRESA ALFA": "11.111.111/0001-11", "EMPRESA BETA": "22.222.222/0001-22"}
GESTORES = ["CLAUDIA REIS", "MARCELO DUARTE", "SIMONE ANDRADE", "RICARDO LOPES",
            "PATRICIA MIRANDA", "EDSON FONSECA", "LUANA BATISTA", "ROBERTO PIRES"]

usados = set()


def novo_nome():
    while True:
        n = f"{random.choice(PRIMEIROS)} {random.choice(SOBRENOMES)} {random.choice(SOBRENOMES)}"
        a, b, c = n.split()
        if b != c and n not in usados:
            usados.add(n)
            return n


seq_cpf = iter(range(100001, 999999))


def cpf_falso():
    # começa com 000: formato de CPF, mas obviamente falso
    k = next(seq_cpf)
    return f"000.{k // 1000:03d}.{k % 1000:03d}-{k % 97:02d}"


def email(nome):
    p = nome.lower().split()
    tr = str.maketrans("áàâãéêíóôõúç", "aaaaeeiooouc")
    return f"{p[0]}.{p[-1]}@empresa-exemplo.com".translate(tr)


def faixa_venc(d):
    if d < 0:
        return "Vencido"
    if d == 0:
        return "Vence hoje"
    if d <= 7:
        return "Em até 7 dias"
    if d <= 15:
        return "De 8 a 15 dias"
    if d <= 30:
        return "De 16 a 30 dias"
    if d <= 60:
        return "De 31 a 60 dias"
    if d <= 90:
        return "De 61 a 90 dias"
    if d <= 120:
        return "De 91 a 120 dias"
    return "Acima de 120 dias"


def faixa_consolidada(f):
    return {"Vencido": "Vencido", "Vence hoje": "Até 7 dias", "Em até 7 dias": "Até 7 dias",
            "De 8 a 15 dias": "8 a 30 dias", "De 16 a 30 dias": "8 a 30 dias",
            "De 31 a 60 dias": "31 a 60 dias", "De 61 a 90 dias": "61 a 90 dias"}.get(f, "Acima de 90 dias")


unid_lista = list(UNIDADES)
unid_peso = [UNIDADES[u][0] for u in unid_lista]

# ----------------------------------------------------------------------------
# tblASOS — base de ativos com controle de vencimento
# ----------------------------------------------------------------------------
N = 2900
asos = []
for i in range(N):
    unid = random.choices(unid_lista, unid_peso)[0]
    func, ghe = random.choice(UNIDADES[unid][1])
    emp = "EMPRESA ALFA" if random.random() < 0.68 else "EMPRESA BETA"
    nome = novo_nome()
    semestral = func in SEMESTRAL
    ciclo = 182 if semestral else 365
    # dias até o vencimento: ~9% vencidos, resto espalhado pelo ciclo
    if random.random() < 0.055:
        dias = -random.randint(1, 45)
    else:
        dias = random.randint(0, ciclo - 1)
    venc = REF + timedelta(days=dias)
    ultimo = venc - timedelta(days=ciclo)
    adm = min(ultimo, REF - timedelta(days=random.randint(200, 4000)))
    status = random.choices(["Ativo", "Férias", "Licença Remunerada", "Afastados"],
                            [87, 7, 2, 4])[0]
    fx = faixa_venc(dias)
    rotulo = fx
    if status == "Afastados":
        rotulo = "AFASTADO"
    elif random.random() < 0.006:
        rotulo = "EXCEÇÃO DIRETORIA"

    # agendamento: só para quem vence em até 90 dias e está na base de cálculo
    ag = None
    if status != "Afastados" and rotulo != "EXCEÇÃO DIRETORIA" and dias <= 90:
        p = 0.78 if dias < 0 else 0.62 if dias <= 7 else 0.62 if dias <= 30 else 0.48 if dias <= 60 else 0.20
        if random.random() < p:
            if dias < 0:
                ag = REF + timedelta(days=random.randint(-12, 10))
            else:
                ag = venc - timedelta(days=random.randint(3, min(45, max(4, dias + 12))))
    st_final = "Enviado" if (dias <= 30 and status != "Afastados") or ag else "Pendente"
    tipo = "PERIÓDICO SEMESTRAL" if semestral else "PERIÓDICO"
    asos.append({
        "CPF": cpf_falso(), "NOME": nome, "CHAPA": f"{'A' if emp == 'EMPRESA ALFA' else 'B'}{10000 + i}",
        "EMPRESA": emp, "CNPJ": EMPRESAS[emp], "POSTO DE TRABALHO - GHE": ghe, "FUNÇÃO": func,
        "UNIDADE DE NEGÓCIO": unid, "ADMISSÃO": adm, "E-MAIL": email(nome),
        "GESTOR DIRETO": random.choice(GESTORES), "VENCIMENTO": rotulo, "DATA DE VENCIMENTO": venc,
        "STATUS": status, "STATUS FINAL": st_final, "ÚLTIMO ASO": ultimo,
        "DATA DE AGENDAMENTO": ag, "TIPO DE ASO": tipo,
    })
df_aso = pd.DataFrame(asos)

# ----------------------------------------------------------------------------
# tblRealizados — ASOs feitos de julho a outubro (histórico de antecipação)
# ----------------------------------------------------------------------------
real = []
mes_qtd = {7: 455, 8: 540, 9: 690, 10: 262}
for mes, qtd in mes_qtd.items():
    ultimo_dia = 31 if mes in (7, 8, 10) else 30
    if mes == 10:
        ultimo_dia = 23
    for _ in range(qtd):
        unid = random.choices(unid_lista, unid_peso)[0]
        emp = "EMPRESA ALFA" if random.random() < 0.68 else "EMPRESA BETA"
        dt = date(2026, mes, random.randint(1, ultimo_dia))
        # setembro teve mutirão: antecipação maior
        r = random.random()
        if r < (0.17 if mes != 9 else 0.10):
            ant = -random.randint(1, 40)
        elif r < 0.30:
            ant = random.randint(0, 7)
        elif r < 0.55:
            ant = random.randint(8, 30)
        elif r < 0.80 or mes != 9:
            ant = random.randint(31, 60 if mes != 9 else 90)
        else:
            ant = random.randint(61, 118)
        tipo = random.choices(["PERIÓDICO", "PERIÓDICO SEMESTRAL", "MUDANÇA DE RISCO"], [84, 12, 4])[0]
        real.append({"CPF": cpf_falso(), "NOME": novo_nome(), "EMPRESA": emp,
                     "UNIDADE DE NEGÓCIO": unid, "TIPO DE ASO": tipo,
                     "DATA DE REALIZAÇÃO": dt, "VENCIMENTO ANTERIOR": dt + timedelta(days=ant),
                     "DIAS ANTECIPADOS": ant})
df_real = pd.DataFrame(real)


def faixa_ant(d):
    if d < 0:
        return "Já vencido"
    if d <= 7:
        return "Até 7 dias"
    if d <= 30:
        return "8 a 30 dias"
    if d <= 60:
        return "31 a 60 dias"
    if d <= 90:
        return "61 a 90 dias"
    return "91 a 120 dias"


df_real["FAIXA DE ANTECIPAÇÃO"] = df_real["DIAS ANTECIPADOS"].map(faixa_ant)

# ----------------------------------------------------------------------------
# tblFaltas — não comparecimento ao exame agendado
# ----------------------------------------------------------------------------
faltas = []
pool = df_aso[df_aso["STATUS"] != "Afastados"].sample(300, random_state=7)
ini, fim = date(2026, 8, 3), date(2026, 10, 21)
span = (fim - ini).days
for k, (_, p) in enumerate(pool.iterrows()):
    vezes = 3 if k < 7 else 2 if k < 46 else 1
    for _ in range(vezes):
        tipo = random.choices(["PERIÓDICO", "DEMISSIONAL", "PERIÓDICO SEMESTRAL", "RETORNO AO TRABALHO"],
                              [74, 16, 7, 3])[0]
        faltas.append({"CPF": p["CPF"], "NOME": p["NOME"], "EMPRESA": p["EMPRESA"],
                       "UNIDADE DE NEGÓCIO": p["UNIDADE DE NEGÓCIO"],
                       "DATA DO AGENDAMENTO": ini + timedelta(days=random.randint(0, span)),
                       "TIPO DE EXAME": tipo,
                       "SITUAÇÃO": "Desligado" if tipo == "DEMISSIONAL" else "Ativo"})
df_faltas = pd.DataFrame(faltas).sort_values("DATA DO AGENDAMENTO")

# ----------------------------------------------------------------------------
# tblDesligados — gerados pelo fluxo de desligamento
# ----------------------------------------------------------------------------
desl = []
for k in range(255):
    unid = random.choices(unid_lista, unid_peso)[0]
    emp = "EMPRESA ALFA" if random.random() < 0.68 else "EMPRESA BETA"
    dt = REF - timedelta(days=random.randint(1, 85))
    while dt.weekday() > 4:
        dt -= timedelta(days=1)
    desl.append({"CPF": cpf_falso(), "NOME": novo_nome(), "CHAPA": f"{'A' if emp == 'EMPRESA ALFA' else 'B'}{9000 + k}",
                 "EMPRESA": emp, "UNIDADE DE NEGÓCIO": unid, "DATA DE DESLIGAMENTO": dt,
                 "ÚLTIMO ASO": dt - timedelta(days=random.randint(30, 360)),
                 "ORIGEM": "Fluxo automático (e-mail de desligamento)",
                 "REGISTRADO EM": dt})
df_desl = pd.DataFrame(desl).sort_values("DATA DE DESLIGAMENTO", ascending=False)

# ----------------------------------------------------------------------------
# Métricas atuais (mesmas regras do dashboard) + tblHistorico semanal
# ----------------------------------------------------------------------------
BASE = {"Ativo", "Férias", "Licença Remunerada"}
FAIXAS_90 = ["Vencido", "Até 7 dias", "8 a 30 dias", "31 a 60 dias", "61 a 90 dias"]


def metricas(df):
    b = df[df["STATUS"].isin(BASE) & ~df["VENCIMENTO"].isin(["EXCEÇÃO DIRETORIA"])].copy()
    b["FX"] = b["VENCIMENTO"].map(faixa_consolidada)
    b = b[b["FX"].isin(FAIXAS_90)]
    ag = b["DATA DE AGENDAMENTO"].notna()
    feito = ag & (b["DATA DE AGENDAMENTO"] <= REF)
    return len(b), int((ag & ~feito).sum()), int(feito.sum()), int((b["FX"] == "Vencido").sum())


hist = []
for u in unid_lista:
    v, a, r, vd = metricas(df_aso[df_aso["UNIDADE DE NEGÓCIO"] == u])
    for w in range(11, -1, -1):
        dt = REF - timedelta(weeks=w)
        f = w / 11
        hist.append({"DATA DO SNAPSHOT": dt, "UNIDADE DE NEGÓCIO": u,
                     "VENCIMENTOS 90 DIAS": max(0, round(v * (1 + 0.12 * f) + random.randint(-2, 2) * (w > 0))),
                     "VENCIDOS": max(0, round(vd * (1 + 1.1 * f) + random.randint(-1, 1) * (w > 0))),
                     "AGENDADOS": max(0, round(a * (1 - 0.45 * f) + random.randint(-2, 2) * (w > 0))),
                     "REALIZADOS": max(0, round(r * (1 - 0.3 * f) + random.randint(-1, 1) * (w > 0)))})
df_hist = pd.DataFrame(hist)

# ----------------------------------------------------------------------------
# Custos de exames — modelo estrela
# ----------------------------------------------------------------------------
EXAMES = [  # exame, grupo, rede, (volA, volB) mensal, preço base
    ("Exame clínico ocupacional", "Clínico", "Própria", (96, 54), 58),
    ("Audiometria tonal", "Audiologia", "Própria", (61, 33), 34),
    ("Audiometria vocal", "Audiologia", "Própria", (61, 0), 22),
    ("Acuidade visual", "Oftalmologia", "Própria", (70, 38), 16),
    ("Avaliação oftalmológica", "Oftalmologia", "Credenciada", (14, 9), 95),
    ("Hemograma completo", "Laboratorial", "Própria", (82, 46), 14),
    ("Glicemia de jejum", "Laboratorial", "Própria", (82, 46), 7),
    ("Gama GT", "Laboratorial", "Própria", (40, 22), 11),
    ("TGO / TGP", "Laboratorial", "Própria", (40, 22), 13),
    ("Creatinina", "Laboratorial", "Própria", (35, 0), 9),
    ("Colesterol total e frações", "Laboratorial", "Própria", (0, 24), 19),
    ("Triglicerídeos", "Laboratorial", "Própria", (0, 24), 10),
    ("Urina tipo I (EAS)", "Laboratorial", "Própria", (30, 18), 9),
    ("Parasitológico de fezes", "Laboratorial", "Credenciada", (12, 7), 12),
    ("Grupo sanguíneo e fator Rh", "Laboratorial", "Própria", (9, 5), 15),
    ("Ácido hipúrico", "Laboratorial", "Credenciada", (0, 6), 48),
    ("Chumbo sérico", "Laboratorial", "Credenciada", (8, 0), 62),
    ("Eletrocardiograma", "Cardiologia", "Própria", (44, 25), 38),
    ("Teste ergométrico", "Cardiologia", "Credenciada", (6, 4), 165),
    ("Eletroencefalograma", "Neurologia", "Credenciada", (22, 13), 88),
    ("Raio-X de tórax PA", "Imagem", "Credenciada", (18, 11), 52),
    ("Raio-X de coluna lombar", "Imagem", "Credenciada", (7, 0), 64),
    ("Espirometria", "Pneumologia", "Própria", (26, 15), 41),
    ("Avaliação psicossocial", "Psicossocial", "Própria", (21, 12), 110),
    ("Exame toxicológico", "Toxicológico", "Credenciada", (15, 8), 145),
    ("Vacina antitetânica", "Imunização", "Própria", (0, 9), 45),
    ("Vacina hepatite B", "Imunização", "Própria", (11, 0), 58),
    ("Avaliação de trabalho em altura", "Clínico", "Própria", (17, 10), 72),
    ("Avaliação de espaço confinado", "Clínico", "Própria", (9, 0), 72),
]
dExame, fVolume, fPreco, fNovo = [], [], [], []
for k, (ex, grp, rede, (va, vb), pb) in enumerate(EXAMES, start=1):
    eid = f"EX{k:03d}"
    dExame.append({"ID_EXAME": eid, "EXAME": ex, "GRUPO": grp, "REDE": rede})
    pa = round(pb * random.uniform(0.92, 1.12), 2)
    pbeta = round(pb * random.uniform(0.95, 1.22), 2)
    if va:
        fVolume.append({"EMPRESA": "EMPRESA ALFA", "ID_EXAME": eid, "VOLUME MENSAL": va})
        fPreco.append({"EMPRESA": "EMPRESA ALFA", "ID_EXAME": eid, "PREÇO ATUAL": pa})
    if vb:
        fVolume.append({"EMPRESA": "EMPRESA BETA", "ID_EXAME": eid, "VOLUME MENSAL": vb})
        fPreco.append({"EMPRESA": "EMPRESA BETA", "ID_EXAME": eid, "PREÇO ATUAL": pbeta})
    precos = [p for p, v in ((pa, va), (pbeta, vb)) if v]
    menor = min(precos)
    if random.random() < 0.22:
        novo, neg = menor, "Sem negociação"
    else:
        novo, neg = round(menor * random.uniform(0.86, 0.97), 2), "Reduzido"
    fNovo.append({"ID_EXAME": eid, "NOVO PREÇO": novo, "SITUAÇÃO DA NEGOCIAÇÃO": neg})

dEmpresa = [
    {"EMPRESA": "EMPRESA ALFA", "VIDAS": 1180, "PER CAPITA ATUAL": 4.90, "PER CAPITA NOVO": 4.20},
    {"EMPRESA": "EMPRESA BETA", "VIDAS": 640, "PER CAPITA ATUAL": 6.40, "PER CAPITA NOVO": 4.20},
]
dPeriodo = [{"PERÍODO": "Mensal", "FATOR": 1, "ORDEM": 1}, {"PERÍODO": "Anual", "FATOR": 12, "ORDEM": 2}]
dCenario = [
    {"CENÁRIO": "Atual (contratos separados)", "ORDEM": 1},
    {"CENÁRIO": "Tabela da Empresa Alfa para todos", "ORDEM": 2},
    {"CENÁRIO": "Tabela da Empresa Beta para todos", "ORDEM": 3},
    {"CENÁRIO": "Menor preço entre as tabelas", "ORDEM": 4},
    {"CENÁRIO": "Novo preço negociado", "ORDEM": 5},
]
premissas_custo = [
    ("Natureza dos dados", "Fictícios, gerados do zero. Não representam empresa, clínica ou contrato reais."),
    ("Custo de exames", "Soma, exame a exame, de volume mensal × preço da tabela aplicada."),
    ("Per capita", "Valor mensal por vida ativa, cobrado à parte dos exames."),
    ("Visão anual", "Visão mensal × 12 (sem sazonalidade)."),
    ("Novo preço", "Preço único da tabela unificada; quando não houve negociação, mantém o menor preço atual."),
    ("Exames exclusivos", "Exames que só uma empresa realiza; no cenário de tabela de outra empresa, mantêm o preço próprio."),
]

# ----------------------------------------------------------------------------
# Gravação dos Excel com tabelas nomeadas
# ----------------------------------------------------------------------------
AZUL = "1F4E5F"


def grava(caminho, abas):
    wb = Workbook()
    wb.remove(wb.active)
    for aba, nome_tab, df in abas:
        ws = wb.create_sheet(aba)
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append([None if (isinstance(v, float) and pd.isna(v)) or v is pd.NaT else v for v in row])
        nlin, ncol = len(df) + 1, len(df.columns)
        ref = f"A1:{get_column_letter(ncol)}{nlin}"
        t = Table(displayName=nome_tab, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)
        for c, col in enumerate(df.columns, start=1):
            letra = get_column_letter(c)
            largura = max(len(str(col)), *(len(str(v)) for v in df[col].head(200))) + 2
            ws.column_dimensions[letra].width = min(largura, 42)
            serie = df[col].dropna()
            if len(serie) and isinstance(serie.iloc[0], date):
                for r in range(2, nlin + 1):
                    ws.cell(r, c).number_format = "dd/mm/yyyy"
            elif "PREÇO" in col or "PER CAPITA" in col:
                for r in range(2, nlin + 1):
                    ws.cell(r, c).number_format = '"R$" #,##0.00'
        ws.freeze_panes = "A2"
    leia = wb.create_sheet("LEIA-ME", 0)
    leia["A1"] = "BASE FICTÍCIA PARA PORTFÓLIO"
    leia["A1"].font = Font(bold=True, size=14, color=AZUL)
    leia["A3"] = "Todos os dados desta pasta foram gerados do zero por script, com semente fixa."
    leia["A4"] = "Nomes, CPFs (começam com 000), CNPJs, e-mails, unidades, preços e volumes são inventados."
    leia["A5"] = f"Data de referência da base: {REF:%d/%m/%Y}."
    leia["A7"] = "Tabelas nesta pasta"
    leia["A7"].font = Font(bold=True)
    for k, (aba, nome_tab, df) in enumerate(abas, start=8):
        leia[f"A{k}"] = f"{nome_tab}  (aba {aba}, {len(df)} linhas)"
    leia.column_dimensions["A"].width = 100
    wb.save(caminho)


grava(OUT / "Base_Controle_ASOs_DEMO.xlsx", [
    ("ASOS", "tblASOS", df_aso),
    ("Realizados", "tblRealizados", df_real),
    ("Faltas", "tblFaltas", df_faltas),
    ("Desligados", "tblDesligados", df_desl),
    ("Histórico", "tblHistorico", df_hist),
])
grava(OUT / "Base_Custos_Exames_DEMO.xlsx", [
    ("dExame", "dExame", pd.DataFrame(dExame)),
    ("dEmpresa", "dEmpresa", pd.DataFrame(dEmpresa)),
    ("fVolume", "fVolume", pd.DataFrame(fVolume)),
    ("fPreco", "fPreco", pd.DataFrame(fPreco)),
    ("fNovoPreco", "fNovoPreco", pd.DataFrame(fNovo)),
    ("dPeriodo", "dPeriodo", pd.DataFrame(dPeriodo)),
    ("dCenario", "dCenario", pd.DataFrame(dCenario)),
    ("Premissas", "tblPremissas", pd.DataFrame(premissas_custo, columns=["ITEM", "REGRA"])),
])

# ----------------------------------------------------------------------------
# JSON compacto para o dashboard web (mesmos dados)
# ----------------------------------------------------------------------------
def d(x):
    return x.isoformat() if isinstance(x, date) else None


web = {
    "ref": REF.isoformat(),
    "unidades": unid_lista,
    "aso": [[r["EMPRESA"][-4:], r["UNIDADE DE NEGÓCIO"], r["STATUS"], r["VENCIMENTO"],
             (r["DATA DE VENCIMENTO"] - REF).days, d(r["DATA DE AGENDAMENTO"]) if pd.notna(r["DATA DE AGENDAMENTO"]) else None]
            for _, r in df_aso.iterrows()],
    "real": [[r["EMPRESA"][-4:], r["UNIDADE DE NEGÓCIO"], r["TIPO DE ASO"], d(r["DATA DE REALIZAÇÃO"]),
              int(r["DIAS ANTECIPADOS"])] for _, r in df_real.iterrows()],
    "faltas": [[r["EMPRESA"][-4:], r["UNIDADE DE NEGÓCIO"], r["NOME"], r["CPF"][-6:], d(r["DATA DO AGENDAMENTO"]),
                r["TIPO DE EXAME"]] for _, r in df_faltas.iterrows()],
    "hist": [[d(r["DATA DO SNAPSHOT"]), r["UNIDADE DE NEGÓCIO"], int(r["VENCIMENTOS 90 DIAS"]), int(r["VENCIDOS"]),
              int(r["AGENDADOS"]), int(r["REALIZADOS"])] for _, r in df_hist.iterrows()],
    "exames": dExame, "volume": fVolume, "preco": fPreco, "novo": fNovo, "empresas": dEmpresa,
}
(OUT / "dados_web.json").write_text(json.dumps(web, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

# resumo para conferência
v, a, r, vd = metricas(df_aso)
print("ASOS", len(df_aso), "| base 90d", v, "agendados", a, "realizados", r, "vencidos", vd)
print("Realizados", len(df_real), "| Faltas", len(df_faltas), "pessoas", df_faltas["CPF"].nunique(),
      "| Desligados", len(df_desl))
vol = pd.DataFrame(fVolume).merge(pd.DataFrame(fPreco))
vol["CUSTO"] = vol["VOLUME MENSAL"] * vol["PREÇO ATUAL"]
print(vol.groupby("EMPRESA")["CUSTO"].sum().round(2).to_dict())
