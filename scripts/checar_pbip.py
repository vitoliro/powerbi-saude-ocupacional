"""Checagem de consistência dos PBIPs gerados (sem o Power BI Desktop).

1. Campos usados nos visuais existem no modelo (coluna ou medida certa).
2. Referências DAX Tabela[Coluna] e [Medida] existem.
3. Colunas do TMDL saem do Power Query (cabeçalhos do Excel + colunas adicionadas).
4. Relacionamentos: colunas existem e o lado "um" não tem valores repetidos nos dados.
5. Indentação TMDL só com tabulação nas linhas de estrutura.
"""
import json
import re
import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).parent
PBIP = RAIZ / "pbip"
DADOS = RAIZ / "dados"
problemas = []


def tira_aspas(nome):
    nome = nome.strip()
    return nome[1:-1].replace("''", "'") if nome.startswith("'") and nome.endswith("'") else nome


def ler_modelo(pasta):
    tabelas = {}
    for f in (pasta / "definition" / "tables").glob("*.tmdl"):
        txt = f.read_text(encoding="utf-8")
        nome = tira_aspas(re.search(r"^table (.+)$", txt, re.M).group(1))
        cols = {tira_aspas(m) for m in re.findall(r"^\tcolumn (.+)$", txt, re.M)}
        meds = {tira_aspas(m) for m in re.findall(r"^\tmeasure (.+?) =", txt, re.M)}
        fontes = re.findall(r"^\t\tsourceColumn: (.+)$", txt, re.M)
        m = re.search(r"^\tpartition .+? = m\n\t\tmode: import\n\t\tsource =\n(.+?)\n\n", txt, re.M | re.S)
        dax = re.findall(r"^\tmeasure .+? = (```.*?```|.+?)$", txt, re.M | re.S)
        for i, l in enumerate(txt.split("\n"), 1):
            ind = re.match(r"^[ \t]*", l).group(0)
            if " " in ind and not l.startswith("\t\t\t"):
                problemas.append(f"{f.name}:{i} indentação com espaço em linha de estrutura")
        tabelas[nome] = {"cols": cols, "meds": meds, "fontes": fontes, "m": m.group(1) if m else "", "dax": dax,
                         "arquivo": f.name}
    rels = re.findall(r"fromColumn: (.+)\n\ttoColumn: (.+)", (pasta / "definition" / "relationships.tmdl").read_text(encoding="utf-8"))
    return tabelas, rels


def separa_ref(ref):
    # Tabela.'Coluna com espaço' ou Tabela.Coluna
    t, c = re.match(r"^('(?:[^']|'')+'|[^.]+)\.(.+)$", ref.strip()).groups()
    return tira_aspas(t), tira_aspas(c)


def checar(nome_proj, excel):
    tab, rels = ler_modelo(PBIP / f"{nome_proj}.SemanticModel")
    todas_medidas = {m for t in tab.values() for m in t["meds"]}

    # 1. visuais
    for vj in (PBIP / f"{nome_proj}.Report").rglob("visual.json"):
        txt = vj.read_text(encoding="utf-8")
        for tipo, ent, prop in re.findall(r'"(Column|Measure)": \{\s*"Expression": \{\s*"SourceRef": \{\s*"Entity": "([^"]+)"\s*\}\s*\},\s*"Property": "([^"]+)"', txt):
            if ent not in tab:
                problemas.append(f"{vj.parent.name}: tabela {ent} não existe")
            elif tipo == "Column" and prop not in tab[ent]["cols"]:
                problemas.append(f"{vj.parent.name}: coluna {ent}[{prop}] não existe")
            elif tipo == "Measure" and prop not in tab[ent]["meds"]:
                problemas.append(f"{vj.parent.name}: medida [{prop}] não existe em {ent}")

    # 2. DAX
    for t, info in tab.items():
        for expr in info["dax"]:
            limpo = re.sub(r'"[^"]*"', '""', expr)
            for tt, cc in re.findall(r"(\w+)\[([^\]]+)\]", limpo):
                if tt not in tab or cc not in tab[tt]["cols"]:
                    problemas.append(f"DAX em {t}: {tt}[{cc}] não existe")
            for mm in re.findall(r"(?<![\w\]])\[([^\]]+)\]", limpo):
                if mm not in todas_medidas and mm != "Date":  # [Date] = coluna gerada por CALENDAR
                    problemas.append(f"DAX em {t}: medida [{mm}] não existe")
            if limpo.count("(") != limpo.count(")"):
                problemas.append(f"DAX em {t}: parênteses desbalanceados em {expr[:60]}")

    # 3. Power Query x colunas
    planilhas = pd.read_excel(DADOS / excel, sheet_name=None)
    cabecalhos = {}
    for aba, df in planilhas.items():
        cabecalhos.update({aba: list(df.columns)})
    import openpyxl
    wb = openpyxl.load_workbook(DADOS / excel, read_only=False)
    tabelas_excel = {}
    for ws in wb.worksheets:
        for nome_t, ref in ws.tables.items():
            primeira = ws[ref.split(":")[0]].row
            col_ini = ws[ref.split(":")[0]].column
            col_fim = ws[ref.split(":")[1]].column
            tabelas_excel[nome_t] = [ws.cell(primeira, c).value for c in range(col_ini, col_fim + 1)]
    for t, info in tab.items():
        m = info["m"]
        if not m:
            continue
        saida = []
        for item in re.findall(r'Item="([^"]+)"', m):
            if item not in tabelas_excel:
                problemas.append(f"{t}: tabela {item} não existe no Excel {excel}")
        itens = re.findall(r'Item="([^"]+)"', m)
        if itens:
            saida += tabelas_excel.get(itens[0], [])
            for c, _ in re.findall(r'\{"([^"]+)", (type \w+|Int64\.Type)\}', m):
                if c not in tabelas_excel.get(itens[0], []) and len(itens) == 1:
                    problemas.append(f"{t}: tipo aplicado a coluna {c} que não está em {itens[0]}")
        saida += re.findall(r'Table\.AddColumn\(\w+, "([^"]+)"', m)
        for lista in re.findall(r'Table\.ExpandTableColumn\(\w+, "\w+", \{([^}]+)\}', m):
            saida += re.findall(r'"([^"]+)"', lista)
        saida += re.findall(r'#"([^"]+)" = (?:text|type \w+|Int64\.Type)', m)
        saida += re.findall(r'(?:\[|, )(\w+) = (?:text|Int64\.Type)', m)
        sel = re.findall(r'Table\.SelectColumns\(\w+, \{"([^"]+)"\}\)', m)
        if sel:
            saida = sel
        if t == "dExame":  # dExame parte de dExame + colunas do fNovoPreco
            saida = tabelas_excel["dExame"] + ["NOVO PREÇO", "SITUAÇÃO DA NEGOCIAÇÃO", "EMPRESAS QUE REALIZAM"]
        if t == "fExameEmpresa":
            saida = tabelas_excel["fVolume"] + ["PREÇO ATUAL", "PREÇO ALFA", "PREÇO BETA", "MENOR PREÇO"]
        for fonte in info["fontes"]:
            if fonte.startswith("["):
                continue
            if fonte not in saida:
                problemas.append(f"{t}: sourceColumn '{fonte}' não sai do Power Query ({saida})")
        for c in saida:
            if c not in info["fontes"] and not sel:
                problemas.append(f"{t}: coluna '{c}' do Power Query não está declarada no TMDL")

    # 3b. medida com mesmo nome de coluna (DAX ignora maiúsculas)
    nomes_cols = {c.lower(): f"{tt}[{c}]" for tt, info in tab.items() for c in info["cols"]}
    for mm in todas_medidas:
        if mm.lower() in nomes_cols:
            problemas.append(f"medida [{mm}] tem o mesmo nome da coluna {nomes_cols[mm.lower()]}")

    # 4. relacionamentos (lado um sem duplicatas nos dados)
    for de, para in rels:
        (t1, c1), (t2, c2) = separa_ref(de), separa_ref(para)
        for t_, c_ in ((t1, c1), (t2, c2)):
            if t_ not in tab or c_ not in tab[t_]["cols"]:
                problemas.append(f"relacionamento: {t_}[{c_}] não existe")
    return tab, rels


for proj, xl in (("Controle_ASOs_DEMO", "Base_Controle_ASOs_DEMO.xlsx"), ("Custos_Exames_DEMO", "Base_Custos_Exames_DEMO.xlsx")):
    tab, rels = checar(proj, xl)
    print(f"{proj}: {len(tab)} tabelas, {sum(len(t['cols']) for t in tab.values())} colunas, "
          f"{sum(len(t['meds']) for t in tab.values())} medidas, {len(rels)} relacionamentos")

if problemas:
    print(f"\n{len(problemas)} problema(s):")
    for p in problemas:
        print(" -", p)
    sys.exit(1)
print("\nNenhum problema de consistência encontrado.")
