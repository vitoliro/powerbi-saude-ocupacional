"""Gera dois projetos Power BI (PBIP: modelo TMDL + relatório PBIR) com dados fictícios.

Saída: portfolio/pbip/
  Controle_ASOs_DEMO.pbip  (+ .SemanticModel / .Report)
  Custos_Exames_DEMO.pbip  (+ .SemanticModel / .Report)

As fontes são os Excel fictícios em portfolio/dados. O caminho da pasta fica no
parâmetro PastaDados do Power Query (editável no Power BI Desktop).
Formato espelha um PBIP salvo pelo Power BI Desktop (schemas de 2024-2025).
"""
import json
import re
import shutil
import uuid
from pathlib import Path

RAIZ = Path(__file__).parent
OUT = RAIZ / "pbip"
BASE_THEME = RAIZ / "CY24SU10.json"
PASTA_PADRAO = "C:\\Portfolio-SSO\\dados\\"

SCH = "https://developer.microsoft.com/json-schemas/fabric"
V_VISUAL = f"{SCH}/item/report/definition/visualContainer/1.5.0/schema.json"
V_PAGE = f"{SCH}/item/report/definition/page/1.3.0/schema.json"
V_PAGES = f"{SCH}/item/report/definition/pagesMetadata/1.0.0/schema.json"
V_REPORT = f"{SCH}/item/report/definition/report/1.2.0/schema.json"
V_VERSION = f"{SCH}/item/report/definition/versionMetadata/1.0.0/schema.json"
V_PLATFORM = f"{SCH}/gitIntegration/platformProperties/2.0.0/schema.json"
V_PBIP = f"{SCH}/pbip/pbipProperties/1.0.0/schema.json"
V_PBIR = f"{SCH}/item/report/definitionProperties/2.0.0/schema.json"
V_PBISM = f"{SCH}/item/semanticModel/definitionProperties/1.0.0/schema.json"

# Paleta própria (neutra, não corporativa)
COR = {
    "tinta": "#22313A", "suave": "#5B6B75", "fundo": "#F3F5F6", "borda": "#E1E6E9",
    "petroleo": "#1F5F7A", "petroleo_claro": "#8DBBCD", "cinza": "#C9D1D6",
    "vermelho": "#B9473A", "ambar": "#D99A3D", "verde": "#3F7D5A",
}
T = "\t"


def w(path: Path, conteudo, is_json=True):
    path.parent.mkdir(parents=True, exist_ok=True)
    if is_json:
        path.write_text(json.dumps(conteudo, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        path.write_text(conteudo, encoding="utf-8")


def gid():
    return str(uuid.uuid4())


# =============================================================================
# TMDL (modelo semântico)
# =============================================================================
def q(nome):
    """Nome TMDL: aspas simples quando há espaço ou caractere especial."""
    simples = nome.replace("_", "").isalnum() and nome.isascii()
    return nome if simples else "'" + nome.replace("'", "''") + "'"


TIPOS = {"text": ("string", "type text"), "date": ("dateTime", "type date"),
         "int": ("int64", "Int64.Type"), "num": ("double", "type number"), "bool": ("boolean", "type logical")}


def coluna(nome, tipo, hidden=False, fmt=None, sort_by=None, summarize="none"):
    dt = TIPOS[tipo][0]
    linhas = [f"{T}column {q(nome)}", f"{T*2}dataType: {dt}"]
    if fmt:
        linhas.append(f"{T*2}formatString: {fmt}")
    elif tipo == "date":
        linhas.append(f"{T*2}formatString: Short Date")
    elif tipo == "int":
        linhas.append(f"{T*2}formatString: 0")
    if hidden:
        linhas.append(f"{T*2}isHidden")
    linhas.append(f"{T*2}summarizeBy: {summarize}")
    linhas.append(f"{T*2}sourceColumn: {nome}")
    if sort_by:
        linhas.append(f"{T*2}sortByColumn: {q(sort_by)}")
    linhas.append("")
    if tipo == "date":
        linhas.append(f"{T*2}annotation UnderlyingDateTimeDataType = Date")
        linhas.append("")
    linhas.append(f"{T*2}annotation SummarizationSetBy = Automatic")
    linhas.append("")
    return "\n".join(linhas)


def medida(nome, dax, fmt, pasta=None, descricao=None):
    linhas = []
    if descricao:
        linhas.append(f"{T}/// {descricao}")
    corpo = dax.strip("\n")
    if "\n" in corpo:
        linhas.append(f"{T}measure {q(nome)} = ```")
        for l in corpo.split("\n"):
            linhas.append(f"{T*3}{l}" if l.strip() else "")
        linhas.append(f"{T*3}```")
    else:
        linhas.append(f"{T}measure {q(nome)} = {corpo}")
    linhas.append(f"{T*2}formatString: {fmt}")
    if pasta:
        linhas.append(f"{T*2}displayFolder: {pasta}")
    linhas.append("")
    return "\n".join(linhas)


def particao_m(nome, m):
    linhas = [f"{T}partition {q(nome)} = m", f"{T*2}mode: import", f"{T*2}source ="]
    for l in m.strip("\n").split("\n"):
        linhas.append(f"{T*4}{l}" if l.strip() else "")
    linhas.append("")
    return "\n".join(linhas)


def tabela(nome, colunas, m, medidas_txt="", descricao=None):
    partes = []
    if descricao:
        partes.append(f"/// {descricao}")
    partes.append(f"table {q(nome)}")
    partes.append("")
    if medidas_txt:
        partes.append(medidas_txt)
    for c in colunas:
        partes.append(coluna(*c) if isinstance(c, tuple) else coluna(**c))
    partes.append(particao_m(nome, m))
    partes.append(f"{T}annotation PBI_ResultType = Table")
    partes.append("")
    return "\n".join(partes)


def tabela_medidas(nome, medidas_txt):
    return "\n".join([
        f"/// Tabela só de medidas",
        f"table {q(nome)}",
        "",
        medidas_txt,
        f"{T}column Value",
        f"{T*2}isHidden",
        f"{T*2}formatString: 0",
        f"{T*2}summarizeBy: sum",
        f"{T*2}isNameInferred",
        f"{T*2}sourceColumn: [Value]",
        "",
        f"{T*2}annotation SummarizationSetBy = Automatic",
        "",
        f"{T}partition {q(nome)} = calculated",
        f"{T*2}mode: import",
        f"{T*2}source = {{0}}",
        "",
        f"{T}annotation PBI_Id = {uuid.uuid4().hex}",
        "",
    ])


def m_excel(arquivo, item, tipos, extra=""):
    """M que lê uma tabela nomeada do Excel fictício e aplica tipos."""
    tl = ", ".join(f'{{"{c}", {TIPOS[t][1]}}}' for c, t in tipos)
    ultimo = "Tipos"
    passos = [
        "let",
        f'    Fonte = Excel.Workbook(File.Contents(PastaDados & "{arquivo}"), null, true),',
        f'    Tabela = Fonte{{[Item="{item}",Kind="Table"]}}[Data],',
        f"    Tipos = Table.TransformColumnTypes(Tabela, {{{tl}}})",
    ]
    if extra:
        passos[-1] += ","
        extra_linhas = extra.strip("\n").split("\n")
        passos += extra_linhas
        ultimo = [re.match(r"^    (\w+) = ", l).group(1) for l in extra_linhas if re.match(r"^    (\w+) = ", l)][-1]
    passos += ["in", f"    {ultimo}"]
    return "\n".join(passos)


def expressoes(params):
    out = []
    for nome, valor, tipo in params:
        out.append(f'expression {nome} = {valor} meta [IsParameterQuery=true, Type="{tipo}", IsParameterQueryRequired=true]')
        out.append("")
        out.append(f"{T}annotation PBI_ResultType = {tipo}")
        out.append("")
    return "\n".join(out)


def relacionamentos(rels):
    out = []
    for de, para in rels:
        (t1, c1), (t2, c2) = de, para
        out.append(f"relationship {gid()}")
        out.append(f"{T}fromColumn: {q(t1)}.{q(c1)}")
        out.append(f"{T}toColumn: {q(t2)}.{q(c2)}")
        out.append("")
    return "\n".join(out)


def modelo(nome_proj, tabelas_txt, ordem_queries, exprs_txt, rels_txt):
    base = OUT / f"{nome_proj}.SemanticModel"
    w(base / "definition.pbism", {"$schema": V_PBISM, "version": "4.1", "settings": {}})
    w(base / ".platform", {"$schema": V_PLATFORM, "metadata": {"type": "SemanticModel", "displayName": nome_proj},
                           "config": {"version": "2.0", "logicalId": gid()}})
    w(base / ".pbi" / "editorSettings.json", {"version": "1.0", "autodetectRelationships": False,
                                               "parallelQueryLoading": True, "relationshipImportEnabled": True,
                                               "shouldNotifyUserOfNameConflictResolution": True})
    d = base / "definition"
    w(d / "database.tmdl", "database\n\tcompatibilityLevel: 1567\n\n", False)
    model = [
        "model Model",
        f"{T}culture: pt-BR",
        f"{T}defaultPowerBIDataSourceVersion: powerBI_V3",
        f"{T}discourageImplicitMeasures",
        f"{T}sourceQueryCulture: pt-BR",
        f"{T}dataAccessOptions",
        f"{T*2}legacyRedirects",
        f"{T*2}returnErrorValuesAsNull",
        "",
        "annotation __PBI_TimeIntelligenceEnabled = 0",
        "",
        "annotation PBI_QueryOrder = " + json.dumps(ordem_queries, ensure_ascii=False),
        "",
        'annotation PBI_ProTooling = ["DevMode"]',
        "",
    ]
    model += [f"ref table {q(n)}" for n in tabelas_txt]
    model.append("")
    w(d / "model.tmdl", "\n".join(model), False)
    w(d / "expressions.tmdl", exprs_txt, False)
    w(d / "relationships.tmdl", rels_txt, False)
    for n, txt in tabelas_txt.items():
        w(d / "tables" / f"{n}.tmdl", txt, False)


# =============================================================================
# PBIR (relatório)
# =============================================================================
def L(v):
    return {"expr": {"Literal": {"Value": v}}}


def b(x):
    return L("true" if x else "false")


def n(x):
    return L(f"{x}D")


def s(x):
    return L(f"'{x}'")


def cor(h):
    return {"solid": {"color": L(f"'{h}'")}}


def props(**kw):
    return [{"properties": kw}]


def col(tab, c):
    return {"Column": {"Expression": {"SourceRef": {"Entity": tab}}, "Property": c}}


def med(tab, m):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": tab}}, "Property": m}}


def proj(campo, nome_exib=None):
    k = "Column" if "Column" in campo else "Measure"
    tab = campo[k]["Expression"]["SourceRef"]["Entity"]
    p = campo[k]["Property"]
    out = {"field": campo, "queryRef": f"{tab}.{p}", "nativeQueryRef": nome_exib or p}
    if nome_exib:
        out["displayName"] = nome_exib
    return out


def container_padrao(titulo=None, mostrar_fundo=True):
    vco = {
        "background": props(show=b(mostrar_fundo), color=cor("#FFFFFF"), transparency=n(0)),
        "border": props(show=b(mostrar_fundo), color=cor(COR["borda"]), radius=n(10), width=n(1)),
        "padding": props(top=n(12), bottom=n(10), left=n(14), right=n(14)),
        "dropShadow": props(show=b(False)),
    }
    if titulo:
        vco["title"] = props(show=b(True), text=s(titulo), fontColor=cor(COR["tinta"]), fontSize=n(12),
                             fontFamily=s("Segoe UI Semibold"))
    else:
        vco["title"] = props(show=b(False))
    return vco


class Pagina:
    def __init__(self, nome, exibicao):
        self.nome, self.exibicao, self.visuais = nome, exibicao, []

    def add(self, nome, x, y, largura, altura, visual):
        z = len(self.visuais) * 1000
        self.visuais.append({
            "$schema": V_VISUAL, "name": nome,
            "position": {"x": x, "y": y, "z": z, "height": altura, "width": largura, "tabOrder": z},
            "visual": visual,
        })


def v_texto(paragrafos):
    """paragrafos: lista de [(texto, tamanho_pt, cor, negrito)]"""
    return {
        "visualType": "textbox",
        "objects": {"general": [{"properties": {"paragraphs": [
            {"textRuns": [{"value": t, "textStyle": {"fontSize": f"{tam}pt", "color": c,
                                                       **({"fontWeight": "bold"} if neg else {}),
                                                       "fontFamily": "Segoe UI"}} for t, tam, c, neg in par]}
            for par in paragrafos]}}]},
        "visualContainerObjects": {"background": props(show=b(False)), "border": props(show=b(False)),
                                   "padding": props(top=n(0), bottom=n(0), left=n(0), right=n(0)),
                                   "title": props(show=b(False))},
        "drillFilterOtherVisuals": True,
    }


def v_slicer(tab, c, rotulo, unico=False):
    objs = {"data": props(mode=s("Dropdown")),
            "header": props(show=b(True), fontColor=cor(COR["suave"]), textSize=n(9)),
            "items": props(fontColor=cor(COR["tinta"]), textSize=n(10))}
    if unico:
        objs["selection"] = props(strictSingleSelect=b(True))
    return {
        "visualType": "slicer",
        "query": {"queryState": {"Values": {"projections": [{**proj(col(tab, c), rotulo), "active": True}]}},
                  "sortDefinition": {"sort": [{"field": col(tab, c), "direction": "Ascending"}], "isDefaultSort": True}},
        "objects": objs,
        "visualContainerObjects": {**container_padrao(), "padding": props(top=n(4), bottom=n(4), left=n(10), right=n(10))},
        "drillFilterOtherVisuals": True,
    }


def v_card(tab_med, m, rotulo, cor_valor=None):
    return {
        "visualType": "card",
        "query": {"queryState": {"Values": {"projections": [proj(med(tab_med, m), rotulo)]}}},
        "objects": {
            "labels": props(color=cor(cor_valor or COR["petroleo"]), fontSize=n(26), fontFamily=s("Segoe UI Semibold")),
            "categoryLabels": props(show=b(True), color=cor(COR["suave"]), fontSize=n(10)),
        },
        "visualContainerObjects": container_padrao(),
        "drillFilterOtherVisuals": True,
    }


def seletor_igual(tab, c, valor):
    return {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": col(tab, c),
                                                  "Right": {"Literal": {"Value": f"'{valor}'"}}}}}]}


def v_barras(tipo, categoria, valores, titulo, serie=None, cores_serie=None, cor_unica=None,
             ordenar_por=None, desc=True, legenda=True, rotulos=True):
    qs = {"Category": {"projections": [{**proj(categoria), "active": True}]},
          "Y": {"projections": [proj(v) for v in valores]}}
    if serie:
        qs["Series"] = {"projections": [proj(serie)]}
    query = {"queryState": qs}
    if ordenar_por is not None:
        query["sortDefinition"] = {"sort": [{"field": ordenar_por, "direction": "Descending" if desc else "Ascending"}]}
    objs = {
        "labels": props(show=b(rotulos), color=cor(COR["tinta"]), fontSize=n(9)),
        "valueAxis": props(show=b(False), gridlineShow=b(False)),
        "categoryAxis": props(showAxisTitle=b(False), labelColor=cor(COR["suave"]), fontSize=n(9)),
        "legend": props(show=b(bool(serie) and legenda), position=s("TopRight"), showTitle=b(False),
                        labelColor=cor(COR["suave"]), fontSize=n(9)),
    }
    if tipo in ("columnChart", "lineChart"):
        objs["valueAxis"] = props(show=b(False), gridlineShow=b(False))
    dp = []
    if cor_unica:
        dp.append({"properties": {"fill": cor(cor_unica)}})
    if serie and cores_serie:
        k = "Column"
        t, c = serie[k]["Expression"]["SourceRef"]["Entity"], serie[k]["Property"]
        for valor, h in cores_serie.items():
            dp.append({"properties": {"fill": cor(h)}, "selector": seletor_igual(t, c, valor)})
    if dp:
        objs["dataPoint"] = dp
    if tipo == "lineChart":
        objs["lineStyles"] = props(strokeWidth=n(3), showMarker=b(True), markerSize=n(5))
        objs["markers"] = props(transparency=n(0))
        objs.pop("dataPoint", None)
        objs["dataPoint"] = [{"properties": {"fill": cor(cor_unica or COR["petroleo"])}}]
    return {"visualType": tipo, "query": query, "objects": objs,
            "visualContainerObjects": container_padrao(titulo), "drillFilterOtherVisuals": True}


def v_tabela(campos, titulo, ordenar_por=None, desc=True):
    query = {"queryState": {"Values": {"projections": [proj(c, nome) for c, nome in campos]}}}
    if ordenar_por is not None:
        query["sortDefinition"] = {"sort": [{"field": ordenar_por, "direction": "Descending" if desc else "Ascending"}]}
    return {
        "visualType": "tableEx", "query": query,
        "objects": {
            "columnHeaders": props(fontColor=cor(COR["suave"]), backColor=cor("#FFFFFF"), fontSize=n(9), bold=b(True)),
            "values": props(fontColor=cor(COR["tinta"]), fontSize=n(9), backColorPrimary=cor("#FFFFFF"),
                            backColorSecondary=cor("#F7F9FA")),
            "grid": props(gridHorizontal=b(True), gridHorizontalColor=cor(COR["borda"]), gridVertical=b(False)),
        },
        "visualContainerObjects": container_padrao(titulo), "drillFilterOtherVisuals": True,
    }


def v_cascata(categoria, valor, titulo):
    return {
        "visualType": "waterfallChart",
        "query": {"queryState": {"Category": {"projections": [{**proj(categoria), "active": True}]},
                                 "Y": {"projections": [proj(valor)]}}},
        "objects": {
            "sentimentColors": props(increaseFill=cor("#8A99A3"), decreaseFill=cor(COR["petroleo"]),
                                     totalFill=cor(COR["tinta"])),
            "labels": props(show=b(True), color=cor(COR["tinta"]), fontSize=n(9)),
            "valueAxis": props(show=b(False), gridlineShow=b(False)),
            "categoryAxis": props(showAxisTitle=b(False), labelColor=cor(COR["suave"]), fontSize=n(9)),
            "legend": props(show=b(False)),
        },
        "visualContainerObjects": container_padrao(titulo), "drillFilterOtherVisuals": True,
    }


def tema(nome):
    return {
        "name": nome,
        "dataColors": [COR["petroleo"], COR["petroleo_claro"], COR["cinza"], COR["vermelho"], COR["ambar"],
                       COR["verde"], "#6C5B7B", "#355C7D"],
        "foreground": COR["tinta"], "foregroundNeutralSecondary": COR["suave"], "foregroundNeutralTertiary": "#8A99A3",
        "background": "#FFFFFF", "backgroundLight": COR["fundo"], "backgroundNeutral": COR["borda"],
        "tableAccent": COR["petroleo"], "good": COR["verde"], "neutral": COR["ambar"], "bad": COR["vermelho"],
        "maximum": COR["petroleo"], "center": COR["petroleo_claro"], "minimum": "#EEF3F5",
        "textClasses": {
            "callout": {"fontSize": 26, "fontFace": "Segoe UI Semibold", "color": COR["petroleo"]},
            "title": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": COR["tinta"]},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": COR["tinta"]},
            "label": {"fontSize": 10, "fontFace": "Segoe UI", "color": COR["suave"]},
        },
        "visualStyles": {
            "page": {"*": {"background": [{"color": {"solid": {"color": COR["fundo"]}}, "transparency": 0}],
                           "outspace": [{"color": {"solid": {"color": COR["fundo"]}}, "transparency": 0}]}},
        },
    }


def relatorio(nome_proj, paginas, nome_tema):
    base = OUT / f"{nome_proj}.Report"
    w(base / "definition.pbir", {"$schema": V_PBIR, "version": "4.0", "datasetReference": {"byPath": {"path": f"../{nome_proj}.SemanticModel"}}})
    w(base / ".platform", {"$schema": V_PLATFORM, "metadata": {"type": "Report", "displayName": nome_proj},
                           "config": {"version": "2.0", "logicalId": gid()}})
    d = base / "definition"
    w(d / "version.json", {"$schema": V_VERSION, "version": "2.0.0"})
    arq_tema = f"{nome_tema}-{uuid.uuid4().hex[:8]}.json"
    w(d / "report.json", {
        "$schema": V_REPORT,
        "themeCollection": {
            "baseTheme": {"name": "CY24SU10", "reportVersionAtImport": "5.61", "type": "SharedResources"},
            "customTheme": {"name": arq_tema, "reportVersionAtImport": "5.61", "type": "RegisteredResources"},
        },
        "layoutOptimization": "None",
        "objects": {"section": props(verticalAlignment=s("Top")),
                    "outspacePane": props(visible=b(True), expanded=b(False))},
        "resourcePackages": [
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": arq_tema, "path": arq_tema, "type": "CustomTheme"}]},
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": "CY24SU10", "path": "BaseThemes/CY24SU10.json", "type": "BaseTheme"}]},
        ],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True, "useDefaultAggregateDisplayName": True},
    })
    w(base / "StaticResources" / "RegisteredResources" / arq_tema, tema(arq_tema))
    destino_base = base / "StaticResources" / "SharedResources" / "BaseThemes" / "CY24SU10.json"
    destino_base.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(BASE_THEME, destino_base)
    w(d / "pages" / "pages.json", {"$schema": V_PAGES, "pageOrder": [p.nome for p in paginas],
                                   "activePageName": paginas[0].nome})
    for p in paginas:
        w(d / "pages" / p.nome / "page.json", {"$schema": V_PAGE, "name": p.nome, "displayName": p.exibicao,
                                               "displayOption": "FitToPage", "height": 720, "width": 1280})
        for v in p.visuais:
            w(d / "pages" / p.nome / "visuals" / v["name"] / "visual.json", v)
    w(OUT / f"{nome_proj}.pbip", {"$schema": V_PBIP, "version": "1.0", "artifacts": [{"report": {"path": f"{nome_proj}.Report"}}],
                                  "settings": {"enableAutoRecovery": True}})


def cabecalho(p, titulo, subtitulo, slicers):
    p.add("titulo", 24, 14, 720, 40, v_texto([[(titulo, 20, COR["tinta"], True)]]))
    p.add("subtitulo", 24, 52, 720, 24, v_texto([[(subtitulo, 10, COR["suave"], False)]]))
    x = 1256
    for nome, visual in reversed(slicers):
        x -= 200
        p.add(nome, x, 10, 192, 76, visual)
        x -= 8


def linha_cards(p, cards, y=96, h=92):
    n_ = len(cards)
    largura = (1232 - 12 * (n_ - 1)) / n_
    for i, (nome, visual) in enumerate(cards):
        p.add(nome, round(24 + i * (largura + 12)), y, round(largura), h, visual)


# =============================================================================
# PROJETO 1 — CONTROLE DE ASOs
# =============================================================================
def projeto_asos():
    nome = "Controle_ASOs_DEMO"
    arq = "Base_Controle_ASOs_DEMO.xlsx"
    M = "Medidas"
    UN = "UNIDADE DE NEGÓCIO"

    t_asos = [("CPF", "text"), ("NOME", "text"), ("CHAPA", "text"), ("EMPRESA", "text"), ("CNPJ", "text"),
              ("POSTO DE TRABALHO - GHE", "text"), ("FUNÇÃO", "text"), (UN, "text"), ("ADMISSÃO", "date"),
              ("E-MAIL", "text"), ("GESTOR DIRETO", "text"), ("VENCIMENTO", "text"), ("DATA DE VENCIMENTO", "date"),
              ("STATUS", "text"), ("STATUS FINAL", "text"), ("ÚLTIMO ASO", "date"), ("DATA DE AGENDAMENTO", "date"),
              ("TIPO DE ASO", "text")]
    extra_asos = """    DiasParaVencer = Table.AddColumn(Tipos, "DIAS PARA VENCER", each Duration.Days([DATA DE VENCIMENTO] - DataReferencia), Int64.Type),
    NaBase = Table.AddColumn(DiasParaVencer, "NA BASE", each List.Contains({"Ativo", "Férias", "Licença Remunerada"}, [STATUS]) and [VENCIMENTO] <> "EXCEÇÃO DIRETORIA", type logical),
    Faixa = Table.AddColumn(NaBase, "FAIXA 90D", each
        if [DIAS PARA VENCER] < 0 then "Vencido"
        else if [DIAS PARA VENCER] <= 7 then "Até 7 dias"
        else if [DIAS PARA VENCER] <= 30 then "8 a 30 dias"
        else if [DIAS PARA VENCER] <= 60 then "31 a 60 dias"
        else if [DIAS PARA VENCER] <= 90 then "61 a 90 dias"
        else "Acima de 90 dias", type text),
    Situacao = Table.AddColumn(Faixa, "SITUAÇÃO AGENDAMENTO", each
        if [DATA DE AGENDAMENTO] = null then "Sem agendamento"
        else if [DATA DE AGENDAMENTO] <= DataReferencia then "Realizado"
        else "Agendado", type text),
    AntecipacaoAgendado = Table.AddColumn(Situacao, "DIAS ANTECIPAÇÃO AGENDADO", each
        if [SITUAÇÃO AGENDAMENTO] = "Agendado" then Duration.Days([DATA DE VENCIMENTO] - [DATA DE AGENDAMENTO]) else null, Int64.Type)"""
    cols_asos = [(c, t) for c, t in t_asos] + [("DIAS PARA VENCER", "int"), ("NA BASE", "bool"), ("FAIXA 90D", "text"),
                                                ("SITUAÇÃO AGENDAMENTO", "text"), ("DIAS ANTECIPAÇÃO AGENDADO", "int")]

    t_real = [("CPF", "text"), ("NOME", "text"), ("EMPRESA", "text"), (UN, "text"), ("TIPO DE ASO", "text"),
              ("DATA DE REALIZAÇÃO", "date"), ("VENCIMENTO ANTERIOR", "date"), ("DIAS ANTECIPADOS", "int"),
              ("FAIXA DE ANTECIPAÇÃO", "text")]
    extra_mes = lambda base, campo: f"""    Mes = Table.AddColumn({base}, "MÊS", each Date.ToText([{campo}], "MMM/yy", "pt-BR"), type text),
    MesOrdem = Table.AddColumn(Mes, "MÊS ORDEM", each Date.Year([{campo}]) * 100 + Date.Month([{campo}]), Int64.Type)"""
    cols_real = [(c, t) for c, t in t_real] + [("MÊS", "text", False, None, "MÊS ORDEM"), ("MÊS ORDEM", "int", True)]

    t_faltas = [("CPF", "text"), ("NOME", "text"), ("EMPRESA", "text"), (UN, "text"), ("DATA DO AGENDAMENTO", "date"),
                ("TIPO DE EXAME", "text"), ("SITUAÇÃO", "text")]
    cols_faltas = [(c, t) for c, t in t_faltas] + [("MÊS", "text", False, None, "MÊS ORDEM"), ("MÊS ORDEM", "int", True)]

    t_desl = [("CPF", "text"), ("NOME", "text"), ("CHAPA", "text"), ("EMPRESA", "text"), (UN, "text"),
              ("DATA DE DESLIGAMENTO", "date"), ("ÚLTIMO ASO", "date"), ("ORIGEM", "text"), ("REGISTRADO EM", "date")]
    t_hist = [("DATA DO SNAPSHOT", "date"), (UN, "text"), ("VENCIMENTOS 90 DIAS", "int"), ("VENCIDOS", "int"),
              ("AGENDADOS", "int"), ("REALIZADOS", "int")]

    def ordem_tab(nome_col, valores):
        linhas = ", ".join(f'{{"{v}", {i}}}' for i, v in enumerate(valores, start=1))
        return f"""let
    Fonte = #table(type table [#"{nome_col}" = text, ORDEM = Int64.Type], {{{linhas}}})
in
    Fonte"""

    medidas_txt = "\n".join([
        medida("Colaboradores na Base", "CALCULATE ( COUNTROWS ( tblASOS ), tblASOS[NA BASE] = TRUE () )", "#,0",
               "1. Vencimentos", "Ativos, em férias ou licença remunerada; afastados e exceções ficam fora."),
        medida("Vencimentos 90d", """CALCULATE (
    COUNTROWS ( tblASOS ),
    tblASOS[NA BASE] = TRUE (),
    tblASOS[DIAS PARA VENCER] <= 90
)""", "#,0", "1. Vencimentos", "ASOs vencidos ou que vencem nos próximos 90 dias, na base de cálculo."),
        medida("ASOs Vencidos", 'CALCULATE ( [Vencimentos 90d], tblASOS[FAIXA 90D] = "Vencido" )', "#,0", "1. Vencimentos"),
        medida("ASOs Agendados", "CALCULATE ( [Vencimentos 90d], tblASOS[SITUAÇÃO AGENDAMENTO] = \"Agendado\" )", "#,0",
               "1. Vencimentos", "Exame com data futura."),
        medida("Realizados a Atualizar", "CALCULATE ( [Vencimentos 90d], tblASOS[SITUAÇÃO AGENDAMENTO] = \"Realizado\" )",
               "#,0", "1. Vencimentos", "Exame já feito, ASO ainda não lançado na base."),
        medida("Sem Agendamento", "CALCULATE ( [Vencimentos 90d], tblASOS[SITUAÇÃO AGENDAMENTO] = \"Sem agendamento\" )",
               "#,0", "1. Vencimentos"),
        medida("% Cobertura", "DIVIDE ( [ASOs Agendados] + [Realizados a Atualizar], [Vencimentos 90d] )", "0.0%",
               "1. Vencimentos", "Parte dos vencimentos em 90 dias que já tem exame agendado ou realizado."),
        medida("Vencidos sem Agendamento", 'CALCULATE ( [Sem Agendamento], tblASOS[FAIXA 90D] = "Vencido" )', "#,0",
               "1. Vencimentos"),
        medida("Vencidos (semanal)", "SUM ( tblHistorico[VENCIDOS] )", "#,0", "2. Histórico"),
        medida("% Cobertura (semanal)", """DIVIDE (
    SUM ( tblHistorico[AGENDADOS] ) + SUM ( tblHistorico[REALIZADOS] ),
    SUM ( tblHistorico[VENCIMENTOS 90 DIAS] )
)""", "0.0%", "2. Histórico"),
        medida("ASOs Realizados", "COUNTROWS ( tblRealizados )", "#,0", "3. Antecipação"),
        medida("Antecipação Média (dias)", "AVERAGE ( tblRealizados[DIAS ANTECIPADOS] )", "#,0", "3. Antecipação",
               "Dias entre o exame e o vencimento anterior; negativo quando já estava vencido."),
        medida("Realizados Já Vencidos", "CALCULATE ( COUNTROWS ( tblRealizados ), tblRealizados[DIAS ANTECIPADOS] < 0 )",
               "#,0", "3. Antecipação"),
        medida("% Realizado Já Vencido", "DIVIDE ( [Realizados Já Vencidos], [ASOs Realizados] )", "0.0%", "3. Antecipação"),
        medida("Antecipação Agendados (dias)",
               "CALCULATE ( AVERAGE ( tblASOS[DIAS ANTECIPAÇÃO AGENDADO] ), tblASOS[NA BASE] = TRUE () )", "#,0",
               "3. Antecipação", "Dias entre a data agendada e o vencimento, para exames com data futura."),
        medida("Total de Faltas", "COUNTROWS ( tblFaltas )", "#,0", "4. Faltas"),
        medida("Pessoas que Faltaram", "DISTINCTCOUNT ( tblFaltas[CPF] )", "#,0", "4. Faltas"),
        medida("Reincidentes 2+", """COUNTROWS (
    FILTER (
        VALUES ( tblFaltas[CPF] ),
        CALCULATE ( COUNTROWS ( tblFaltas ) ) >= 2
    )
)""", "#,0", "4. Faltas", "Colaboradores com duas faltas ou mais no período."),
        medida("% Reincidentes", "DIVIDE ( [Reincidentes 2+], [Pessoas que Faltaram] )", "0.0%", "4. Faltas"),
        medida("Faltas em Demissional", 'CALCULATE ( [Total de Faltas], tblFaltas[TIPO DE EXAME] = "DEMISSIONAL" )', "#,0",
               "4. Faltas"),
        medida("Desligamentos Registrados", "COUNTROWS ( tblDesligados )", "#,0", "5. Desligamentos"),
    ])

    tabs = {
        "tblASOS": tabela("tblASOS", cols_asos, m_excel(arq, "tblASOS", t_asos, extra_asos),
                          descricao="Um colaborador por linha (base fictícia)."),
        "tblRealizados": tabela("tblRealizados", cols_real,
                                m_excel(arq, "tblRealizados", t_real, extra_mes("Tipos", "DATA DE REALIZAÇÃO"))),
        "tblFaltas": tabela("tblFaltas", cols_faltas,
                            m_excel(arq, "tblFaltas", t_faltas, extra_mes("Tipos", "DATA DO AGENDAMENTO"))),
        "tblDesligados": tabela("tblDesligados", t_desl, m_excel(arq, "tblDesligados", t_desl)),
        "tblHistorico": tabela("tblHistorico", t_hist, m_excel(arq, "tblHistorico", t_hist)),
        "dEmpresa": tabela("dEmpresa", [("EMPRESA", "text")], """let
    Fonte = Table.Distinct(Table.SelectColumns(tblASOS, {"EMPRESA"}))
in
    Fonte"""),
        "dUnidade": tabela("dUnidade", [(UN, "text")], f"""let
    Fonte = Table.Distinct(Table.SelectColumns(tblASOS, {{"{UN}"}}))
in
    Fonte"""),
        "dFaixa": tabela("dFaixa", [("FAIXA", "text", False, None, "ORDEM"), ("ORDEM", "int", True)],
                         ordem_tab("FAIXA", ["Vencido", "Até 7 dias", "8 a 30 dias", "31 a 60 dias", "61 a 90 dias",
                                             "Acima de 90 dias"])),
        "dFaixaAntecipacao": tabela("dFaixaAntecipacao", [("FAIXA", "text", False, None, "ORDEM"), ("ORDEM", "int", True)],
                                    ordem_tab("FAIXA", ["Já vencido", "Até 7 dias", "8 a 30 dias", "31 a 60 dias",
                                                        "61 a 90 dias", "91 a 120 dias"])),
        "dSituacao": tabela("dSituacao", [("SITUAÇÃO", "text", False, None, "ORDEM"), ("ORDEM", "int", True)],
                            ordem_tab("SITUAÇÃO", ["Realizado", "Agendado", "Sem agendamento"])),
        M: tabela_medidas(M, medidas_txt),
    }
    rels = [
        (("tblASOS", "EMPRESA"), ("dEmpresa", "EMPRESA")), (("tblASOS", UN), ("dUnidade", UN)),
        (("tblASOS", "FAIXA 90D"), ("dFaixa", "FAIXA")), (("tblASOS", "SITUAÇÃO AGENDAMENTO"), ("dSituacao", "SITUAÇÃO")),
        (("tblRealizados", "EMPRESA"), ("dEmpresa", "EMPRESA")), (("tblRealizados", UN), ("dUnidade", UN)),
        (("tblRealizados", "FAIXA DE ANTECIPAÇÃO"), ("dFaixaAntecipacao", "FAIXA")),
        (("tblFaltas", "EMPRESA"), ("dEmpresa", "EMPRESA")), (("tblFaltas", UN), ("dUnidade", UN)),
        (("tblDesligados", "EMPRESA"), ("dEmpresa", "EMPRESA")), (("tblDesligados", UN), ("dUnidade", UN)),
        (("tblHistorico", UN), ("dUnidade", UN)),
    ]
    ordem = ["PastaDados", "DataReferencia", "tblASOS", "tblRealizados", "tblFaltas", "tblDesligados", "tblHistorico",
             "dEmpresa", "dUnidade", "dFaixa", "dFaixaAntecipacao", "dSituacao", M]
    exprs = expressoes([("PastaDados", f'"{PASTA_PADRAO}"', "Text"), ("DataReferencia", "#date(2026, 10, 26)", "Date")])
    modelo(nome, tabs, ordem, exprs, relacionamentos(rels))

    # ------------------------------------------------------------- relatório
    slicers = [("filtroEmpresa", v_slicer("dEmpresa", "EMPRESA", "Empresa")),
               ("filtroUnidade", v_slicer("dUnidade", UN, "Unidade de negócio"))]
    sit = col("dSituacao", "SITUAÇÃO")
    cores_sit = {"Realizado": COR["petroleo"], "Agendado": COR["petroleo_claro"], "Sem agendamento": COR["cinza"]}

    p1 = Pagina("Vencimentos", "Vencimentos")
    cabecalho(p1, "Gestão de ASOs", "Vencimentos nos próximos 90 dias · dados fictícios · referência 26/10/2026", slicers)
    linha_cards(p1, [
        ("cardVencimentos", v_card(M, "Vencimentos 90d", "Vencem em até 90 dias")),
        ("cardVencidos", v_card(M, "ASOs Vencidos", "Vencidos", COR["vermelho"])),
        ("cardCobertura", v_card(M, "% Cobertura", "Cobertura (agendado + realizado)")),
        ("cardAgendados", v_card(M, "ASOs Agendados", "Agendados")),
        ("cardSemAgendamento", v_card(M, "Sem Agendamento", "Sem agendamento", COR["ambar"])),
    ])
    p1.add("barFaixas", 24, 200, 740, 252, v_barras(
        "barChart", col("dFaixa", "FAIXA"), [med(M, "Vencimentos 90d")], "Situação por faixa de vencimento",
        serie=sit, cores_serie=cores_sit, ordenar_por=col("dFaixa", "FAIXA"), desc=False))
    p1.add("barUnidades", 776, 200, 480, 496, v_barras(
        "barChart", col("dUnidade", UN), [med(M, "Vencimentos 90d")], "Vencimentos por unidade (clique para filtrar)",
        serie=sit, cores_serie=cores_sit, ordenar_por=med(M, "Vencimentos 90d"), legenda=False))
    p1.add("linhaCobertura", 24, 464, 364, 232, v_barras(
        "lineChart", col("tblHistorico", "DATA DO SNAPSHOT"), [med(M, "% Cobertura (semanal)")],
        "Cobertura nas últimas 12 semanas", ordenar_por=col("tblHistorico", "DATA DO SNAPSHOT"), desc=False))
    p1.add("colVencidosSemana", 400, 464, 364, 232, v_barras(
        "columnChart", col("tblHistorico", "DATA DO SNAPSHOT"), [med(M, "Vencidos (semanal)")],
        "Vencidos por semana", cor_unica=COR["vermelho"], ordenar_por=col("tblHistorico", "DATA DO SNAPSHOT"), desc=False))

    p2 = Pagina("Antecipacao", "Antecipação")
    cabecalho(p2, "Antecipação de ASOs", "Exames realizados de julho a outubro de 2026 · dados fictícios", slicers)
    linha_cards(p2, [
        ("cardRealizados", v_card(M, "ASOs Realizados", "ASOs realizados")),
        ("cardAntecipacao", v_card(M, "Antecipação Média (dias)", "Antecipação média (dias)")),
        ("cardJaVencido", v_card(M, "% Realizado Já Vencido", "Feitos depois de vencer", COR["vermelho"])),
        ("cardAntecipacaoAgendados", v_card(M, "Antecipação Agendados (dias)", "Antecipação dos agendados (dias)")),
    ])
    p2.add("barFaixaAntecipacao", 24, 200, 616, 496, v_barras(
        "barChart", col("dFaixaAntecipacao", "FAIXA"), [med(M, "ASOs Realizados")], "Realizados por faixa de antecipação",
        cor_unica=COR["petroleo"], ordenar_por=col("dFaixaAntecipacao", "FAIXA"), desc=False))
    p2.add("colRealizadosMes", 652, 200, 604, 240, v_barras(
        "columnChart", col("tblRealizados", "MÊS"), [med(M, "ASOs Realizados")], "ASOs realizados por mês",
        cor_unica=COR["petroleo"], ordenar_por=col("tblRealizados", "MÊS"), desc=False))
    p2.add("linhaAntecipacaoMes", 652, 456, 604, 240, v_barras(
        "lineChart", col("tblRealizados", "MÊS"), [med(M, "Antecipação Média (dias)")],
        "Antecipação média por mês (dias)", ordenar_por=col("tblRealizados", "MÊS"), desc=False))

    p3 = Pagina("Faltas", "Faltas")
    cabecalho(p3, "Faltas em exames agendados", "Agosto a outubro de 2026 · nomes fictícios", slicers)
    linha_cards(p3, [
        ("cardFaltas", v_card(M, "Total de Faltas", "Faltas", COR["vermelho"])),
        ("cardPessoas", v_card(M, "Pessoas que Faltaram", "Pessoas que faltaram")),
        ("cardReincidentes", v_card(M, "Reincidentes 2+", "Reincidentes (2+ faltas)")),
        ("cardPctReincidentes", v_card(M, "% Reincidentes", "Taxa de reincidência")),
        ("cardDemissional", v_card(M, "Faltas em Demissional", "Faltas em demissional", COR["ambar"])),
    ])
    p3.add("barFaltasUnidade", 24, 200, 400, 496, v_barras(
        "barChart", col("dUnidade", UN), [med(M, "Total de Faltas")], "Faltas por unidade (clique para filtrar)",
        cor_unica=COR["petroleo"], ordenar_por=med(M, "Total de Faltas")))
    p3.add("barFaltasTipo", 436, 200, 400, 240, v_barras(
        "barChart", col("tblFaltas", "TIPO DE EXAME"), [med(M, "Total de Faltas")], "Faltas por tipo de exame",
        cor_unica=COR["petroleo"], ordenar_por=med(M, "Total de Faltas")))
    p3.add("colFaltasMes", 436, 456, 400, 240, v_barras(
        "columnChart", col("tblFaltas", "MÊS"), [med(M, "Total de Faltas")], "Faltas por mês",
        cor_unica=COR["vermelho"], ordenar_por=col("tblFaltas", "MÊS"), desc=False))
    p3.add("tabelaFaltas", 848, 200, 408, 496, v_tabela(
        [(col("tblFaltas", "NOME"), "Colaborador"), (col("tblFaltas", UN), "Unidade"),
         (med(M, "Total de Faltas"), "Faltas")], "Colaboradores com falta", ordenar_por=med(M, "Total de Faltas")))

    relatorio(nome, [p1, p2, p3], "TemaSSO")


# =============================================================================
# PROJETO 2 — CUSTOS DE EXAMES
# =============================================================================
def projeto_custos():
    nome = "Custos_Exames_DEMO"
    arq = "Base_Custos_Exames_DEMO.xlsx"
    M = "Medidas"
    alfa, beta = "EMPRESA ALFA", "EMPRESA BETA"

    t_exame = [("ID_EXAME", "text"), ("EXAME", "text"), ("GRUPO", "text"), ("REDE", "text")]
    t_novo = [("ID_EXAME", "text"), ("NOVO PREÇO", "num"), ("SITUAÇÃO DA NEGOCIAÇÃO", "text")]
    t_vol = [("EMPRESA", "text"), ("ID_EXAME", "text"), ("VOLUME MENSAL", "int")]
    t_preco = [("EMPRESA", "text"), ("ID_EXAME", "text"), ("PREÇO ATUAL", "num")]
    t_emp = [("EMPRESA", "text"), ("VIDAS", "int"), ("PER CAPITA ATUAL", "num"), ("PER CAPITA NOVO", "num")]
    t_per = [("PERÍODO", "text"), ("FATOR", "int"), ("ORDEM", "int")]
    t_cen = [("CENÁRIO", "text"), ("ORDEM", "int")]
    t_prem = [("ITEM", "text"), ("REGRA", "text")]

    m_dexame = f"""let
    Fonte = Excel.Workbook(File.Contents(PastaDados & "{arq}"), null, true),
    Exames = Table.TransformColumnTypes(Fonte{{[Item="dExame",Kind="Table"]}}[Data], {{{{"ID_EXAME", type text}}, {{"EXAME", type text}}, {{"GRUPO", type text}}, {{"REDE", type text}}}}),
    Novo = Table.TransformColumnTypes(Fonte{{[Item="fNovoPreco",Kind="Table"]}}[Data], {{{{"ID_EXAME", type text}}, {{"NOVO PREÇO", type number}}, {{"SITUAÇÃO DA NEGOCIAÇÃO", type text}}}}),
    Volume = Table.TransformColumnTypes(Fonte{{[Item="fVolume",Kind="Table"]}}[Data], {{{{"EMPRESA", type text}}, {{"ID_EXAME", type text}}}}),
    ComNovo = Table.NestedJoin(Exames, {{"ID_EXAME"}}, Novo, {{"ID_EXAME"}}, "N", JoinKind.LeftOuter),
    Expandido = Table.ExpandTableColumn(ComNovo, "N", {{"NOVO PREÇO", "SITUAÇÃO DA NEGOCIAÇÃO"}}),
    Empresas = Table.AddColumn(Expandido, "EMPRESAS QUE REALIZAM", (linha) =>
        let
            Lista = Table.SelectRows(Volume, each [ID_EXAME] = linha[ID_EXAME])[EMPRESA],
            TemAlfa = List.Contains(Lista, "{alfa}"),
            TemBeta = List.Contains(Lista, "{beta}")
        in
            if TemAlfa and TemBeta then "Ambas" else if TemAlfa then "Só Alfa" else "Só Beta", type text)
in
    Empresas"""

    m_fato = f"""let
    Fonte = Excel.Workbook(File.Contents(PastaDados & "{arq}"), null, true),
    Volume = Table.TransformColumnTypes(Fonte{{[Item="fVolume",Kind="Table"]}}[Data], {{{{"EMPRESA", type text}}, {{"ID_EXAME", type text}}, {{"VOLUME MENSAL", Int64.Type}}}}),
    Preco = Table.TransformColumnTypes(Fonte{{[Item="fPreco",Kind="Table"]}}[Data], {{{{"EMPRESA", type text}}, {{"ID_EXAME", type text}}, {{"PREÇO ATUAL", type number}}}}),
    ComPreco = Table.NestedJoin(Volume, {{"EMPRESA", "ID_EXAME"}}, Preco, {{"EMPRESA", "ID_EXAME"}}, "P", JoinKind.LeftOuter),
    Fato = Table.ExpandTableColumn(ComPreco, "P", {{"PREÇO ATUAL"}}),
    PrecoDe = (empresa as text, id as text) as nullable number =>
        let Linhas = Table.SelectRows(Preco, each [EMPRESA] = empresa and [ID_EXAME] = id)
        in if Table.IsEmpty(Linhas) then null else Linhas{{0}}[PREÇO ATUAL],
    ComAlfa = Table.AddColumn(Fato, "PREÇO ALFA", each PrecoDe("{alfa}", [ID_EXAME]), type number),
    ComBeta = Table.AddColumn(ComAlfa, "PREÇO BETA", each PrecoDe("{beta}", [ID_EXAME]), type number),
    Menor = Table.AddColumn(ComBeta, "MENOR PREÇO", each List.Min(List.RemoveNulls({{[PREÇO ALFA], [PREÇO BETA]}})), type number)
in
    Menor"""

    def m_estatico(nome_tab, valores, colunas_tipos):
        tipo_m = {"text": "text", "int": "Int64.Type", "num": "number", "date": "date", "bool": "logical"}
        tipo = ", ".join(f'#"{c}" = {tipo_m[t]}' for c, t in colunas_tipos)
        linhas = ", ".join("{" + ", ".join(json.dumps(v, ensure_ascii=False) for v in row) + "}" for row in valores)
        return f"""let
    Fonte = #table(type table [{tipo}], {{{linhas}}})
in
    Fonte"""

    rs = "\\R\\$ #,0"
    rs2 = "\\R\\$ #,0.00"
    F = "1. Base"
    medidas_txt = "\n".join([
        medida("Fator Período", "SELECTEDVALUE ( dPeriodo[FATOR], 1 )", "0", F, "1 para mensal, 12 para anual."),
        medida("Volume do Mês", "SUM ( fExameEmpresa[VOLUME MENSAL] )", "#,0", F),
        medida("Custo Exames Atual", """SUMX (
    fExameEmpresa,
    fExameEmpresa[VOLUME MENSAL] * fExameEmpresa[PREÇO ATUAL]
) * [Fator Período]""", rs, "2. Atual x Novo", "Volume × preço da tabela atual de cada empresa."),
        medida("Custo Per Capita Atual", "SUMX ( dEmpresa, dEmpresa[VIDAS] * dEmpresa[PER CAPITA ATUAL] ) * [Fator Período]", rs,
               "2. Atual x Novo"),
        medida("Custo Atual", "[Custo Exames Atual] + [Custo Per Capita Atual]", rs, "2. Atual x Novo"),
        medida("Custo Exames Novo", """SUMX (
    fExameEmpresa,
    fExameEmpresa[VOLUME MENSAL] * RELATED ( dExame[NOVO PREÇO] )
) * [Fator Período]""", rs, "2. Atual x Novo", "Volume × preço único da tabela negociada."),
        medida("Custo Per Capita Novo", "SUMX ( dEmpresa, dEmpresa[VIDAS] * dEmpresa[PER CAPITA NOVO] ) * [Fator Período]", rs,
               "2. Atual x Novo"),
        medida("Custo Novo", "[Custo Exames Novo] + [Custo Per Capita Novo]", rs, "2. Atual x Novo"),
        medida("Economia", "[Custo Atual] - [Custo Novo]", rs, "2. Atual x Novo"),
        medida("Economia %", "DIVIDE ( [Economia], [Custo Atual] )", "0.0%", "2. Atual x Novo"),
        medida("Economia Exames", "[Custo Exames Atual] - [Custo Exames Novo]", rs, "2. Atual x Novo",
               "Só a parte de exames, sem per capita: usar em visuais por exame."),
        medida("Preço Tabela Alfa", f'CALCULATE ( MAX ( fExameEmpresa[PREÇO ATUAL] ), dEmpresa[EMPRESA] = "{alfa}" )', rs2,
               "3. Preços"),
        medida("Preço Tabela Beta", f'CALCULATE ( MAX ( fExameEmpresa[PREÇO ATUAL] ), dEmpresa[EMPRESA] = "{beta}" )', rs2,
               "3. Preços"),
        medida("Preço Negociado", "MAX ( dExame[NOVO PREÇO] )", rs2, "3. Preços"),
        medida("Custo Cenário", f"""VAR Cenario = SELECTEDVALUE ( dCenario[ORDEM] )
VAR Vidas = SUM ( dEmpresa[VIDAS] )
VAR PcAlfa = CALCULATE ( MAX ( dEmpresa[PER CAPITA ATUAL] ), dEmpresa[EMPRESA] = "{alfa}" )
VAR PcBeta = CALCULATE ( MAX ( dEmpresa[PER CAPITA ATUAL] ), dEmpresa[EMPRESA] = "{beta}" )
VAR ExAtual = SUMX ( fExameEmpresa, fExameEmpresa[VOLUME MENSAL] * fExameEmpresa[PREÇO ATUAL] )
VAR ExAlfa = SUMX ( fExameEmpresa, fExameEmpresa[VOLUME MENSAL] * COALESCE ( fExameEmpresa[PREÇO ALFA], fExameEmpresa[PREÇO ATUAL] ) )
VAR ExBeta = SUMX ( fExameEmpresa, fExameEmpresa[VOLUME MENSAL] * COALESCE ( fExameEmpresa[PREÇO BETA], fExameEmpresa[PREÇO ATUAL] ) )
VAR ExMenor = SUMX ( fExameEmpresa, fExameEmpresa[VOLUME MENSAL] * fExameEmpresa[MENOR PREÇO] )
VAR ExNovo = SUMX ( fExameEmpresa, fExameEmpresa[VOLUME MENSAL] * RELATED ( dExame[NOVO PREÇO] ) )
VAR PcAtual = SUMX ( dEmpresa, dEmpresa[VIDAS] * dEmpresa[PER CAPITA ATUAL] )
VAR PcNovo = SUMX ( dEmpresa, dEmpresa[VIDAS] * dEmpresa[PER CAPITA NOVO] )
VAR Mensal =
    SWITCH (
        Cenario,
        1, ExAtual + PcAtual,
        2, ExAlfa + Vidas * PcAlfa,
        3, ExBeta + Vidas * PcBeta,
        4, ExMenor + Vidas * MIN ( PcAlfa, PcBeta ),
        5, ExNovo + PcNovo
    )
RETURN
    Mensal * [Fator Período]""", rs, "4. Cenários", "Custo total (exames + per capita) em cada cenário de tabela."),
        medida("Valor Ponte", """VAR Etapa = SELECTEDVALUE ( dPonte[ORDEM] )
VAR Reducao = ( fExameEmpresa[VOLUME MENSAL] * ( fExameEmpresa[PREÇO ATUAL] - RELATED ( dExame[NOVO PREÇO] ) ) )
VAR Comum = SUMX ( FILTER ( fExameEmpresa, RELATED ( dExame[EMPRESAS QUE REALIZAM] ) = "Ambas" ), fExameEmpresa[VOLUME MENSAL] * ( fExameEmpresa[PREÇO ATUAL] - RELATED ( dExame[NOVO PREÇO] ) ) )
VAR SoAlfa = SUMX ( FILTER ( fExameEmpresa, RELATED ( dExame[EMPRESAS QUE REALIZAM] ) = "Só Alfa" ), fExameEmpresa[VOLUME MENSAL] * ( fExameEmpresa[PREÇO ATUAL] - RELATED ( dExame[NOVO PREÇO] ) ) )
VAR SoBeta = SUMX ( FILTER ( fExameEmpresa, RELATED ( dExame[EMPRESAS QUE REALIZAM] ) = "Só Beta" ), fExameEmpresa[VOLUME MENSAL] * ( fExameEmpresa[PREÇO ATUAL] - RELATED ( dExame[NOVO PREÇO] ) ) )
VAR PerCapita = SUMX ( dEmpresa, dEmpresa[VIDAS] * ( dEmpresa[PER CAPITA ATUAL] - dEmpresa[PER CAPITA NOVO] ) )
RETURN
    SWITCH (
        Etapa,
        1, [Custo Atual],
        2, -Comum * [Fator Período],
        3, -SoAlfa * [Fator Período],
        4, -SoBeta * [Fator Período],
        5, -PerCapita * [Fator Período]
    )""", rs, "4. Cenários", "Etapas do custo atual ao negociado; o total da cascata é o custo negociado."),
    ])
    # a variável Reducao não é usada; removida para não confundir quem ler o DAX
    medidas_txt = medidas_txt.replace(
        "\t\t\tVAR Reducao = ( fExameEmpresa[VOLUME MENSAL] * ( fExameEmpresa[PREÇO ATUAL] - RELATED ( dExame[NOVO PREÇO] ) ) )\n", "")

    tabs = {
        "dExame": tabela("dExame", t_exame + t_novo[1:] + [("EMPRESAS QUE REALIZAM", "text")], m_dexame,
                         descricao="Catálogo de exames com o novo preço negociado."),
        "dEmpresa": tabela("dEmpresa", t_emp, m_excel(arq, "dEmpresa", t_emp)),
        "fExameEmpresa": tabela("fExameEmpresa", t_vol + [("PREÇO ATUAL", "num"), ("PREÇO ALFA", "num", True),
                                                          ("PREÇO BETA", "num", True), ("MENOR PREÇO", "num", True)],
                                m_fato, descricao="Fato: volume mensal e preço atual por empresa e exame."),
        "dPeriodo": tabela("dPeriodo", [("PERÍODO", "text", False, None, "ORDEM"), ("FATOR", "int"), ("ORDEM", "int", True)],
                           m_excel(arq, "dPeriodo", t_per)),
        "dCenario": tabela("dCenario", [("CENÁRIO", "text", False, None, "ORDEM"), ("ORDEM", "int", True)],
                           m_excel(arq, "dCenario", t_cen)),
        "dPonte": tabela("dPonte", [("ETAPA", "text", False, None, "ORDEM"), ("ORDEM", "int", True)],
                         m_estatico("dPonte", [["Custo atual", 1], ["Exames em comum", 2], ["Exclusivos da Alfa", 3],
                                               ["Exclusivos da Beta", 4], ["Per capita por vida", 5]],
                                    [("ETAPA", "text"), ("ORDEM", "int")])),
        "tblPremissas": tabela("tblPremissas", t_prem, m_excel(arq, "tblPremissas", t_prem)),
        M: tabela_medidas(M, medidas_txt),
    }
    rels = [(("fExameEmpresa", "ID_EXAME"), ("dExame", "ID_EXAME")),
            (("fExameEmpresa", "EMPRESA"), ("dEmpresa", "EMPRESA"))]
    ordem = ["PastaDados", "dExame", "dEmpresa", "fExameEmpresa", "dPeriodo", "dCenario", "dPonte", "tblPremissas", M]
    modelo(nome, tabs, ordem, expressoes([("PastaDados", f'"{PASTA_PADRAO}"', "Text")]), relacionamentos(rels))

    # ------------------------------------------------------------- relatório
    f_periodo = ("filtroPeriodo", v_slicer("dPeriodo", "PERÍODO", "Período", unico=True))
    p1 = Pagina("Resumo", "Resumo")
    cabecalho(p1, "Simulação de custos de exames ocupacionais",
              "Duas unidades de negócio, cinco cenários de tabela · dados fictícios", [f_periodo])
    linha_cards(p1, [
        ("cardCustoAtual", v_card(M, "Custo Atual", "Custo atual (contratos separados)", COR["tinta"])),
        ("cardCustoNovo", v_card(M, "Custo Novo", "Custo com novo preço")),
        ("cardEconomia", v_card(M, "Economia", "Economia", COR["verde"])),
        ("cardEconomiaPct", v_card(M, "Economia %", "Redução sobre o custo atual", COR["verde"])),
    ])
    p1.add("cascataPonte", 24, 200, 616, 496, v_cascata(col("dPonte", "ETAPA"), med(M, "Valor Ponte"),
                                                          "Do custo atual ao negociado (Total = custo negociado)"))
    p1.add("barCenarios", 652, 200, 604, 496, v_barras(
        "barChart", col("dCenario", "CENÁRIO"), [med(M, "Custo Cenário")], "Custo total por cenário de tabela",
        cor_unica=COR["petroleo"], ordenar_por=col("dCenario", "CENÁRIO"), desc=False))

    p2 = Pagina("PorExame", "Por exame")
    cabecalho(p2, "Custo por exame", "Valores conforme o período escolhido · dados fictícios",
              [("filtroEmpresa", v_slicer("dEmpresa", "EMPRESA", "Empresa")), f_periodo])
    p2.add("tabelaExames", 24, 92, 820, 604, v_tabela(
        [(col("dExame", "EXAME"), "Exame"), (col("dExame", "REDE"), "Rede"),
         (col("dExame", "EMPRESAS QUE REALIZAM"), "Empresas"), (med(M, "Volume do Mês"), "Volume/mês"),
         (med(M, "Preço Tabela Alfa"), "Preço Alfa"), (med(M, "Preço Tabela Beta"), "Preço Beta"), (med(M, "Preço Negociado"), "Novo preço"),
         (med(M, "Custo Exames Atual"), "Custo atual"), (med(M, "Custo Exames Novo"), "Custo novo"),
         (med(M, "Economia Exames"), "Economia"), (col("dExame", "SITUAÇÃO DA NEGOCIAÇÃO"), "Negociação")],
        "Tabela de exames", ordenar_por=med(M, "Economia Exames")))
    p2.add("barEconomiaExame", 856, 92, 400, 604, v_barras(
        "barChart", col("dExame", "EXAME"), [med(M, "Economia Exames")], "Economia por exame",
        cor_unica=COR["verde"], ordenar_por=med(M, "Economia Exames")))

    p3 = Pagina("Premissas", "Premissas")
    cabecalho(p3, "Premissas e regras de cálculo", "Como os números deste relatório são calculados", [])
    p3.add("tabelaPremissas", 24, 92, 1232, 380, v_tabela(
        [(col("tblPremissas", "ITEM"), "Item"), (col("tblPremissas", "REGRA"), "Regra")], "Premissas"))
    p3.add("notaModelo", 24, 488, 1232, 120, v_texto([
        [("Modelo: ", 11, COR["tinta"], True),
         ("estrela, com a fato fExameEmpresa (volume × preço por empresa) ligada a dExame e dEmpresa. "
          "dPeriodo, dCenario e dPonte são tabelas de parâmetro desconectadas, lidas por SELECTEDVALUE nas medidas.",
          11, COR["suave"], False)],
        [("Origem: ", 11, COR["tinta"], True),
         ("planilha Excel fictícia carregada por Power Query; o caminho fica no parâmetro PastaDados.", 11, COR["suave"], False)],
    ]))

    relatorio(nome, [p1, p2, p3], "TemaSSO")


def gitignore():
    w(OUT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n", False)


if __name__ == "__main__":
    if OUT.exists():
        shutil.rmtree(OUT)
    projeto_asos()
    projeto_custos()
    gitignore()
    print("ok", sum(1 for _ in OUT.rglob("*") if _.is_file()), "arquivos")
