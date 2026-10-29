"""Versão 2 dos dois projetos Power BI do portfólio (PBIP: modelo TMDL + relatório PBIR).

Mais páginas, navegação lateral, cartões modernos, mapa de calor, Pareto, cascata,
comparativos por empresa e filtros sincronizados entre páginas. Dados 100% fictícios.

Reaproveita os auxiliares de TMDL/M de gerar_pbip.py.
"""
import json
import shutil
import uuid

from gerar_pbip import (BASE_THEME, OUT, PASTA_PADRAO, SCH, TIPOS, V_PAGE, V_PAGES, V_PBIP, V_PBIR, V_PLATFORM,
                        V_REPORT, V_VERSION, V_VISUAL, L, b, col, expressoes, gid, m_excel, med, medida, modelo, n,
                        props, q, relacionamentos, s, tabela, tabela_medidas, w)

# ----------------------------------------------------------------------------- identidade visual
C = {
    "lateral": "#102A36", "lateral_txt": "#C9D9E0", "tinta": "#14303C", "suave": "#5E6E78", "fundo": "#EEF2F4",
    "borda": "#DDE4E8", "branco": "#FFFFFF",
    "alfa": "#2563A8", "beta": "#E08E45", "novo": "#2E9E5B", "cinza": "#B8C4CB",
    "claro": "#93BDE6", "vermelho": "#C0504D", "ambar": "#E0A13A", "verde": "#2E8B6E",
}
X0, LARG = 216, 1040          # área de conteúdo (à direita da barra lateral)
GAP = 12
M = "Medidas"


def cor(h):
    return {"solid": {"color": L(f"'{h}'")}}


def proj(campo, nome=None):
    k = "Column" if "Column" in campo else "Measure"
    t, p = campo[k]["Expression"]["SourceRef"]["Entity"], campo[k]["Property"]
    out = {"field": campo, "queryRef": f"{t}.{p}", "nativeQueryRef": nome or p}
    if nome:
        out["displayName"] = nome
    return out


def qref(campo):
    k = "Column" if "Column" in campo else "Measure"
    return f'{campo[k]["Expression"]["SourceRef"]["Entity"]}.{campo[k]["Property"]}'


# ----------------------------------------------------------------------------- contêineres e peças
def vco(titulo=None, fundo=True, sub=None):
    o = {
        "background": props(show=b(fundo), color=cor(C["branco"]), transparency=n(0)),
        "border": props(show=b(fundo), color=cor(C["borda"]), radius=n(12), width=n(1)),
        "padding": props(top=n(10), bottom=n(8), left=n(14), right=n(14)),
        "dropShadow": props(show=b(False)),
        "title": props(show=b(bool(titulo)), **({"text": s(titulo), "fontColor": cor(C["tinta"]), "fontSize": n(12),
                                               "fontFamily": s("Segoe UI Semibold")} if titulo else {})),
    }
    if sub:
        o["subTitle"] = props(show=b(True), text=s(sub), fontColor=cor(C["suave"]), fontSize=n(9))
    return o


def filtro(tab, coluna, valores, negar=False, nome=None):
    cond = {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": coluna}}],
                   "Values": [[{"Literal": {"Value": f"'{v}'"}}] for v in valores]}}
    if negar:
        cond = {"Not": {"Expression": cond}}
    return {"name": nome or ("Filtro" + uuid.uuid4().hex[:20]), "field": col(tab, coluna), "type": "Categorical",
            "filter": {"Version": 2, "From": [{"Name": "t", "Entity": tab, "Type": 0}], "Where": [{"Condition": cond}]},
            "howCreated": "User"}


class Pagina:
    def __init__(self, nome, exibicao, filtros=None):
        self.nome, self.exibicao, self.visuais, self.filtros = nome, exibicao, [], filtros or []

    def add(self, nome, x, y, largura, altura, visual, filtros=None):
        z = len(self.visuais) * 1000
        v = {"$schema": V_VISUAL, "name": nome,
             "position": {"x": x, "y": y, "z": z, "height": altura, "width": largura, "tabOrder": z}, "visual": visual}
        if filtros:
            v["filterConfig"] = {"filters": filtros}
        self.visuais.append(v)


def v_shape(fill):
    return {"visualType": "shape",
            "objects": {"shape": [{"properties": {"tileShape": s("rectangle")}}],
                        "fill": [{"properties": {"fillColor": cor(fill), "transparency": n(0)}, "selector": {"id": "default"}}],
                        "outline": [{"properties": {"show": b(False)}, "selector": {"id": "default"}}]},
            "visualContainerObjects": {"background": props(show=b(False)), "border": props(show=b(False)),
                                       "padding": props(top=n(0), bottom=n(0), left=n(0), right=n(0)),
                                       "title": props(show=b(False))},
            "drillFilterOtherVisuals": True}


def v_texto(paragrafos):
    return {"visualType": "textbox",
            "objects": {"general": [{"properties": {"paragraphs": [
                {"textRuns": [{"value": t, "textStyle": {"fontSize": f"{tam}pt", "color": c, "fontFamily": "Segoe UI",
                                                         **({"fontWeight": "bold"} if neg else {})}}
                              for t, tam, c, neg in par]} for par in paragrafos]}}]},
            "visualContainerObjects": {"background": props(show=b(False)), "border": props(show=b(False)),
                                       "padding": props(top=n(0), bottom=n(0), left=n(0), right=n(0)),
                                       "title": props(show=b(False))},
            "drillFilterOtherVisuals": True}


def v_nav(qtd):
    sel = lambda est, **kw: {"properties": kw, "selector": {"id": est}}
    return {"visualType": "pageNavigator",
            "objects": {
                "layout": [{"properties": {"orientation": L("1D"), "cellPadding": L("4L")}}],
                "shape": [{"properties": {"tileShape": s("rectangleRoundedByPixel"), "roundEdge": L("8L")}},
                          {"properties": {"tileShape": s("rectangleRoundedByPixel")}, "selector": {"id": "default"}}],
                "fill": [{"properties": {"show": b(True)}},
                         sel("default", fillColor=cor(C["lateral"]), transparency=n(0)),
                         sel("hover", fillColor=cor("#1C4252"), transparency=n(0)),
                         sel("selected", fillColor=cor(C["alfa"]), transparency=n(0))],
                "text": [{"properties": {"show": b(True)}},
                         sel("default", fontColor=cor(C["lateral_txt"]), fontSize=n(10), fontFamily=s("Segoe UI"),
                             horizontalAlignment=s("left"), leftMargin=L("12L")),
                         sel("hover", fontColor=cor(C["branco"])),
                         sel("selected", fontColor=cor(C["branco"]), bold=b(True))],
                "outline": [{"properties": {"show": b(False)}}],
            },
            "visualContainerObjects": {"background": props(show=b(False)), "border": props(show=b(False)),
                                       "padding": props(top=n(0), bottom=n(0), left=n(0), right=n(0)),
                                       "title": props(show=b(False))},
            "drillFilterOtherVisuals": True}


def v_card(m, rotulo, acento, cor_valor=None):
    d = {"id": "default"}
    return {"visualType": "cardVisual",
            "query": {"queryState": {"Data": {"projections": [proj(med(M, m), rotulo)]}}},
            "objects": {
                "value": [{"properties": {"fontSize": n(22), "fontColor": cor(cor_valor or C["tinta"]),
                                          "fontFamily": s("Segoe UI Semibold"), "labelDisplayUnits": L("1D")},
                           "selector": d}],
                "label": [{"properties": {"show": b(True), "fontSize": n(10), "fontColor": cor(C["suave"]),
                                          "position": s("belowValue")}, "selector": d}],
                "accentBar": [{"properties": {"show": b(True), "position": s("Left"), "color": cor(acento),
                                              "width": n(5)}, "selector": d}],
                "padding": [{"properties": {"paddingUniform": n(10)}, "selector": d}],
                "layout": [{"properties": {"paddingUniform": n(0)}, "selector": d}],
            },
            "visualContainerObjects": {**vco(), "padding": props(top=n(4), bottom=n(4), left=n(4), right=n(8))},
            "drillFilterOtherVisuals": True}


def v_slicer(tab, c, rotulo, modo="HorizontalList", grupo=None, unico=False):
    horizontal = modo == "HorizontalList"
    objs = {"data": props(mode=s("Basic" if horizontal else modo)),
            "header": props(show=b(True), text=s(rotulo), fontColor=cor(C["suave"]), textSize=n(9)),
            "items": props(fontColor=cor(C["tinta"]), textSize=n(10))}
    if horizontal:
        objs["general"] = props(orientation=L("1D"))
        objs["items"] = props(fontColor=cor(C["tinta"]), textSize=n(10), background=cor("#F4F7F8"))
    if unico:
        objs["selection"] = props(strictSingleSelect=b(True))
    v = {"visualType": "slicer",
         "query": {"queryState": {"Values": {"projections": [{**proj(col(tab, c), rotulo), "active": True}]}},
                   "sortDefinition": {"sort": [{"field": col(tab, c), "direction": "Ascending"}], "isDefaultSort": True}},
         "objects": objs,
         "visualContainerObjects": {**vco(), "padding": props(top=n(4), bottom=n(4), left=n(10), right=n(10))},
         "drillFilterOtherVisuals": True}
    if grupo:
        v["syncGroup"] = {"groupName": grupo, "fieldChanges": True, "filterChanges": True}
    return v


def seletor_valor(campo_col, valor):
    return {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": campo_col,
                                                  "Right": {"Literal": {"Value": f"'{valor}'"}}}}}]}


def gradiente(medida_nome, minimo, maximo, meio=None):
    regra = {"min": {"color": L(f"'{minimo}'")["expr"]}, "max": {"color": L(f"'{maximo}'")["expr"]},
             "nullColoringStrategy": {"strategy": {"Literal": {"Value": "'asZero'"}}}}
    chave = "linearGradient2"
    if meio:
        regra["mid"] = {"color": L(f"'{meio}'")["expr"]}
        chave = "linearGradient3"
    return {"solid": {"color": {"expr": {"FillRule": {"Input": med(M, medida_nome), "FillRule": {chave: regra}}}}}}


def v_grafico(tipo, categoria, ys, titulo, serie=None, cores_serie=None, cores_y=None, cor_unica=None, y2=None,
              ordenar=None, desc=True, rotulos=True, legenda=None, sub=None, grad=None, totais=False, fmt_rotulo=None):
    qs = {"Category": {"projections": [{**proj(categoria), "active": True}]},
          "Y": {"projections": [proj(*y) if isinstance(y, tuple) else proj(y) for y in ys]}}
    if y2:
        qs["Y2"] = {"projections": [proj(*y) if isinstance(y, tuple) else proj(y) for y in y2]}
    if serie:
        qs["Series"] = {"projections": [proj(serie)]}
    query = {"queryState": qs}
    if ordenar is not None:
        query["sortDefinition"] = {"sort": [{"field": ordenar, "direction": "Descending" if desc else "Ascending"}]}
    mostra_leg = legenda if legenda is not None else (bool(serie) or len(ys) + len(y2 or []) > 1)
    objs = {
        "labels": props(show=b(rotulos), color=cor(C["tinta"]), fontSize=n(9),
                        **({"labelDisplayUnits": L(f"{fmt_rotulo}D")} if fmt_rotulo else {})),
        "valueAxis": props(show=b(False), gridlineShow=b(False)),
        "categoryAxis": props(showAxisTitle=b(False), labelColor=cor(C["tinta"]), fontSize=n(9)),
        "legend": props(show=b(mostra_leg), position=s("TopRight"), showTitle=b(False), labelColor=cor(C["suave"]),
                        fontSize=n(9)),
    }
    dp = []
    if cor_unica:
        dp.append({"properties": {"fill": cor(cor_unica)}})
    if grad:
        dp.append({"properties": {"fill": grad}, "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}})
    if serie and cores_serie:
        for valor, h in cores_serie.items():
            dp.append({"properties": {"fill": cor(h)}, "selector": seletor_valor(serie, valor)})
    if cores_y:
        for campo, h in cores_y:
            dp.append({"properties": {"fill": cor(h)}, "selector": {"metadata": qref(campo)}})
    if dp:
        objs["dataPoint"] = dp
    if tipo in ("lineChart", "lineClusteredColumnComboChart"):
        objs["lineStyles"] = props(strokeWidth=n(3), showMarker=b(True), markerSize=n(5))
    if totais:
        objs["totals"] = props(show=b(True))
    return {"visualType": tipo, "query": query, "objects": objs, "visualContainerObjects": vco(titulo, sub=sub),
            "drillFilterOtherVisuals": True}


def v_rosca(categoria, y, titulo, cores=None, sub=None):
    objs = {"labels": props(show=b(True), labelStyle=s("Category, percent of total"), color=cor(C["tinta"]),
                            fontSize=n(9), position=s("outside"), background=s("off")),
            "legend": props(show=b(False))}
    if cores:
        objs["dataPoint"] = [{"properties": {"fill": cor(h)}, "selector": seletor_valor(categoria, v)} for v, h in cores.items()]
    return {"visualType": "donutChart",
            "query": {"queryState": {"Category": {"projections": [{**proj(categoria), "active": True}]},
                                     "Y": {"projections": [proj(y)]}}},
            "objects": objs, "visualContainerObjects": vco(titulo, sub=sub), "drillFilterOtherVisuals": True}


def v_tabela(campos, titulo, ordenar=None, desc=True, barras=None, grad=None, sub=None):
    """campos: [(campo, nome)]; barras: [campo medida]; grad: [(campo medida, min, max, mid)]"""
    query = {"queryState": {"Values": {"projections": [proj(c, nm) for c, nm in campos]}}}
    if ordenar is not None:
        query["sortDefinition"] = {"sort": [{"field": ordenar, "direction": "Descending" if desc else "Ascending"}]}
    objs = {
        "columnHeaders": props(fontColor=cor(C["suave"]), backColor=cor(C["branco"]), fontSize=n(9), bold=b(True)),
        "values": props(fontColor=cor(C["tinta"]), fontSize=n(9), backColorPrimary=cor(C["branco"]),
                        backColorSecondary=cor("#F6F9FA")),
        "grid": props(gridHorizontal=b(True), gridHorizontalColor=cor(C["borda"]), gridVertical=b(False)),
        "total": props(totals=b(False)),
    }
    if barras:
        objs["columnFormatting"] = [{"properties": {"dataBars": {
            "positiveColor": cor(c_), "negativeColor": cor(C["vermelho"]), "axisColor": cor(C["cinza"]),
            "reverseDirection": b(False), "hideText": b(False)}}, "selector": {"metadata": qref(campo)}}
            for campo, c_ in barras]
    if grad:
        for campo, mn, mx, md in grad:
            nome_m = campo["Measure"]["Property"]
            objs["values"].append({"properties": {"backColor": gradiente(nome_m, mn, mx, md)},
                                   "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}],
                                                "metadata": qref(campo)}})
    return {"visualType": "tableEx", "query": query, "objects": objs, "visualContainerObjects": vco(titulo, sub=sub),
            "drillFilterOtherVisuals": True}


def v_matriz(linha, coluna, valor, titulo, grad=None, sub=None, total_linhas=True):
    objs = {
        "columnHeaders": props(fontColor=cor(C["suave"]), backColor=cor(C["branco"]), fontSize=n(9), bold=b(True)),
        "rowHeaders": props(fontColor=cor(C["tinta"]), fontSize=n(9)),
        "values": props(fontColor=cor(C["tinta"]), fontSize=n(10)),
        "grid": props(gridHorizontal=b(True), gridHorizontalColor=cor(C["borda"]), gridVertical=b(True),
                      gridVerticalColor=cor(C["borda"])),
    }
    if not total_linhas:
        objs["subTotals"] = props(rowSubtotals=b(False))
    if grad:
        mn, mx = grad
        objs["values"].append({"properties": {"backColor": gradiente(valor["Measure"]["Property"], mn, mx)},
                               "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": qref(valor)}})
    return {"visualType": "pivotTable",
            "query": {"queryState": {"Rows": {"projections": [{**proj(linha), "active": True}]},
                                     "Columns": {"projections": [{**proj(coluna), "active": True}]},
                                     "Values": {"projections": [proj(valor)]}}},
            "objects": objs, "visualContainerObjects": vco(titulo, sub=sub), "drillFilterOtherVisuals": True}


def v_cascata(categoria, valor, titulo, sub=None):
    return {"visualType": "waterfallChart",
            "query": {"queryState": {"Category": {"projections": [{**proj(categoria), "active": True}]},
                                     "Y": {"projections": [proj(valor)]}},
                      "sortDefinition": {"sort": [{"field": categoria, "direction": "Ascending"}]}},
            "objects": {"sentimentColors": props(increaseFill=cor(C["cinza"]), decreaseFill=cor(C["novo"]),
                                                 totalFill=cor(C["alfa"])),
                        "labels": props(show=b(True), color=cor(C["tinta"]), fontSize=n(9), labelDisplayUnits=L("1D"),
                                        labelPrecision=L("0L")),
                        "valueAxis": props(show=b(False), gridlineShow=b(False)),
                        "categoryAxis": props(showAxisTitle=b(False), labelColor=cor(C["tinta"]), fontSize=n(9)),
                        "legend": props(show=b(False))},
            "visualContainerObjects": vco(titulo, sub=sub), "drillFilterOtherVisuals": True}


# ----------------------------------------------------------------------------- moldura de página
def moldura(p, titulo, subtitulo, marca, qtd_paginas, slicers=()):
    p.add("barraLateral", 0, 0, 200, 720, v_shape(C["lateral"]))
    p.add("marca", 16, 18, 172, 64, v_texto([[(marca, 13, C["branco"], True)],
                                              [("dados fictícios · portfólio", 9, C["lateral_txt"], False)]]))
    p.add("navegacao", 8, 96, 184, min(36 * qtd_paginas, 600), v_nav(qtd_paginas))
    p.add("rodape", 16, 676, 172, 32, v_texto([[("Victor Santos · Power BI", 8, C["lateral_txt"], False)]]))
    ocupado = sum(lg + 8 for _, lg, _ in slicers)
    larg_titulo = min(640, LARG - ocupado - 12)
    p.add("titulo", X0, 12, larg_titulo, 36, v_texto([[(titulo, 20 if len(titulo) <= 34 else 17, C["tinta"], True)]]))
    p.add("subtitulo", X0, 48, larg_titulo, 24, v_texto([[(subtitulo, 10 if len(subtitulo) <= 70 else 9, C["suave"], False)]]))
    x = X0 + LARG
    for nome, largura, visual in reversed(list(slicers)):
        x -= largura
        p.add(nome, x, 6, largura, 78, visual)
        x -= 8


def cards(p, itens, y=94, h=88, x=X0, largura_total=LARG):
    k = len(itens)
    wd = (largura_total - GAP * (k - 1)) / k
    for i, (nome, visual) in enumerate(itens):
        p.add(nome, round(x + i * (wd + GAP)), y, round(wd), h, visual)


def tema(arq):
    return {
        "name": arq,
        "dataColors": [C["alfa"], C["beta"], C["novo"], C["claro"], C["vermelho"], C["ambar"], "#6C5B7B", "#355C7D"],
        "foreground": C["tinta"], "foregroundNeutralSecondary": C["suave"], "foregroundNeutralTertiary": "#8A99A3",
        "background": C["branco"], "backgroundLight": C["fundo"], "backgroundNeutral": C["borda"],
        "tableAccent": C["alfa"], "good": C["verde"], "neutral": C["ambar"], "bad": C["vermelho"],
        "maximum": C["alfa"], "center": C["claro"], "minimum": "#F1F7F9",
        "textClasses": {
            "callout": {"fontSize": 22, "fontFace": "Segoe UI Semibold", "color": C["tinta"]},
            "title": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": C["tinta"]},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": C["tinta"]},
            "label": {"fontSize": 10, "fontFace": "Segoe UI", "color": C["suave"]},
        },
        "visualStyles": {"page": {"*": {"background": [{"color": {"solid": {"color": C["fundo"]}}, "transparency": 0}],
                                        "outspace": [{"color": {"solid": {"color": C["fundo"]}}, "transparency": 0}]}}},
    }


def relatorio(nome_proj, paginas):
    base = OUT / f"{nome_proj}.Report"
    w(base / "definition.pbir", {"$schema": V_PBIR, "version": "4.0",
                                 "datasetReference": {"byPath": {"path": f"../{nome_proj}.SemanticModel"}}})
    w(base / ".platform", {"$schema": V_PLATFORM, "metadata": {"type": "Report", "displayName": nome_proj},
                           "config": {"version": "2.0", "logicalId": gid()}})
    d = base / "definition"
    w(d / "version.json", {"$schema": V_VERSION, "version": "2.0.0"})
    arq_tema = f"TemaSSO-{uuid.uuid4().hex[:8]}.json"
    w(d / "report.json", {
        "$schema": V_REPORT,
        "themeCollection": {
            "baseTheme": {"name": "CY24SU10", "reportVersionAtImport": "5.61", "type": "SharedResources"},
            "customTheme": {"name": arq_tema, "reportVersionAtImport": "5.61", "type": "RegisteredResources"}},
        "layoutOptimization": "None",
        "objects": {"section": props(verticalAlignment=s("Top")),
                    "outspacePane": props(visible=b(True), expanded=b(False))},
        "resourcePackages": [
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": arq_tema, "path": arq_tema, "type": "CustomTheme"}]},
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": "CY24SU10", "path": "BaseThemes/CY24SU10.json", "type": "BaseTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True, "useDefaultAggregateDisplayName": True},
    })
    w(base / "StaticResources" / "RegisteredResources" / arq_tema, tema(arq_tema))
    destino = base / "StaticResources" / "SharedResources" / "BaseThemes" / "CY24SU10.json"
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(BASE_THEME, destino)
    w(d / "pages" / "pages.json", {"$schema": V_PAGES, "pageOrder": [p.nome for p in paginas],
                                   "activePageName": paginas[0].nome})
    for p in paginas:
        pj = {"$schema": V_PAGE, "name": p.nome, "displayName": p.exibicao, "displayOption": "FitToPage",
              "height": 720, "width": 1280}
        if p.filtros:
            pj["filterConfig"] = {"filters": p.filtros}
        w(d / "pages" / p.nome / "page.json", pj)
        for v in p.visuais:
            w(d / "pages" / p.nome / "visuals" / v["name"] / "visual.json", v)
    w(OUT / f"{nome_proj}.pbip", {"$schema": V_PBIP, "version": "1.0",
                                  "artifacts": [{"report": {"path": f"{nome_proj}.Report"}}],
                                  "settings": {"enableAutoRecovery": True}})


def m_tabela_fixa(colunas, linhas):
    tipo_m = {"text": "text", "int": "Int64.Type", "num": "number"}
    tipo = ", ".join(f'#"{c}" = {tipo_m[t]}' for c, t in colunas)
    vals = ", ".join("{" + ", ".join(json.dumps(v, ensure_ascii=False) for v in row) + "}" for row in linhas)
    return f"let\n    Fonte = #table(type table [{tipo}], {{{vals}}})\nin\n    Fonte"


def tab_fixa(nome, colunas, linhas, ordenar_por=None):
    defs = []
    for c, t in colunas:
        if ordenar_por and c != ordenar_por and t == "text":
            defs.append((c, t, False, None, ordenar_por))
        elif c == ordenar_por:
            defs.append((c, t, True))
        else:
            defs.append((c, t))
    return tabela(nome, defs, m_tabela_fixa(colunas, linhas))


# =============================================================================
# PROJETO 1 — CONTROLE DE ASOs
# =============================================================================
def projeto_asos():
    nome = "Controle_ASOs_DEMO"
    arq = "Base_Controle_ASOs_DEMO.xlsx"
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
    cols_asos = t_asos + [("DIAS PARA VENCER", "int"), ("NA BASE", "bool"), ("FAIXA 90D", "text"),
                          ("SITUAÇÃO AGENDAMENTO", "text"), ("DIAS ANTECIPAÇÃO AGENDADO", "int")]

    def extra_mes(campo):
        return f"""    Mes = Table.AddColumn(Tipos, "MÊS", each Date.ToText([{campo}], "MMM/yy", "pt-BR"), type text),
    MesOrdem = Table.AddColumn(Mes, "MÊS ORDEM", each Date.Year([{campo}]) * 100 + Date.Month([{campo}]), Int64.Type)"""
    mes_cols = [("MÊS", "text", False, None, "MÊS ORDEM"), ("MÊS ORDEM", "int", True)]

    t_real = [("CPF", "text"), ("NOME", "text"), ("EMPRESA", "text"), (UN, "text"), ("TIPO DE ASO", "text"),
              ("DATA DE REALIZAÇÃO", "date"), ("VENCIMENTO ANTERIOR", "date"), ("DIAS ANTECIPADOS", "int"),
              ("FAIXA DE ANTECIPAÇÃO", "text")]
    t_faltas = [("CPF", "text"), ("NOME", "text"), ("EMPRESA", "text"), (UN, "text"), ("DATA DO AGENDAMENTO", "date"),
                ("TIPO DE EXAME", "text"), ("SITUAÇÃO", "text")]
    t_desl = [("CPF", "text"), ("NOME", "text"), ("CHAPA", "text"), ("EMPRESA", "text"), (UN, "text"),
              ("DATA DE DESLIGAMENTO", "date"), ("ÚLTIMO ASO", "date"), ("ORIGEM", "text"), ("REGISTRADO EM", "date")]
    t_hist = [("DATA DO SNAPSHOT", "date"), (UN, "text"), ("VENCIMENTOS 90 DIAS", "int"), ("VENCIDOS", "int"),
              ("AGENDADOS", "int"), ("REALIZADOS", "int")]

    F1, F2, F3, F4, F5 = "1. Vencimentos", "2. Histórico", "3. Antecipação", "4. Faltas", "5. Desligamentos"
    medidas = "\n".join([
        medida("Colaboradores na Base", "CALCULATE ( COUNTROWS ( tblASOS ), tblASOS[NA BASE] = TRUE () )", "#,0", F1,
               "Ativos, em férias ou licença remunerada; afastados e exceções ficam fora."),
        medida("Vencimentos 90d", """CALCULATE (
    COUNTROWS ( tblASOS ),
    tblASOS[NA BASE] = TRUE (),
    tblASOS[DIAS PARA VENCER] <= 90
)""", "#,0", F1, "ASOs vencidos ou que vencem nos próximos 90 dias, na base de cálculo."),
        medida("ASOs Vencidos", 'CALCULATE ( [Vencimentos 90d], tblASOS[FAIXA 90D] = "Vencido" )', "#,0", F1),
        medida("ASOs Agendados", 'CALCULATE ( [Vencimentos 90d], tblASOS[SITUAÇÃO AGENDAMENTO] = "Agendado" )', "#,0", F1),
        medida("Realizados a Atualizar", 'CALCULATE ( [Vencimentos 90d], tblASOS[SITUAÇÃO AGENDAMENTO] = "Realizado" )',
               "#,0", F1, "Exame feito, ASO novo ainda não lançado na base."),
        medida("Sem Agendamento", 'CALCULATE ( [Vencimentos 90d], tblASOS[SITUAÇÃO AGENDAMENTO] = "Sem agendamento" )',
               "#,0", F1),
        medida("% Cobertura", "DIVIDE ( [ASOs Agendados] + [Realizados a Atualizar], [Vencimentos 90d] )", "0.0%", F1,
               "Parte dos vencimentos em 90 dias que já tem exame agendado ou realizado."),
        medida("Pendentes Críticos", """CALCULATE (
    COUNTROWS ( tblASOS ),
    tblASOS[NA BASE] = TRUE (),
    tblASOS[SITUAÇÃO AGENDAMENTO] = "Sem agendamento",
    tblASOS[DIAS PARA VENCER] <= 30
)""", "#,0", F1, "Sem agendamento e vencendo em até 30 dias (ou já vencidos)."),
        medida("Dias até Vencer (pendente)", "IF ( [Pendentes Críticos] > 0, MIN ( tblASOS[DIAS PARA VENCER] ) )", "#,0",
               F1, "Usada na lista de pendências: só aparece para quem está pendente."),
        medida("Vencidos (semanal)", "SUM ( tblHistorico[VENCIDOS] )", "#,0", F2),
        medida("% Cobertura (semanal)", """DIVIDE (
    SUM ( tblHistorico[AGENDADOS] ) + SUM ( tblHistorico[REALIZADOS] ),
    SUM ( tblHistorico[VENCIMENTOS 90 DIAS] )
)""", "0.0%", F2),
        medida("ASOs Realizados", "COUNTROWS ( tblRealizados )", "#,0", F3),
        medida("Antecipação Média (dias)", "AVERAGE ( tblRealizados[DIAS ANTECIPADOS] )", "#,0", F3,
               "Dias entre o exame e o vencimento anterior; negativo quando já estava vencido."),
        medida("Realizados Já Vencidos", "CALCULATE ( COUNTROWS ( tblRealizados ), tblRealizados[DIAS ANTECIPADOS] < 0 )",
               "#,0", F3),
        medida("% Realizado Já Vencido", "DIVIDE ( [Realizados Já Vencidos], [ASOs Realizados] )", "0.0%", F3),
        medida("Antecipação Agendados (dias)",
               "CALCULATE ( AVERAGE ( tblASOS[DIAS ANTECIPAÇÃO AGENDADO] ), tblASOS[NA BASE] = TRUE () )", "#,0", F3),
        medida("Total de Faltas", "COUNTROWS ( tblFaltas )", "#,0", F4),
        medida("Pessoas que Faltaram", "DISTINCTCOUNT ( tblFaltas[CPF] )", "#,0", F4),
        medida("Reincidentes 2+", """COUNTROWS (
    FILTER (
        VALUES ( tblFaltas[CPF] ),
        CALCULATE ( COUNTROWS ( tblFaltas ) ) >= 2
    )
)""", "#,0", F4, "Colaboradores com duas faltas ou mais no período."),
        medida("% Reincidentes", "DIVIDE ( [Reincidentes 2+], [Pessoas que Faltaram] )", "0.0%", F4),
        medida("Faltas em Demissional", 'CALCULATE ( [Total de Faltas], tblFaltas[TIPO DE EXAME] = "DEMISSIONAL" )', "#,0", F4),
        medida("Desligamentos Registrados", "COUNTROWS ( tblDesligados )", "#,0", F5,
               "Registrados automaticamente pelo fluxo que lê o comunicado de desligamento."),
        medida("Desligamentos por Dia Útil", """VAR Inicio = MIN ( tblDesligados[DATA DE DESLIGAMENTO] )
VAR Fim = MAX ( tblDesligados[DATA DE DESLIGAMENTO] )
RETURN
    IF (
        NOT ISBLANK ( Inicio ),
        DIVIDE (
            [Desligamentos Registrados],
            COUNTROWS ( FILTER ( CALENDAR ( Inicio, Fim ), WEEKDAY ( [Date], 2 ) < 6 ) )
        )
    )""", "#,0.0", F5),
        medida("Horas Manuais Evitadas", "[Desligamentos Registrados] * 12 / 60", "#,0", F5,
               "Estimativa: cerca de 12 minutos de trabalho manual por desligamento."),
    ])

    tabs = {
        "tblASOS": tabela("tblASOS", cols_asos, m_excel(arq, "tblASOS", t_asos, extra_asos),
                          descricao="Um colaborador por linha (base fictícia)."),
        "tblRealizados": tabela("tblRealizados", t_real + mes_cols, m_excel(arq, "tblRealizados", t_real,
                                                                             extra_mes("DATA DE REALIZAÇÃO"))),
        "tblFaltas": tabela("tblFaltas", t_faltas + mes_cols, m_excel(arq, "tblFaltas", t_faltas,
                                                                       extra_mes("DATA DO AGENDAMENTO"))),
        "tblDesligados": tabela("tblDesligados", t_desl + mes_cols, m_excel(arq, "tblDesligados", t_desl,
                                                                             extra_mes("DATA DE DESLIGAMENTO"))),
        "tblHistorico": tabela("tblHistorico", t_hist, m_excel(arq, "tblHistorico", t_hist)),
        "dEmpresa": tabela("dEmpresa", [("EMPRESA", "text")],
                           'let\n    Fonte = Table.Distinct(Table.SelectColumns(tblASOS, {"EMPRESA"}))\nin\n    Fonte'),
        "dUnidade": tabela("dUnidade", [(UN, "text")],
                           f'let\n    Fonte = Table.Distinct(Table.SelectColumns(tblASOS, {{"{UN}"}}))\nin\n    Fonte'),
        "dFaixa": tab_fixa("dFaixa", [("FAIXA", "text"), ("ORDEM", "int")],
                           [[f, i] for i, f in enumerate(["Vencido", "Até 7 dias", "8 a 30 dias", "31 a 60 dias",
                                                          "61 a 90 dias", "Acima de 90 dias"], 1)], "ORDEM"),
        "dFaixaAntecipacao": tab_fixa("dFaixaAntecipacao", [("FAIXA", "text"), ("ORDEM", "int")],
                                      [[f, i] for i, f in enumerate(["Já vencido", "Até 7 dias", "8 a 30 dias",
                                                                     "31 a 60 dias", "61 a 90 dias", "91 a 120 dias"], 1)],
                                      "ORDEM"),
        "dSituacao": tab_fixa("dSituacao", [("SITUAÇÃO", "text"), ("ORDEM", "int")],
                              [["Realizado", 1], ["Agendado", 2], ["Sem agendamento", 3]], "ORDEM"),
        M: tabela_medidas(M, medidas),
    }
    rels = [(("tblASOS", "EMPRESA"), ("dEmpresa", "EMPRESA")), (("tblASOS", UN), ("dUnidade", UN)),
            (("tblASOS", "FAIXA 90D"), ("dFaixa", "FAIXA")), (("tblASOS", "SITUAÇÃO AGENDAMENTO"), ("dSituacao", "SITUAÇÃO")),
            (("tblRealizados", "EMPRESA"), ("dEmpresa", "EMPRESA")), (("tblRealizados", UN), ("dUnidade", UN)),
            (("tblRealizados", "FAIXA DE ANTECIPAÇÃO"), ("dFaixaAntecipacao", "FAIXA")),
            (("tblFaltas", "EMPRESA"), ("dEmpresa", "EMPRESA")), (("tblFaltas", UN), ("dUnidade", UN)),
            (("tblDesligados", "EMPRESA"), ("dEmpresa", "EMPRESA")), (("tblDesligados", UN), ("dUnidade", UN)),
            (("tblHistorico", UN), ("dUnidade", UN))]
    ordem = ["PastaDados", "DataReferencia", "tblASOS", "tblRealizados", "tblFaltas", "tblDesligados", "tblHistorico",
             "dEmpresa", "dUnidade", "dFaixa", "dFaixaAntecipacao", "dSituacao", M]
    modelo(nome, tabs, ordem, expressoes([("PastaDados", f'"{PASTA_PADRAO}"', "Text"),
                                          ("DataReferencia", "#date(2026, 10, 26)", "Date")]), relacionamentos(rels))

    # ---------------------------------------------------------------- relatório
    MARCA = "Gestão de ASOs"
    NP = 5
    sl = [("filtroEmpresa", 236, v_slicer("dEmpresa", "EMPRESA", "Empresa", "HorizontalList", "Empresa")),
          ("filtroUnidade", 220, v_slicer("dUnidade", UN, "Unidade de negócio", "Dropdown", "Unidade"))]
    sit = col("dSituacao", "SITUAÇÃO")
    cores_sit = {"Realizado": C["alfa"], "Agendado": C["claro"], "Sem agendamento": C["cinza"]}
    faixa = col("dFaixa", "FAIXA")
    unidade = col("dUnidade", UN)
    pgs = []

    p = Pagina("VisaoGeral", "Visão geral")
    moldura(p, "Visão geral dos ASOs", "Vencimentos nos próximos 90 dias · referência 26/10/2026", MARCA, NP, sl)
    cards(p, [("kVenc", v_card("Vencimentos 90d", "Vencem em até 90 dias", C["alfa"])),
              ("kVencidos", v_card("ASOs Vencidos", "Já vencidos", C["vermelho"], C["vermelho"])),
              ("kCob", v_card("% Cobertura", "Cobertura", C["novo"], C["novo"])),
              ("kAg", v_card("ASOs Agendados", "Agendados", C["claro"])),
              ("kSem", v_card("Sem Agendamento", "Sem agendamento", C["ambar"]))])
    p.add("barFaixas", X0, 192, 620, 256, v_grafico(
        "barChart", faixa, [med(M, "Vencimentos 90d")], "Situação por faixa de vencimento", serie=sit,
        cores_serie=cores_sit, ordenar=faixa, desc=False, sub="O trecho cinza ainda precisa de agendamento"))
    p.add("roscaSituacao", X0 + 632, 192, 408, 256, v_rosca(sit, med(M, "Vencimentos 90d"), "Como estão os vencimentos",
                                                             cores_sit))
    p.add("comboSemanal", X0, 460, LARG, 248, v_grafico(
        "lineClusteredColumnComboChart", col("tblHistorico", "DATA DO SNAPSHOT"),
        [(med(M, "Vencidos (semanal)"), "Vencidos")], "Evolução nas últimas 12 semanas",
        y2=[(med(M, "% Cobertura (semanal)"), "Cobertura")],
        ordenar=col("tblHistorico", "DATA DO SNAPSHOT"), desc=False,
        sub="Vencidos caem enquanto a cobertura sobe: o efeito dos lembretes automáticos"))
    pgs.append(p)

    p = Pagina("PorUnidade", "Por unidade")
    moldura(p, "Onde agir primeiro", "Mapa de calor por unidade e lista de pendências críticas", MARCA, NP, sl)
    cards(p, [("kBase", v_card("Colaboradores na Base", "Colaboradores na base", C["alfa"])),
              ("kVenc2", v_card("Vencimentos 90d", "Vencem em até 90 dias", C["claro"])),
              ("kCob2", v_card("% Cobertura", "Cobertura", C["novo"], C["novo"])),
              ("kCrit", v_card("Pendentes Críticos", "Pendentes críticos (≤ 30 dias)", C["vermelho"], C["vermelho"]))])
    p.add("mapaCalor", X0, 192, 620, 300, v_matriz(unidade, faixa, med(M, "Vencimentos 90d"),
                                                   "Vencimentos por unidade e faixa", grad=("#F1F7F9", "#5FA8C1"),
                                                   sub="Quanto mais escuro, mais ASOs vencendo"))
    p.add("barCobertura", X0 + 632, 192, 408, 300, v_grafico(
        "barChart", unidade, [med(M, "% Cobertura")], "Cobertura por unidade (clique para filtrar)",
        grad=gradiente("% Cobertura", C["vermelho"], C["novo"], C["ambar"]), ordenar=med(M, "% Cobertura"), desc=False,
        sub="Pior cobertura no topo"))
    p.add("tabPendentes", X0, 504, LARG, 204, v_tabela(
        [(col("tblASOS", "NOME"), "Colaborador"), (col("tblASOS", UN), "Unidade"), (col("tblASOS", "EMPRESA"), "Empresa"),
         (col("tblASOS", "GESTOR DIRETO"), "Gestor"), (col("tblASOS", "DATA DE VENCIMENTO"), "Vencimento"),
         (med(M, "Dias até Vencer (pendente)"), "Dias para vencer")],
        "Pendências críticas: sem agendamento e vencendo em até 30 dias", ordenar=med(M, "Dias até Vencer (pendente)"),
        desc=False, grad=[(med(M, "Dias até Vencer (pendente)"), "#F4C7C3", "#FFFFFF", None)]))
    pgs.append(p)

    p = Pagina("Antecipacao", "Antecipação")
    moldura(p, "Antecipação dos exames", "ASOs realizados de julho a outubro de 2026", MARCA, NP, sl)
    cards(p, [("kReal", v_card("ASOs Realizados", "ASOs realizados", C["alfa"])),
              ("kAnt", v_card("Antecipação Média (dias)", "Antecipação média (dias)", C["novo"])),
              ("kJaVenc", v_card("% Realizado Já Vencido", "Feitos depois de vencer", C["vermelho"], C["vermelho"])),
              ("kAntAg", v_card("Antecipação Agendados (dias)", "Antecipação dos agendados (dias)", C["claro"]))])
    p.add("barFaixaAnt", X0, 192, 500, 516, v_grafico(
        "barChart", col("dFaixaAntecipacao", "FAIXA"), [med(M, "ASOs Realizados")], "Realizados por faixa de antecipação",
        cor_unica=C["alfa"], ordenar=col("dFaixaAntecipacao", "FAIXA"), desc=False,
        sub="Quanto antes do vencimento o exame foi feito"))
    p.add("comboMes", X0 + 512, 192, 528, 256, v_grafico(
        "lineClusteredColumnComboChart", col("tblRealizados", "MÊS"), [(med(M, "ASOs Realizados"), "Realizados")],
        "Realizados e antecipação média por mês", y2=[(med(M, "Antecipação Média (dias)"), "Antecipação (dias)")],
        ordenar=col("tblRealizados", "MÊS"), desc=False))
    p.add("barAntUnidade", X0 + 512, 460, 528, 248, v_grafico(
        "clusteredBarChart", unidade, [(med(M, "Antecipação Média (dias)"), "Realizados"),
                                      (med(M, "Antecipação Agendados (dias)"), "Agendados")],
        "Antecipação média por unidade: realizados x agendados",
        cores_y=[(med(M, "Antecipação Média (dias)"), C["cinza"]), (med(M, "Antecipação Agendados (dias)"), C["alfa"])],
        ordenar=med(M, "Antecipação Média (dias)")))
    pgs.append(p)

    p = Pagina("Faltas", "Faltas")
    moldura(p, "Faltas em exames agendados", "Agosto a outubro de 2026 · nomes fictícios", MARCA, NP, sl)
    cards(p, [("kFaltas", v_card("Total de Faltas", "Faltas", C["vermelho"], C["vermelho"])),
              ("kPessoas", v_card("Pessoas que Faltaram", "Pessoas que faltaram", C["alfa"])),
              ("kReinc", v_card("Reincidentes 2+", "Reincidentes (2+ faltas)", C["ambar"])),
              ("kPctReinc", v_card("% Reincidentes", "Taxa de reincidência", C["ambar"])),
              ("kDemiss", v_card("Faltas em Demissional", "Faltas em demissional", C["beta"]))])
    p.add("barFaltasUnid", X0, 192, 340, 516, v_grafico(
        "barChart", unidade, [med(M, "Total de Faltas")], "Faltas por unidade (clique para filtrar)",
        cor_unica=C["vermelho"], ordenar=med(M, "Total de Faltas")))
    p.add("roscaTipo", X0 + 352, 192, 340, 256, v_rosca(col("tblFaltas", "TIPO DE EXAME"), med(M, "Total de Faltas"),
                                                         "Faltas por tipo de exame", sub="Demissional atrasa a rescisão"))
    p.add("colFaltasMes", X0 + 352, 460, 340, 248, v_grafico(
        "columnChart", col("tblFaltas", "MÊS"), [med(M, "Total de Faltas")], "Faltas por mês", cor_unica=C["vermelho"],
        ordenar=col("tblFaltas", "MÊS"), desc=False))
    p.add("tabFaltas", X0 + 704, 192, 336, 516, v_tabela(
        [(col("tblFaltas", "NOME"), "Colaborador"), (col("tblFaltas", UN), "Unidade"), (med(M, "Total de Faltas"), "Faltas")],
        "Quem mais faltou", ordenar=med(M, "Total de Faltas"), barras=[(med(M, "Total de Faltas"), C["vermelho"])]))
    pgs.append(p)

    p = Pagina("Desligamentos", "Desligamentos")
    moldura(p, "Desligamentos registrados automaticamente",
            "Fluxo do Power Automate lê o comunicado e move o colaborador para a aba de desligados", MARCA, NP, sl)
    cards(p, [("kDesl", v_card("Desligamentos Registrados", "Desligamentos registrados", C["alfa"])),
              ("kDia", v_card("Desligamentos por Dia Útil", "Média por dia útil", C["claro"])),
              ("kHoras", v_card("Horas Manuais Evitadas", "Horas manuais evitadas (estimativa)", C["novo"], C["novo"]))])
    p.add("colDeslMes", X0, 192, 620, 256, v_grafico(
        "columnChart", col("tblDesligados", "MÊS"), [med(M, "Desligamentos Registrados")], "Desligamentos por mês",
        cor_unica=C["alfa"], ordenar=col("tblDesligados", "MÊS"), desc=False))
    p.add("roscaDeslEmp", X0 + 632, 192, 408, 256, v_rosca(col("dEmpresa", "EMPRESA"), med(M, "Desligamentos Registrados"),
                                                            "Por empresa", {"EMPRESA ALFA": C["alfa"],
                                                                            "EMPRESA BETA": C["beta"]}))
    p.add("barDeslUnid", X0, 460, 400, 248, v_grafico(
        "barChart", unidade, [med(M, "Desligamentos Registrados")], "Por unidade", cor_unica=C["alfa"],
        ordenar=med(M, "Desligamentos Registrados")))
    p.add("tabDesl", X0 + 412, 460, 628, 248, v_tabela(
        [(col("tblDesligados", "NOME"), "Colaborador"), (col("tblDesligados", UN), "Unidade"),
         (col("tblDesligados", "DATA DE DESLIGAMENTO"), "Desligamento"), (col("tblDesligados", "ÚLTIMO ASO"), "Último ASO"),
         (med(M, "Desligamentos Registrados"), "Registros")],
        "Últimos registros", ordenar=col("tblDesligados", "DATA DE DESLIGAMENTO")))
    pgs.append(p)

    assert len(pgs) == NP
    relatorio(nome, pgs)


# =============================================================================
# PROJETO 2 — CUSTOS DE EXAMES (estrutura espelhada do BI real, com nomes fictícios)
# =============================================================================
def projeto_custos():
    nome = "Custos_Exames_DEMO"
    arq = "Base_Custos_Exames_DEMO.xlsx"
    alfa, beta = "EMPRESA ALFA", "EMPRESA BETA"

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
    t_emp = [("EMPRESA", "text"), ("VIDAS", "int"), ("PER CAPITA ATUAL", "num"), ("PER CAPITA NOVO", "num")]
    t_prem = [("ITEM", "text"), ("REGRA", "text")]

    rs, rs2 = "\\R\\$ #,0", "\\R\\$ #,0.00"
    B0, B1, B2, B3, B4, B5, B6 = ("0. Base", "1. Atual x Novo", "2. Por empresa", "3. Cenários", "4. Preços e rede",
                                  "5. Negociação", "6. Ranking")
    vol_preco = "fExameEmpresa[VOLUME MENSAL] * fExameEmpresa[PREÇO ATUAL]"
    vol_novo = "fExameEmpresa[VOLUME MENSAL] * RELATED ( dExame[NOVO PREÇO] )"
    medidas = "\n".join([
        # base
        medida("Fator Período", "SELECTEDVALUE ( dPeriodo[FATOR], 1 )", "0", B0, "1 para mensal, 12 para anual."),
        medida("Inclui Per Capita", "SELECTEDVALUE ( dPerCapita[INCLUI], 1 )", "0", B0, "1 com per capita, 0 sem."),
        medida("Volume do Mês", "SUM ( fExameEmpresa[VOLUME MENSAL] )", "#,0", B0),
        medida("Total de Vidas", "SUM ( dEmpresa[VIDAS] )", "#,0", B0),
        medida("Qtd Exames", "COUNTROWS ( dExame )", "#,0", B0),
        medida("Exames Atual (mês)", f"SUMX ( fExameEmpresa, {vol_preco} )", rs, B0),
        medida("Exames Novo (mês)", f"SUMX ( fExameEmpresa, {vol_novo} )", rs, B0),
        medida("Per Capita Atual (mês)", "SUMX ( dEmpresa, dEmpresa[VIDAS] * dEmpresa[PER CAPITA ATUAL] )", rs, B0),
        medida("Per Capita Novo (mês)", "SUMX ( dEmpresa, dEmpresa[VIDAS] * dEmpresa[PER CAPITA NOVO] )", rs, B0),
        # atual x novo
        medida("Custo Exames Atual", "[Exames Atual (mês)] * [Fator Período]", rs, B1),
        medida("Custo Exames Novo", "[Exames Novo (mês)] * [Fator Período]", rs, B1),
        medida("Custo Per Capita Atual", "[Per Capita Atual (mês)] * [Inclui Per Capita] * [Fator Período]", rs, B1),
        medida("Custo Per Capita Novo", "[Per Capita Novo (mês)] * [Inclui Per Capita] * [Fator Período]", rs, B1),
        medida("Custo Atual", "[Custo Exames Atual] + [Custo Per Capita Atual]", rs, B1,
               "Exames na tabela atual de cada empresa + per capita (se incluído)."),
        medida("Custo Novo", "[Custo Exames Novo] + [Custo Per Capita Novo]", rs, B1,
               "Exames na tabela unificada negociada + per capita novo (se incluído)."),
        medida("Economia", "[Custo Atual] - [Custo Novo]", rs, B1),
        medida("Economia %", "DIVIDE ( [Economia], [Custo Atual] )", "0.0%", B1),
        medida("Economia Exames", "[Custo Exames Atual] - [Custo Exames Novo]", rs, B1),
        medida("Economia Per Capita", "[Custo Per Capita Atual] - [Custo Per Capita Novo]", rs, B1),
        # por empresa / comparativo
        medida("Custo Atual Alfa", f'CALCULATE ( [Custo Atual], dEmpresa[EMPRESA] = "{alfa}" )', rs, B2),
        medida("Custo Atual Beta", f'CALCULATE ( [Custo Atual], dEmpresa[EMPRESA] = "{beta}" )', rs, B2),
        medida("Parte Exames", f"""SWITCH (
    SELECTEDVALUE ( dComparativo[ORDEM] ),
    1, CALCULATE ( [Custo Exames Atual], dEmpresa[EMPRESA] = "{alfa}" ),
    2, CALCULATE ( [Custo Exames Atual], dEmpresa[EMPRESA] = "{beta}" ),
    3, [Custo Exames Atual],
    4, [Custo Exames Novo]
)""", rs, B2, "Parte de exames em cada coluna do comparativo."),
        medida("Parte Per Capita", f"""SWITCH (
    SELECTEDVALUE ( dComparativo[ORDEM] ),
    1, CALCULATE ( [Custo Per Capita Atual], dEmpresa[EMPRESA] = "{alfa}" ),
    2, CALCULATE ( [Custo Per Capita Atual], dEmpresa[EMPRESA] = "{beta}" ),
    3, [Custo Per Capita Atual],
    4, [Custo Per Capita Novo]
)""", rs, B2, "Parte de per capita em cada coluna do comparativo."),
        # cenários
        medida("Custo Cenário", f"""VAR Cenario = SELECTEDVALUE ( dCenario[ORDEM] )
VAR Vidas = SUM ( dEmpresa[VIDAS] )
VAR PcAlfa = CALCULATE ( MAX ( dEmpresa[PER CAPITA ATUAL] ), dEmpresa[EMPRESA] = "{alfa}" )
VAR PcBeta = CALCULATE ( MAX ( dEmpresa[PER CAPITA ATUAL] ), dEmpresa[EMPRESA] = "{beta}" )
VAR ExAlfa = SUMX ( fExameEmpresa, fExameEmpresa[VOLUME MENSAL] * COALESCE ( fExameEmpresa[PREÇO ALFA], fExameEmpresa[PREÇO ATUAL] ) )
VAR ExBeta = SUMX ( fExameEmpresa, fExameEmpresa[VOLUME MENSAL] * COALESCE ( fExameEmpresa[PREÇO BETA], fExameEmpresa[PREÇO ATUAL] ) )
VAR ExMenor = SUMX ( fExameEmpresa, fExameEmpresa[VOLUME MENSAL] * fExameEmpresa[MENOR PREÇO] )
VAR Mensal =
    SWITCH (
        Cenario,
        1, [Exames Atual (mês)] + [Per Capita Atual (mês)] * [Inclui Per Capita],
        2, ExAlfa + Vidas * PcAlfa * [Inclui Per Capita],
        3, ExBeta + Vidas * PcBeta * [Inclui Per Capita],
        4, ExMenor + Vidas * MIN ( PcAlfa, PcBeta ) * [Inclui Per Capita],
        5, [Exames Novo (mês)] + [Per Capita Novo (mês)] * [Inclui Per Capita]
    )
RETURN
    Mensal * [Fator Período]""", rs, B3, "Custo total em cada cenário de tabela; respeita empresa, período e per capita."),
        medida("Variação vs Atual %", "DIVIDE ( [Custo Cenário] - [Custo Atual], [Custo Atual] )", "+0.0%;-0.0%;0.0%", B3),
        medida("Valor Ponte", f"""VAR Etapa = SELECTEDVALUE ( dPonte[ORDEM] )
VAR Comum = SUMX ( FILTER ( fExameEmpresa, RELATED ( dExame[EMPRESAS QUE REALIZAM] ) = "Ambas" ), fExameEmpresa[VOLUME MENSAL] * ( fExameEmpresa[PREÇO ATUAL] - RELATED ( dExame[NOVO PREÇO] ) ) )
VAR SoAlfa = SUMX ( FILTER ( fExameEmpresa, RELATED ( dExame[EMPRESAS QUE REALIZAM] ) = "Só Alfa" ), fExameEmpresa[VOLUME MENSAL] * ( fExameEmpresa[PREÇO ATUAL] - RELATED ( dExame[NOVO PREÇO] ) ) )
VAR SoBeta = SUMX ( FILTER ( fExameEmpresa, RELATED ( dExame[EMPRESAS QUE REALIZAM] ) = "Só Beta" ), fExameEmpresa[VOLUME MENSAL] * ( fExameEmpresa[PREÇO ATUAL] - RELATED ( dExame[NOVO PREÇO] ) ) )
VAR PerCapita = SUMX ( dEmpresa, dEmpresa[VIDAS] * ( dEmpresa[PER CAPITA ATUAL] - dEmpresa[PER CAPITA NOVO] ) ) * [Inclui Per Capita]
RETURN
    SWITCH (
        Etapa,
        1, [Custo Atual],
        2, -Comum * [Fator Período],
        3, -SoAlfa * [Fator Período],
        4, -SoBeta * [Fator Período],
        5, -PerCapita * [Fator Período]
    )""", rs, B3, "Etapas do custo atual ao negociado; o total da cascata é o custo negociado."),
        medida("Economia por Origem", "IF ( SELECTEDVALUE ( dPonte[ORDEM] ) > 1, - [Valor Ponte] )", rs, B3),
        # preços e rede
        medida("Preço Tabela Alfa", f'CALCULATE ( MAX ( fExameEmpresa[PREÇO ATUAL] ), dEmpresa[EMPRESA] = "{alfa}" )', rs2, B4),
        medida("Preço Tabela Beta", f'CALCULATE ( MAX ( fExameEmpresa[PREÇO ATUAL] ), dEmpresa[EMPRESA] = "{beta}" )', rs2, B4),
        medida("Preço Negociado", "MAX ( dExame[NOVO PREÇO] )", rs2, B4),
        medida("Preço Médio Alfa", f"""DIVIDE (
    CALCULATE ( [Exames Atual (mês)], dEmpresa[EMPRESA] = "{alfa}" ),
    CALCULATE ( [Volume do Mês], dEmpresa[EMPRESA] = "{alfa}" )
)""", rs2, B4, "Preço médio ponderado pelo volume."),
        medida("Preço Médio Beta", f"""DIVIDE (
    CALCULATE ( [Exames Atual (mês)], dEmpresa[EMPRESA] = "{beta}" ),
    CALCULATE ( [Volume do Mês], dEmpresa[EMPRESA] = "{beta}" )
)""", rs2, B4),
        medida("Preço Médio Negociado", "DIVIDE ( [Exames Novo (mês)], [Volume do Mês] )", rs2, B4),
        medida("Diferença Beta vs Alfa %", """VAR A = [Preço Tabela Alfa]
VAR B = [Preço Tabela Beta]
RETURN
    IF ( NOT ISBLANK ( A ) && NOT ISBLANK ( B ), DIVIDE ( B - A, A ) )""", "+0.0%;-0.0%;0.0%", B4),
        medida("Redução vs Menor Preço %", """VAR A = [Preço Tabela Alfa]
VAR B = [Preço Tabela Beta]
VAR Menor = IF ( ISBLANK ( A ), B, IF ( ISBLANK ( B ), A, MIN ( A, B ) ) )
RETURN
    DIVIDE ( Menor - [Preço Negociado], Menor )""", "0.0%", B4),
        # negociação / exclusivos
        medida("Exames Reduzidos", 'CALCULATE ( [Qtd Exames], dExame[SITUAÇÃO DA NEGOCIAÇÃO] = "Reduzido" )', "#,0", B5),
        medida("Exames sem Negociação", 'CALCULATE ( [Qtd Exames], dExame[SITUAÇÃO DA NEGOCIAÇÃO] = "Sem negociação" )',
               "#,0", B5),
        medida("Custo Novo sem Negociação",
               'CALCULATE ( [Custo Exames Novo], dExame[SITUAÇÃO DA NEGOCIAÇÃO] = "Sem negociação" )', rs, B5,
               "Custo dos exames que mantiveram preço: onde ainda há espaço para negociar."),
        medida("Qtd Exames Exclusivos", 'CALCULATE ( [Qtd Exames], dExame[EMPRESAS QUE REALIZAM] <> "Ambas" )', "#,0", B5),
        medida("Qtd Exames em Comum", 'CALCULATE ( [Qtd Exames], dExame[EMPRESAS QUE REALIZAM] = "Ambas" )', "#,0", B5),
        medida("Custo Exclusivos Atual", 'CALCULATE ( [Custo Exames Atual], dExame[EMPRESAS QUE REALIZAM] <> "Ambas" )', rs, B5),
        medida("Custo Exclusivos Novo", 'CALCULATE ( [Custo Exames Novo], dExame[EMPRESAS QUE REALIZAM] <> "Ambas" )', rs, B5),
        medida("Economia Exames em Comum", 'CALCULATE ( [Economia Exames], dExame[EMPRESAS QUE REALIZAM] = "Ambas" )', rs, B5),
        # ranking
        medida("Top N", "SELECTEDVALUE ( dTopN[N], 10 )", "0", B6),
        medida("Ranking Custo Novo", """IF (
    NOT ISBLANK ( [Custo Exames Novo] ),
    RANKX ( ALL ( dExame[EXAME], dExame[GRUPO] ), [Custo Exames Novo], , DESC, DENSE )
)""", "0", B6),
        medida("Custo Novo Top N", "IF ( [Ranking Custo Novo] <= [Top N], [Custo Exames Novo] )", rs, B6),
        medida("% Acumulado Top N", """VAR Valor = [Custo Exames Novo]
VAR Total = CALCULATE ( [Custo Exames Novo], ALL ( dExame ) )
VAR Acumulado =
    SUMX ( FILTER ( ALL ( dExame[EXAME], dExame[GRUPO] ), [Custo Exames Novo] >= Valor ), [Custo Exames Novo] )
RETURN
    IF ( [Ranking Custo Novo] <= [Top N], DIVIDE ( Acumulado, Total ) )""", "0.0%", B6, "Curva de Pareto dos exames."),
        medida("% do Total Top N", """IF (
    [Ranking Custo Novo] <= [Top N],
    DIVIDE ( [Custo Exames Novo], CALCULATE ( [Custo Exames Novo], ALL ( dExame ) ) )
)""", "0.0%", B6),
        medida("Concentração Top N", """DIVIDE (
    SUMX ( FILTER ( ALL ( dExame[EXAME], dExame[GRUPO] ), [Ranking Custo Novo] <= [Top N] ), [Custo Exames Novo] ),
    CALCULATE ( [Custo Exames Novo], ALL ( dExame ) )
)""", "0.0%", B6, "Parte do custo concentrada nos N exames mais caros."),
    ])

    tabs = {
        "dExame": tabela("dExame", [("ID_EXAME", "text"), ("EXAME", "text"), ("GRUPO", "text"), ("REDE", "text"),
                                    ("NOVO PREÇO", "num"), ("SITUAÇÃO DA NEGOCIAÇÃO", "text"),
                                    ("EMPRESAS QUE REALIZAM", "text")], m_dexame,
                         descricao="Catálogo de exames com o novo preço negociado."),
        "dEmpresa": tabela("dEmpresa", t_emp, m_excel(arq, "dEmpresa", t_emp)),
        "fExameEmpresa": tabela("fExameEmpresa", [("EMPRESA", "text"), ("ID_EXAME", "text"), ("VOLUME MENSAL", "int"),
                                                  ("PREÇO ATUAL", "num"), ("PREÇO ALFA", "num", True),
                                                  ("PREÇO BETA", "num", True), ("MENOR PREÇO", "num", True)], m_fato,
                                descricao="Fato: volume mensal e preço atual por empresa e exame."),
        "dPeriodo": tab_fixa("dPeriodo", [("PERÍODO", "text"), ("FATOR", "int"), ("ORDEM", "int")],
                             [["Mensal", 1, 1], ["Anual", 12, 2]], "ORDEM"),
        "dPerCapita": tab_fixa("dPerCapita", [("PER CAPITA", "text"), ("INCLUI", "int"), ("ORDEM", "int")],
                               [["Com per capita", 1, 1], ["Sem per capita", 0, 2]], "ORDEM"),
        "dCenario": tab_fixa("dCenario", [("CENÁRIO", "text"), ("ORDEM", "int")],
                             [["Atual (contratos separados)", 1], ["Tabela da Alfa para todos", 2],
                              ["Tabela da Beta para todos", 3], ["Menor preço entre tabelas", 4],
                              ["Novo preço negociado", 5]], "ORDEM"),
        "dComparativo": tab_fixa("dComparativo", [("VISÃO", "text"), ("ORDEM", "int")],
                                 [["Alfa · atual", 1], ["Beta · atual", 2], ["Alfa + Beta · atual", 3],
                                  ["Unificado · novo preço", 4]], "ORDEM"),
        "dPonte": tab_fixa("dPonte", [("ETAPA", "text"), ("ORDEM", "int")],
                           [["Custo atual", 1], ["Exames em comum", 2], ["Exclusivos da Alfa", 3],
                            ["Exclusivos da Beta", 4], ["Per capita por vida", 5]], "ORDEM"),
        "dTopN": tab_fixa("dTopN", [("N", "int")], [[5], [10], [15], [20]]),
        "tblPremissas": tabela("tblPremissas", t_prem, m_excel(arq, "tblPremissas", t_prem)),
        M: tabela_medidas(M, medidas),
    }
    rels = [(("fExameEmpresa", "ID_EXAME"), ("dExame", "ID_EXAME")), (("fExameEmpresa", "EMPRESA"), ("dEmpresa", "EMPRESA"))]
    ordem = ["PastaDados", "dExame", "dEmpresa", "fExameEmpresa", "dPeriodo", "dPerCapita", "dCenario", "dComparativo",
             "dPonte", "dTopN", "tblPremissas", M]
    modelo(nome, tabs, ordem, expressoes([("PastaDados", f'"{PASTA_PADRAO}"', "Text")]), relacionamentos(rels))

    # ---------------------------------------------------------------- relatório
    MARCA = "Custos de Exames"
    NP = 12
    f_per = lambda: ("fPeriodo", 196, v_slicer("dPeriodo", "PERÍODO", "Período", "HorizontalList", "Periodo", unico=True))
    f_pc = lambda: ("fPerCapita", 268, v_slicer("dPerCapita", "PER CAPITA", "Per capita", "HorizontalList", "PerCapita",
                                                unico=True))
    sl = [f_per(), f_pc()]
    exame, grupo, rede = col("dExame", "EXAME"), col("dExame", "GRUPO"), col("dExame", "REDE")
    empresa = col("dEmpresa", "EMPRESA")
    cores_emp = {alfa: C["alfa"], beta: C["beta"]}
    pgs = []

    # 1. Resumo
    p = Pagina("Resumo", "Resumo")
    moldura(p, "Custo atual x tabela unificada", "Duas empresas, um prestador · valores conforme período e per capita",
            MARCA, NP, sl)
    cards(p, [("kAtual", v_card("Custo Atual", "Custo atual (Alfa + Beta)", C["cinza"])),
              ("kNovo", v_card("Custo Novo", "Custo com tabela unificada", C["novo"])),
              ("kEcon", v_card("Economia", "Economia", C["novo"], C["novo"])),
              ("kEconPct", v_card("Economia %", "Redução sobre o atual", C["novo"], C["novo"]))])
    comp = col("dComparativo", "VISÃO")
    p.add("comparativo", X0, 192, 512, 516, v_grafico(
        "columnChart", comp, [(med(M, "Parte Exames"), "Exames"), (med(M, "Parte Per Capita"), "Per capita")],
        "Alfa, Beta e unificado", cores_y=[(med(M, "Parte Exames"), C["alfa"]), (med(M, "Parte Per Capita"), C["claro"])],
        ordenar=comp, desc=False, totais=True, sub="Composição do custo: exames + per capita"))
    p.add("cascata", X0 + 524, 192, 516, 300, v_cascata(col("dPonte", "ETAPA"), med(M, "Valor Ponte"),
                                                        "Do custo atual ao negociado",
                                                        sub="Barra final (Total) = custo com tabela unificada"))
    p.add("roscaOrigem", X0 + 524, 504, 516, 204, v_rosca(col("dPonte", "ETAPA"), med(M, "Economia por Origem"),
                                                          "De onde vem a economia"))
    pgs.append(p)

    # 2 e 3. Empresas
    for chave, nome_emp, cor_emp in (("Alfa", alfa, C["alfa"]), ("Beta", beta, C["beta"])):
        p = Pagina(f"Empresa{chave}", f"Empresa {chave}", filtros=[filtro("dEmpresa", "EMPRESA", [nome_emp])])
        moldura(p, f"Empresa {chave}: hoje x tabela unificada", f"Página filtrada para a Empresa {chave}", MARCA, NP, sl)
        cards(p, [(f"kAt{chave}", v_card("Custo Atual", "Custo atual", cor_emp)),
                  (f"kNv{chave}", v_card("Custo Novo", "Com tabela unificada", C["novo"])),
                  (f"kEc{chave}", v_card("Economia", "Economia", C["novo"], C["novo"])),
                  (f"kEp{chave}", v_card("Economia %", "Redução", C["novo"], C["novo"])),
                  (f"kVi{chave}", v_card("Total de Vidas", "Vidas", C["cinza"]))])
        p.add(f"barExames{chave}", X0, 192, 620, 516, v_grafico(
            "clusteredBarChart", exame, [(med(M, "Custo Exames Atual"), "Atual"), (med(M, "Custo Exames Novo"), "Novo")],
            "Custo por exame: atual x novo", cores_y=[(med(M, "Custo Exames Atual"), cor_emp),
                                                       (med(M, "Custo Exames Novo"), C["novo"])],
            ordenar=med(M, "Custo Exames Atual"), rotulos=False, sub="Ordenado pelo custo atual"))
        p.add(f"barGrupo{chave}", X0 + 632, 192, 408, 256, v_grafico(
            "barChart", grupo, [med(M, "Custo Exames Atual")], "Custo atual por grupo de exame", cor_unica=cor_emp,
            ordenar=med(M, "Custo Exames Atual"), fmt_rotulo=1000))
        p.add(f"tab{chave}", X0 + 632, 460, 408, 248, v_tabela(
            [(exame, "Exame"), (med(M, "Volume do Mês"), "Vol./mês"), (med(M, "Economia Exames"), "Economia")],
            "Economia por exame", ordenar=med(M, "Economia Exames"), barras=[(med(M, "Economia Exames"), C["novo"])]))
        pgs.append(p)

    # 4. Unificado
    p = Pagina("Unificado", "Unificado")
    moldura(p, "Empresa unificada: uma tabela só", "Volume das duas empresas no preço negociado", MARCA, NP, sl)
    cards(p, [("kUnNovo", v_card("Custo Novo", "Custo unificado", C["novo"])),
              ("kUnVol", v_card("Volume do Mês", "Exames por mês", C["alfa"])),
              ("kUnVidas", v_card("Total de Vidas", "Vidas", C["claro"])),
              ("kUnExames", v_card("Qtd Exames", "Exames na tabela", C["cinza"]))])
    p.add("colGrupoEmpresa", X0, 192, 620, 516, v_grafico(
        "columnChart", grupo, [med(M, "Custo Exames Novo")], "Custo unificado por grupo e empresa", serie=empresa,
        cores_serie=cores_emp, ordenar=med(M, "Custo Exames Novo"), totais=True, rotulos=False,
        sub="Quem consome cada grupo de exame na tabela unificada"))
    p.add("tabUnificado", X0 + 632, 192, 408, 516, v_tabela(
        [(exame, "Exame"), (med(M, "Volume do Mês"), "Vol./mês"), (med(M, "Preço Negociado"), "Preço"),
         (med(M, "Custo Exames Novo"), "Custo")], "Tabela unificada", ordenar=med(M, "Custo Exames Novo"),
        barras=[(med(M, "Custo Exames Novo"), C["novo"])]))
    pgs.append(p)

    # 5. Cenários
    p = Pagina("Cenarios", "Cenários A × B")
    moldura(p, "Cenários de tabela: Alfa × Beta", "E se uma empresa adotasse a tabela da outra?", MARCA, NP, sl)
    cen = col("dCenario", "CENÁRIO")
    p.add("barCenarios", X0, 92, 620, 320, v_grafico(
        "barChart", cen, [med(M, "Custo Cenário")], "Custo total por cenário", serie=empresa, cores_serie=cores_emp,
        ordenar=cen, desc=False, totais=True, rotulos=False, sub="Empilhado por empresa"))
    p.add("barVariacao", X0 + 632, 92, 408, 320, v_grafico(
        "barChart", cen, [med(M, "Variação vs Atual %")], "Variação contra o custo atual",
        grad=gradiente("Variação vs Atual %", C["novo"], C["vermelho"], "#F2F2F2"), ordenar=cen, desc=False,
        sub="Negativo = mais barato"))
    p.add("matrizCenarios", X0, 424, LARG, 284, v_matriz(cen, empresa, med(M, "Custo Cenário"),
                                                         "Custo por cenário e empresa",
                                                         sub="O total da linha é o custo das duas empresas",
                                                         total_linhas=False))
    pgs.append(p)

    # 6. Exclusivos
    p = Pagina("Exclusivos", "Exames exclusivos",
               filtros=[filtro("dExame", "EMPRESAS QUE REALIZAM", ["Ambas"], negar=True)])
    moldura(p, "Exames que só uma empresa realiza", "Página filtrada: exames exclusivos de cada empresa", MARCA, NP, sl)
    cards(p, [("kExQtd", v_card("Qtd Exames", "Exames exclusivos", C["alfa"])),
              ("kExAt", v_card("Custo Exames Atual", "Custo atual", C["cinza"])),
              ("kExNv", v_card("Custo Exames Novo", "Custo com tabela unificada", C["novo"])),
              ("kExEc", v_card("Economia Exames", "Economia", C["novo"], C["novo"]))])
    quem = col("dExame", "EMPRESAS QUE REALIZAM")
    p.add("roscaExcl", X0, 192, 400, 516, v_rosca(quem, med(M, "Custo Exames Atual"), "Custo exclusivo por empresa",
                                                  {"Só Alfa": C["alfa"], "Só Beta": C["beta"]}))
    p.add("tabExcl", X0 + 412, 192, 628, 516, v_tabela(
        [(exame, "Exame"), (quem, "Empresa"), (rede, "Rede"), (med(M, "Volume do Mês"), "Vol./mês"),
         (med(M, "Custo Exames Atual"), "Atual"), (med(M, "Custo Exames Novo"), "Novo"),
         (med(M, "Economia Exames"), "Economia")], "Lista de exames exclusivos", ordenar=med(M, "Custo Exames Atual"),
        barras=[(med(M, "Custo Exames Atual"), C["cinza"])]))
    pgs.append(p)

    # 7. Preço por rede
    p = Pagina("PrecoRede", "Preço por rede")
    moldura(p, "Comparativo de preço de tabela", "Rede própria x credenciada, preço médio ponderado pelo volume",
            MARCA, NP, sl)
    precos = [(med(M, "Preço Médio Alfa"), "Alfa"), (med(M, "Preço Médio Beta"), "Beta"),
              (med(M, "Preço Médio Negociado"), "Negociado")]
    cores_precos = [(med(M, "Preço Médio Alfa"), C["alfa"]), (med(M, "Preço Médio Beta"), C["beta"]),
                    (med(M, "Preço Médio Negociado"), C["novo"])]
    p.add("colRede", X0, 92, 400, 320, v_grafico("clusteredColumnChart", rede, precos, "Preço médio por rede",
                                                 cores_y=cores_precos, ordenar=rede, desc=False))
    p.add("barGrupo", X0 + 412, 92, 628, 320, v_grafico("clusteredBarChart", grupo, precos, "Preço médio por grupo",
                                                         cores_y=cores_precos, ordenar=med(M, "Preço Médio Negociado"),
                                                         rotulos=False))
    p.add("tabPrecos", X0, 424, LARG, 284, v_tabela(
        [(exame, "Exame"), (rede, "Rede"), (med(M, "Preço Tabela Alfa"), "Tabela Alfa"),
         (med(M, "Preço Tabela Beta"), "Tabela Beta"), (med(M, "Diferença Beta vs Alfa %"), "Beta x Alfa"),
         (med(M, "Preço Negociado"), "Negociado"), (med(M, "Redução vs Menor Preço %"), "Redução vs menor")],
        "Preço de tabela por exame", ordenar=med(M, "Diferença Beta vs Alfa %"),
        grad=[(med(M, "Diferença Beta vs Alfa %"), "#CFE8DD", "#F6C9C2", "#FFFFFF")],
        barras=[(med(M, "Redução vs Menor Preço %"), C["novo"])]))
    pgs.append(p)

    # 8. Novo preço · impacto
    p = Pagina("Impacto", "Novo preço · impacto")
    moldura(p, "Novo preço: impacto total", "Quanto a tabela negociada economiza e de onde vem", MARCA, NP, sl)
    cards(p, [("kImpEx", v_card("Economia Exames", "Economia em exames", C["novo"], C["novo"])),
              ("kImpPc", v_card("Economia Per Capita", "Economia em per capita", C["claro"])),
              ("kImpTot", v_card("Economia", "Economia total", C["novo"], C["novo"])),
              ("kImpRed", v_card("Exames Reduzidos", "Exames com preço reduzido", C["alfa"]))])
    p.add("cascataImp", X0, 192, 620, 516, v_cascata(col("dPonte", "ETAPA"), med(M, "Valor Ponte"),
                                                     "Ponte do custo atual ao negociado"))
    p.add("roscaImp", X0 + 632, 192, 408, 256, v_rosca(col("dPonte", "ETAPA"), med(M, "Economia por Origem"),
                                                       "Composição da economia"))
    p.add("colEmpImp", X0 + 632, 460, 408, 248, v_grafico(
        "clusteredColumnChart", empresa, [(med(M, "Custo Atual"), "Atual"), (med(M, "Custo Novo"), "Novo")],
        "Atual x novo por empresa", cores_y=[(med(M, "Custo Atual"), C["cinza"]), (med(M, "Custo Novo"), C["novo"])],
        ordenar=empresa, desc=False, fmt_rotulo=1000))
    pgs.append(p)

    # 9. Novo preço · exames em comum
    p = Pagina("Comum", "Novo preço · em comum", filtros=[filtro("dExame", "EMPRESAS QUE REALIZAM", ["Ambas"])])
    moldura(p, "Exames em comum", "Exames que as duas empresas realizam: três preços lado a lado", MARCA, NP, sl)
    cards(p, [("kCmQtd", v_card("Qtd Exames", "Exames em comum", C["alfa"])),
              ("kCmAt", v_card("Custo Exames Atual", "Custo atual", C["cinza"])),
              ("kCmNv", v_card("Custo Exames Novo", "Custo novo", C["novo"])),
              ("kCmEc", v_card("Economia Exames", "Economia", C["novo"], C["novo"]))])
    p.add("barComum", X0, 192, 400, 516, v_grafico(
        "barChart", exame, [med(M, "Economia Exames")], "Economia por exame", cor_unica=C["novo"],
        ordenar=med(M, "Economia Exames"), rotulos=True, fmt_rotulo=1))
    p.add("tabComum", X0 + 412, 192, 628, 516, v_tabela(
        [(exame, "Exame"), (med(M, "Preço Tabela Alfa"), "Alfa"), (med(M, "Preço Tabela Beta"), "Beta"),
         (med(M, "Preço Negociado"), "Negociado"), (med(M, "Redução vs Menor Preço %"), "Redução vs menor"),
         (med(M, "Economia Exames"), "Economia")], "Três preços por exame", ordenar=med(M, "Economia Exames"),
        barras=[(med(M, "Economia Exames"), C["novo"])]))
    pgs.append(p)

    # 10. Novo preço · negociação
    p = Pagina("Negociacao", "Novo preço · negociação")
    moldura(p, "Itens sem negociação x reduções", "Onde a negociação reduziu preço e onde ainda há espaço", MARCA, NP, sl)
    situ = col("dExame", "SITUAÇÃO DA NEGOCIAÇÃO")
    cores_situ = {"Reduzido": C["novo"], "Sem negociação": C["ambar"]}
    cards(p, [("kNgRed", v_card("Exames Reduzidos", "Exames com redução", C["novo"])),
              ("kNgSem", v_card("Exames sem Negociação", "Exames sem negociação", C["ambar"])),
              ("kNgCusto", v_card("Custo Novo sem Negociação", "Custo ainda sem negociação", C["ambar"])),
              ("kNgEc", v_card("Economia Exames", "Economia em exames", C["novo"], C["novo"]))])
    p.add("roscaNeg", X0, 192, 400, 256, v_rosca(situ, med(M, "Qtd Exames"), "Exames por situação", cores_situ))
    p.add("colNeg", X0, 460, 400, 248, v_grafico(
        "clusteredColumnChart", situ, [(med(M, "Custo Exames Atual"), "Atual"), (med(M, "Custo Exames Novo"), "Novo")],
        "Custo atual x novo por situação", cores_y=[(med(M, "Custo Exames Atual"), C["cinza"]),
                                                    (med(M, "Custo Exames Novo"), C["novo"])], ordenar=situ, desc=False,
        fmt_rotulo=1000))
    p.add("tabSemNeg", X0 + 412, 192, 628, 516, v_tabela(
        [(exame, "Exame"), (rede, "Rede"), (med(M, "Volume do Mês"), "Vol./mês"), (med(M, "Preço Negociado"), "Preço"),
         (med(M, "Custo Exames Novo"), "Custo novo")], "Exames que mantiveram o preço: próximos alvos",
        ordenar=med(M, "Custo Exames Novo"), barras=[(med(M, "Custo Exames Novo"), C["ambar"])]),
        filtros=[filtro("dExame", "SITUAÇÃO DA NEGOCIAÇÃO", ["Sem negociação"])])
    pgs.append(p)

    # 11. Novo preço · maiores custos (Pareto)
    p = Pagina("MaioresCustos", "Novo preço · maiores custos")
    moldura(p, "Maiores custos na tabela unificada", "Ranking e curva de Pareto · escolha quantos exames ver", MARCA, NP,
            [("fTopN", 236, v_slicer("dTopN", "N", "Top N", "HorizontalList", None, unico=True)), f_per()])
    cards(p, [("kTopConc", v_card("Concentração Top N", "Parte do custo nos N maiores", C["alfa"])),
              ("kTopCusto", v_card("Custo Exames Novo", "Custo de exames (novo)", C["novo"])),
              ("kTopQtd", v_card("Qtd Exames", "Exames na tabela", C["cinza"]))])
    p.add("pareto", X0, 192, LARG, 300, v_grafico(
        "lineClusteredColumnComboChart", exame, [(med(M, "Custo Novo Top N"), "Custo novo")],
        "Pareto dos exames mais caros", y2=[(med(M, "% Acumulado Top N"), "% acumulado")],
        ordenar=med(M, "Custo Novo Top N"), rotulos=False))
    p.add("tabRanking", X0, 504, LARG, 204, v_tabela(
        [(med(M, "Ranking Custo Novo"), "#"), (exame, "Exame"), (grupo, "Grupo"), (med(M, "Custo Novo Top N"), "Custo novo"),
         (med(M, "% do Total Top N"), "% do total"), (med(M, "% Acumulado Top N"), "% acumulado")],
        "Ranking", ordenar=med(M, "Custo Novo Top N"), barras=[(med(M, "Custo Novo Top N"), C["alfa"])]))
    pgs.append(p)

    # 12. Premissas
    p = Pagina("Premissas", "Premissas")
    moldura(p, "Premissas e regras de cálculo", "Como os números deste relatório são calculados", MARCA, NP)
    p.add("tabPremissas", X0, 92, LARG, 300, v_tabela([(col("tblPremissas", "ITEM"), "Item"),
                                                       (col("tblPremissas", "REGRA"), "Regra")], "Premissas"))
    p.add("notaModelo", X0, 408, LARG, 300, v_texto([
        [("Modelo", 12, C["tinta"], True)],
        [("Estrela: a fato fExameEmpresa (volume e preço por empresa e exame) liga-se a dExame e dEmpresa. ", 10,
          C["suave"], False)],
        [("dPeriodo, dPerCapita, dCenario, dComparativo, dPonte e dTopN são tabelas de parâmetro desconectadas, lidas "
          "com SELECTEDVALUE dentro das medidas. Assim os mesmos visuais respondem a mensal/anual e com/sem per capita.",
          10, C["suave"], False)],
        [("Técnicas", 12, C["tinta"], True)],
        [("SUMX exame a exame, CALCULATE por empresa, RANKX com Top N dinâmico, curva de Pareto, cascata de economia, "
          "filtros de página e slicers sincronizados.", 10, C["suave"], False)],
        [("Origem", 12, C["tinta"], True)],
        [("Planilha Excel fictícia carregada por Power Query; o caminho fica no parâmetro PastaDados.", 10, C["suave"],
          False)],
    ]))
    pgs.append(p)

    assert len(pgs) == NP, len(pgs)
    relatorio(nome, pgs)


if __name__ == "__main__":
    if OUT.exists():
        shutil.rmtree(OUT)
    projeto_asos()
    projeto_custos()
    w(OUT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n", False)
    print("ok", sum(1 for f in OUT.rglob("*") if f.is_file()), "arquivos")
