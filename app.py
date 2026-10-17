"""Dashboard Streamlit – Levantamento de BD (CDVs e Sinais, Vias 1 e 2) + CRUZAMENTOS.

Rodar:  streamlit run app.py
Os dois .xlsx ficam na pasta  dados/  (ou podem ser enviados pela barra lateral).
"""
import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Levantamento BD • Vias 1 e 2", page_icon="🚆",
                   layout="wide", initial_sidebar_state="collapsed")
st.markdown("""<style>
.block-container{padding-top:1.5rem;max-width:1500px}
[data-testid=stMetricValue]{font-size:2.4rem;font-weight:700}
[data-testid=stMetricLabel]{font-size:1rem}
</style>""", unsafe_allow_html=True)

# ───────────────────────── caminhos (relativos: funcionam em qualquer PC) ─────────────────────────
BASE = Path(__file__).parent
PASTA = BASE / "dados"
ARQ_CDV = PASTA / "levantamento_BD_CDVs_via_1_e_2.xlsx"
ARQ_SIN = PASTA / "levantamento_BD_Sinais_via_1_e_2.xlsx"

COR = {"Preenchido": "#2E9E6B", "Não utilizado": "#B8C0CC", "Pendente": "#E8833A"}
FUNC_CDV = [("Codificação", "NCTP"), ("Codificação", "SCTP"), ("Codificação", "N12-CTP"),
            ("Indicação", "TPR"), ("Indicação", "TPRN")]
FUNC_SIN = [("Vermelho", "RGE"), ("Vermelho", "RGENX"), ("Amarelo", "YGE"), ("Amarelo", "YGENX")]
PADRAO = {"Locação": re.compile(r"^L[BC] [A-Z]{1,2}-\d+$"),
          "Armário TB": re.compile(r"^TB\d-[A-Z]\d+$"),
          "Interface": re.compile(r"^B\d{2} \d{2}$")}


# ───────────────────────── leitura e tratamento ─────────────────────────
def txt(v):
    return "" if pd.isna(v) else str(v).strip()


def status_de(v):  # vazio = pendente | "-" = não utilizado
    return "Pendente" if v == "" else "Não utilizado" if v == "-" else "Preenchido"


def tipo_local(loc):
    l = loc.lower()
    for ini, t in (("locação", "Locação"), ("armári", "Armário TB"), ("armari", "Armário TB"),
                   ("interface", "Interface"), ("bastidor", "Bastidor")):
        if l.startswith(ini):
            return t
    return "Outro"


def fora_padrao(tipo, ponto):
    p = PADRAO.get(tipo)
    if not p or ponto in ("", "-"):
        return False
    return not all(p.match(x.strip()) for x in ponto.split("/"))


def familia(tipo, ponto):
    if ponto in ("", "-"):
        return ""
    p = ponto.split("/")[0].strip()
    if tipo == "Locação":
        return p[:2]
    if tipo == "Armário TB":
        m = re.match(r"TB\d*", p)
        return m.group() if m else "Fora do padrão"
    return tipo


def parse(raw, ativo, prefixo, cols, funcs, dados):
    """Lê os blocos lado a lado e devolve uma tabela 'longa' (1 linha por ponto)."""
    regs, cont = [], {}
    for c0 in cols:
        r = 0
        while r < len(raw):
            cel = txt(raw.iat[r, c0])
            if not cel.upper().startswith(prefixo):
                r += 1
                continue
            nome = re.sub(r"\s*LUZ$", "", cel[len(prefixo):].strip(), flags=re.I).strip()
            cont[nome] = cont.get(nome, 0) + 1
            bid = nome if cont[nome] == 1 else f"{nome} (#{cont[nome]})"
            m = re.match(r"^(\d)([NS])", nome) if ativo == "CDV" else None
            via = f"Via {m[1]}" if m else "—"
            sent = {"N": "Norte", "S": "Sul"}[m[2]] if m else "—"
            codigos = [txt(raw.iat[r + 3, c0 + 1 + k]) for k in range(len(funcs))]
            rr = r + dados
            while rr < len(raw):
                loc = txt(raw.iat[rr, c0])
                if not loc or loc.upper().startswith(prefixo):
                    break
                tipo = tipo_local(loc)
                local = loc if tipo in ("Locação", "Bastidor") else tipo
                for k, (cat, fun) in enumerate(funcs):
                    ponto = txt(raw.iat[rr, c0 + 1 + k])
                    cod = codigos[k] or fun
                    if ativo == "Sinal" and cod == fun:
                        cod = nome + fun
                    regs.append(dict(ativo=ativo, id=bid, nome=nome, via=via, sentido=sent,
                                     local=local, tipo_local=tipo, categoria=cat, funcao=fun,
                                     codigo=cod, ponto=ponto, status=status_de(ponto),
                                     familia=familia(tipo, ponto),
                                     fora_padrao=fora_padrao(tipo, ponto)))
                rr += 1
            r = rr
    for x in regs:
        x["duplicado"] = cont[x["nome"]] > 1
    return pd.DataFrame(regs)


@st.cache_data
def ler(b):
    return pd.read_excel(BytesIO(b), header=None)


def carregar(caminho, rotulo):
    if caminho.exists():
        return caminho.read_bytes()
    up = st.sidebar.file_uploader(rotulo, type="xlsx")
    return up.getvalue() if up else None


# ───────────────────────── cruzamento ─────────────────────────
def chave(p):
    """Normaliza o ponto para comparar: maiúsculas, sem espaços duplicados."""
    return re.sub(r"\s+", " ", p.strip().upper())


def pontos_unitarios(d):
    """Explode células como 'LB A-20 / LB B-11' em um ponto por linha (só preenchidos)."""
    d = d[d.status == "Preenchido"].copy()
    d["ponto_un"] = d.ponto.str.split("/")
    d = d.explode("ponto_un")
    d["ponto_un"] = d.ponto_un.map(chave)
    return d


def placa(p):
    """Interface 'B13 22' -> placa 'B13'."""
    m = re.match(r"^B\s?(\d{2})", p)
    return f"B{m[1]}" if m else None


b_cdv, b_sin = carregar(ARQ_CDV, "Planilha de CDVs"), carregar(ARQ_SIN, "Planilha de Sinais")
if b_cdv is None or b_sin is None:
    st.info("Coloque os dois arquivos .xlsx na pasta **dados/** ou envie-os pela barra lateral.")
    st.stop()

df = pd.concat([parse(ler(b_cdv), "CDV", "CDV ", (1, 8), FUNC_CDV, 4),
                parse(ler(b_sin), "Sinal", "SINAL ", (0, 6, 12, 18), FUNC_SIN, 5)],
               ignore_index=True)

# ───────────────────────── filtros ─────────────────────────
with st.sidebar:
    st.header("Filtros")
    vias_all = sorted(df.via[df.ativo == "CDV"].unique())
    vias = st.multiselect("Via (CDVs)", vias_all, default=vias_all)
    locs_all = sorted(df.tipo_local.unique())
    locs = st.multiselect("Tipo de local", locs_all, default=locs_all)
f = df[df.tipo_local.isin(locs) & ((df.ativo != "CDV") | df.via.isin(vias))]
cdv, sin = f[f.ativo == "CDV"], f[f.ativo == "Sinal"]

# ───────────────────────── cabeçalho / KPIs ─────────────────────────
st.title("🚆 Levantamento de BD — Vias 1 e 2")
st.caption("CDVs e Sinais · **Preenchido** = ponto informado · **Não utilizado** = “-” · "
           "**Pendente** = célula vazia")

pre, pen = (f.status == "Preenchido").sum(), (f.status == "Pendente").sum()
taxa = pre / (pre + pen) if (pre + pen) else 0
k = st.columns(5)
k[0].metric("CDVs mapeados", cdv.nome.nunique())
k[1].metric("Sinais mapeados", sin.nome.nunique())
k[2].metric("Pontos preenchidos", f"{pre:,}".replace(",", "."))
k[3].metric("Taxa de preenchimento", f"{taxa:.1%}")
k[4].metric("Pendências", int(pen))


def mostrar(fig, h=420):
    fig.update_layout(template="plotly_white", height=h, margin=dict(t=60, b=20, l=10, r=10),
                      legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)


def aba_ativo(d, nome, ordem):
    if d.empty:
        st.info("Sem dados para os filtros selecionados.")
        return
    ok = d[d.status == "Preenchido"]
    ids = list(ok.groupby("id").size().sort_values(ascending=False).index)
    p = d.groupby(["id", "status"]).size().reset_index(name="pontos")
    mostrar(px.bar(p, x="id", y="pontos", color="status", color_discrete_map=COR,
                   category_orders={"id": ids}, title=f"Pontos por {nome}"), 400)
    c1, c2 = st.columns(2)
    with c1:
        q = ok.groupby(["funcao", "tipo_local"]).size().reset_index(name="pontos")
        mostrar(px.bar(q, x="funcao", y="pontos", color="tipo_local", text_auto=True,
                       category_orders={"funcao": ordem},
                       title="Pontos preenchidos por função e tipo de local"))
    with c2:
        q = ok.groupby(["tipo_local", "familia"]).size().reset_index(name="pontos")
        mostrar(px.sunburst(q, path=["tipo_local", "familia"], values="pontos",
                            title="Composição dos pontos (local → família)"))
    h = ok.pivot_table(index="id", columns="funcao", values="ponto", aggfunc="count",
                       fill_value=0).reindex(columns=ordem, fill_value=0).sort_index()
    if not h.empty:
        fig = px.imshow(h, text_auto=True, aspect="auto", color_continuous_scale="Teal",
                        title=f"Mapa de calor: {nome} × função (nº de pontos preenchidos)")
        mostrar(fig, max(360, 24 * len(h) + 120))


# ───────────────────────── abas ─────────────────────────
t1, t2, t3, t4, t5, t6 = st.tabs(["Visão geral", "CDVs", "Sinais", "🔗 Cruzamentos",
                                  "Qualidade dos dados", "Dados"])

with t1:
    c1, c2 = st.columns(2)
    with c1:
        a = f.groupby(["ativo", "status"]).size().reset_index(name="pontos")
        mostrar(px.bar(a, x="ativo", y="pontos", color="status", color_discrete_map=COR,
                       text_auto=True, title="Situação dos pontos por tipo de ativo"))
    with c2:
        b = f[f.status == "Preenchido"].groupby(["tipo_local", "ativo"]).size().reset_index(name="pontos")
        mostrar(px.bar(b, x="tipo_local", y="pontos", color="ativo", barmode="group",
                       text_auto=True, title="Pontos preenchidos por tipo de local"))
    c = f[f.status == "Preenchido"].groupby(["ativo", "tipo_local", "familia"]).size().reset_index(name="pontos")
    mostrar(px.sunburst(c, path=["ativo", "tipo_local", "familia"], values="pontos",
                        title="De onde vêm os pontos (ativo → local → família)"), 520)

with t2:
    aba_ativo(cdv, "CDV", [x[1] for x in FUNC_CDV])
with t3:
    aba_ativo(sin, "Sinal", [x[1] for x in FUNC_SIN])

with t4:
    st.subheader("Cruzamento CDVs × Sinais")
    st.caption("Dois registros são considerados *semelhantes* quando apontam para o mesmo ponto "
               "físico (mesma locação, armário TB ou interface), ignorando maiúsculas e espaços extras. "
               "Células com “/” são separadas em um ponto por linha.")
    uc, us = pontos_unitarios(cdv), pontos_unitarios(sin)

    # 1) pontos em comum entre CDV e Sinal
    comuns = sorted(set(uc.ponto_un) & set(us.ponto_un))
    a_ = (uc[uc.ponto_un.isin(comuns)].groupby("ponto_un")
          .agg(CDVs=("id", lambda s: ", ".join(sorted(set(s)))),
               funcoes_cdv=("funcao", lambda s: ", ".join(sorted(set(s))))))
    b_ = (us[us.ponto_un.isin(comuns)].groupby("ponto_un")
          .agg(Sinais=("id", lambda s: ", ".join(sorted(set(s)))),
               funcoes_sinal=("funcao", lambda s: ", ".join(sorted(set(s))))))
    cruz = a_.join(b_).reset_index().rename(columns={"ponto_un": "Ponto"})

    # 2) placas de interface em comum (ex.: B13 usada por CDV e por Sinal)
    ic, i_s = uc[uc.tipo_local == "Interface"].copy(), us[us.tipo_local == "Interface"].copy()
    ic["placa"], i_s["placa"] = ic.ponto_un.map(placa), i_s.ponto_un.map(placa)
    placas = sorted((set(ic.placa.dropna()) & set(i_s.placa.dropna())))

    m = st.columns(4)
    m[0].metric("Pontos únicos CDV", uc.ponto_un.nunique())
    m[1].metric("Pontos únicos Sinal", us.ponto_un.nunique())
    m[2].metric("Pontos em comum", len(comuns))
    m[3].metric("Placas de interface em comum", len(placas))

    c1, c2 = st.columns(2)
    with c1:
        v = pd.DataFrame({"Grupo": ["Só CDV", "Em comum", "Só Sinal"],
                          "Pontos": [len(set(uc.ponto_un) - set(us.ponto_un)), len(comuns),
                                     len(set(us.ponto_un) - set(uc.ponto_un))]})
        mostrar(px.bar(v, x="Grupo", y="Pontos", text_auto=True, color="Grupo",
                       color_discrete_sequence=["#2F6FDE", "#2E9E6B", "#E8833A"],
                       title="Pontos exclusivos × compartilhados"), 380)
    with c2:
        ambos = pd.concat([uc.assign(o="CDV"), us.assign(o="Sinal")])
        t = (ambos.groupby(["tipo_local", "o"]).ponto_un.nunique().reset_index(name="pontos_unicos"))
        mostrar(px.bar(t, x="tipo_local", y="pontos_unicos", color="o", barmode="group", text_auto=True,
                       title="Pontos únicos por tipo de local"), 380)

    st.markdown("##### Pontos usados por CDV **e** Sinal ao mesmo tempo")
    if cruz.empty:
        st.info("Nenhum ponto em comum para os filtros selecionados.")
    else:
        st.dataframe(cruz, hide_index=True, use_container_width=True)

    st.markdown("##### Reuso de pontos dentro do mesmo ativo (mesmo ponto em vários blocos)")
    ativo_sel = st.radio("Ativo", ["CDV", "Sinal"], horizontal=True)
    base = uc if ativo_sel == "CDV" else us
    reuso = (base.groupby("ponto_un").agg(blocos=("id", "nunique"),
                                          ids=("id", lambda s: ", ".join(sorted(set(s)))),
                                          funcoes=("funcao", lambda s: ", ".join(sorted(set(s)))))
             .query("blocos > 1").sort_values("blocos", ascending=False).reset_index()
             .rename(columns={"ponto_un": "Ponto"}))
    st.caption("Reuso de LB A-19, LB A-21… na função N12-CTP é esperado (código comum). "
               "Reuso em outras funções merece conferência.")
    st.dataframe(reuso, hide_index=True, use_container_width=True)

    st.markdown("##### Matriz CDV × Sinal (nº de pontos compartilhados)")
    if comuns:
        pares = (uc[uc.ponto_un.isin(comuns)][["id", "ponto_un"]].rename(columns={"id": "CDV"})
                 .merge(us[us.ponto_un.isin(comuns)][["id", "ponto_un"]].rename(columns={"id": "Sinal"}),
                        on="ponto_un").drop_duplicates())
        mat = pares.pivot_table(index="CDV", columns="Sinal", values="ponto_un",
                                aggfunc="count", fill_value=0)
        mostrar(px.imshow(mat, text_auto=True, aspect="auto", color_continuous_scale="Teal",
                          title="Quais CDVs dividem pontos com quais Sinais"), max(320, 40 * len(mat) + 120))
    else:
        st.info("Sem pontos compartilhados para montar a matriz.")

    st.markdown("##### Placas de interface compartilhadas")
    if placas:
        pl = pd.concat([ic[ic.placa.isin(placas)].assign(ativo="CDV"),
                        i_s[i_s.placa.isin(placas)].assign(ativo="Sinal")])
        st.dataframe(pl.groupby(["placa", "ativo"]).id.agg(lambda s: ", ".join(sorted(set(s))))
                     .unstack().reset_index(), hide_index=True, use_container_width=True)
    else:
        st.info("Nenhuma placa de interface compartilhada.")

with t5:
    dup = f[f.duplicado].groupby(["ativo", "nome"]).id.nunique().reset_index(name="ocorrências")
    q = st.columns(3)
    q[0].metric("Células pendentes", int(pen))
    q[1].metric("Pontos fora do padrão", int(f.fora_padrao.sum()))
    q[2].metric("Blocos repetidos", len(dup))
    st.markdown("##### Blocos com o mesmo nome aparecendo mais de uma vez")
    st.dataframe(dup.rename(columns={"ativo": "Ativo", "nome": "Nome"}), hide_index=True)
    cols = ["ativo", "id", "local", "funcao", "codigo", "ponto"]
    st.markdown("##### Pontos fora do padrão de escrita (revisar)")
    st.dataframe(f[f.fora_padrao][cols], hide_index=True)
    st.markdown("##### Pendências (células vazias)")
    st.dataframe(f[f.status == "Pendente"][cols[:-1]], hide_index=True)

with t6:
    st.dataframe(f.drop(columns=["nome", "duplicado"]), hide_index=True)
    st.download_button("⬇️ Baixar dados tratados (CSV)", f.to_csv(index=False).encode("utf-8-sig"),
                       "levantamento_bd_tratado.csv", "text/csv")
