#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# -*- coding: utf-8 -*-
"""
Relatórios Espelhos + Pfizer — Seleção Interativa, Padronização e Análises
----------------------------------------------------------------------------

• Lê a pasta de ESPELHOS (clientes múltiplos; alguns com várias LOJAS) e o arquivo PFIZER (estrutura distinta).
• Interface com checkboxes para selecionar clientes e, quando houver, as LOJAS de cada cliente.
• Junta todos clientes selecionados (exceto Pfizer) em um único DataFrame: df_espelhos.
• Lê Pfizer separadamente em df_pfizer.
• Remove linhas em branco e padroniza nomes de colunas.
• Gera análises com TABELAS (estilizadas para HTML) e GRÁFICOS (PNG) para inserir no corpo de e-mail mais tarde.
• Cada análise é implementada em um bloco (função) separado, salvando HTML e PNG em disco.

Como executar:
1) Instale dependências (se necessário):
   pip install streamlit pandas numpy matplotlib openpyxl xlrd python-dateutil unidecode

2) Rode o app:
   streamlit run relatorios_espelhos_pfizer_streamlit.py

Observações:
- Tabelas são salvas como HTML estilizado (pronto para inserir no corpo do e-mail).
- Gráficos são salvos como PNG.
- Você pode ajustar os caminhos em CONFIG.
- O código tenta detectar automaticamente colunas equivalentes (ex.: nomes diferentes para as mesmas variáveis).
"""

import os
import re
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from unidecode import unidecode
import streamlit as st

# ===============================
# CONFIGURAÇÕES
# ===============================
PASTA_ESPELHOS = r"P:\Relatórios\2020\06. 2025"  # Pasta com vários arquivos de clientes (estrutura igual)
ARQ_PFIZER    = r"P:\Relatórios\2020\Pfizer.xlsx" # Arquivo único com estrutura diferente

# Pasta de saída (HTML + PNG) — use a que você já utiliza nos e-mails
PASTA_SAIDA   = r"C:\Users\fabio\Reunião python"
os.makedirs(PASTA_SAIDA, exist_ok=True)

# Datas de referência
HOJE = datetime.now()
INICIO_ANO = datetime(HOJE.year, 1, 1)
INICIO_MES = datetime(HOJE.year, HOJE.month, 1)

# Paleta/cor para listras (azul/branco)
COR_AZUL = "#e6f0ff"
COR_BRANCO = "#ffffff"

# ===============================
# UTILITÁRIOS GERAIS
# ===============================

def normalize_colname(col: str) -> str:
    """Normaliza nome de coluna: sem acentos, maiúscula, tira espaços duplicados."""
    c = unidecode(str(col)).strip()
    c = re.sub(r"\s+", " ", c)
    return c.upper()


def padronizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_colname(c) for c in df.columns]
    return df


def encontrar_coluna(df: pd.DataFrame, candidatos):
    """Retorna o nome da primeira coluna existente no DF dentre os candidatos (case-insensitive normalizado)."""
    norm_cols = {normalize_colname(c): c for c in df.columns}
    for c in candidatos:
        key = normalize_colname(c)
        if key in norm_cols:
            return norm_cols[key]
    return None


def detectar_coluna_valor(df: pd.DataFrame):
    candidatos = [
        "VALOR", "VALOR TOTAL", "TOTAL", "CUSTO", "R$", "VALOR R$", "TOTAL R$", "VALOR TOTAL R$",
    ]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_data_aprov(df: pd.DataFrame):
    candidatos = [
        "DATA APROVACAO", "DATA APROVAÇÃO", "DT APROVACAO", "APROVACAO", "APROVAÇÃO"
    ]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_data_exec(df: pd.DataFrame):
    candidatos = [
        "DATA EXECUCAO", "DATA EXECUÇÃO", "DT EXECUCAO", "EXECUCAO", "EXECUÇÃO"
    ]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_os(df: pd.DataFrame):
    candidatos = ["OS", "Nº OS", "NUMERO OS", "N OS", "ORDEM SERVICO", "ORDEM DE SERVICO"]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_cliente(df: pd.DataFrame):
    candidatos = ["CLIENTE", "NOME CLIENTE", "RAZAO SOCIAL", "CLIENTE/LOJA"]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_loja(df: pd.DataFrame):
    candidatos = ["LOJA", "FILIAL", "UNIDADE", "PDV", "STORE"]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_equip(df: pd.DataFrame):
    candidatos = ["EQUIPAMENTO", "EQUIP", "MAQUINA", "MÁQUINA", "SERIE EQUIPAMENTO", "EQUIPAMENTO/SETOR"]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_motivo(df: pd.DataFrame):
    candidatos = ["MOTIVO", "MOTIVO CORRETIVA", "MOTIVO MAU USO", "MOTIVO GARANTIA", "TIPO MOTIVO"]
    # Retorna MOTIVO (texto base) — os outros específicos são consultados diretamente pelo nome
    return encontrar_coluna(df, candidatos)


def detectar_coluna_tecnico(df: pd.DataFrame):
    candidatos = ["TECNICO", "TÉCNICO", "RESPONSAVEL TECNICO", "RESPONSAVEL", "EXECUTOR"]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_parada(df: pd.DataFrame):
    candidatos = ["PARADA", "DATA PARADA", "DT PARADA"]
    return encontrar_coluna(df, candidatos)


def detectar_coluna_liberada(df: pd.DataFrame):
    candidatos = ["LIBERADA", "DATA LIBERADA", "DT LIBERADA", "LIBERACAO", "LIBERAÇÃO"]
    return encontrar_coluna(df, candidatos)


def to_datetime_safe(s: pd.Series):
    return pd.to_datetime(s, errors='coerce', dayfirst=True)


def formatar_reais_serie(s: pd.Series) -> pd.Series:
    return s.fillna(0).map(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))


from pandas.io.formats.style import Styler  # ou apenas use pandas.Styler

def estilizar_tabela(df: pd.DataFrame, index=False):

    styler = df.style.hide(axis='index' if not index else None)
    styler = styler.set_properties(**{
        'border': '1px solid #B0B0B0',
        'text-align': 'center',
        'padding': '6px'
    })
    styler = styler.set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#d9e6ff'), ('text-align', 'center'), ('border', '1px solid #808080'), ('padding', '6px')]},
        {'selector': 'table', 'props': [('border-collapse', 'collapse'), ('width', '100%')]},
        {'selector': 'tbody tr:nth-child(odd)', 'props': [('background-color', COR_BRANCO)]},
        {'selector': 'tbody tr:nth-child(even)', 'props': [('background-color', COR_AZUL)]},
    ])
    return styler


def salvar_tabela_html(df: pd.DataFrame, caminho_html: str, index=False):
    styler = estilizar_tabela(df, index=index)
    html = styler.to_html()
    with open(caminho_html, 'w', encoding='utf-8') as f:
        f.write(html)


def salvar_fig(fig, caminho_png: str):
    fig.tight_layout()
    fig.savefig(caminho_png, dpi=160, bbox_inches='tight')
    plt.close(fig)

# ===============================
# LEITURA DE ARQUIVOS
# ===============================

def listar_arquivos_espelhos(pasta: str):
    exts = (".xlsx", ".xls", ".xlsm")
    arquivos = [str(p) for p in Path(pasta).glob("**/*") if p.suffix.lower() in exts]
    return sorted(arquivos)


def carregar_espelho(path_arq: str) -> pd.DataFrame:
    """Lê um espelho padrão e padroniza colunas chave."""
    df = pd.read_excel(path_arq)
    df = padronizar_colunas(df)
    # Rename básicos opcionais podem ser adicionados aqui, se precisar
    return df


def carregar_pfizer(path_arq: str) -> pd.DataFrame:
    """Lê o Pfizer (estrutura diferente) e mapeia para colunas comuns quando possível."""
    df = pd.read_excel(path_arq)
    df = padronizar_colunas(df)
    # Se necessário, mapear colunas do Pfizer para chaves comuns
    # Exemplo de mapeamento (ajuste conforme o real):
    possiveis_mapas = {
        # 'COLUNA_PFIZER_ORIGINAL': 'EQUIVALENTE_COMUM'
        # Ex.: 'UNIDADE' -> 'LOJA'
    }
    for k, v in possiveis_mapas.items():
        if k in df.columns and v not in df.columns:
            df[v] = df[k]
    return df


def remover_linhas_em_branco(df: pd.DataFrame) -> pd.DataFrame:
    # Remove linhas totalmente vazias e também onde todas colunas chave estão vazias
    df2 = df.dropna(how='all')
    # Se existir uma coluna OS ou LOJA, removemos linhas sem essas chaves
    col_os = detectar_coluna_os(df2)
    col_loja = detectar_coluna_loja(df2)
    if col_os:
        df2 = df2[~df2[col_os].isna()]
    if col_loja:
        df2 = df2[~df2[col_loja].astype(str).str.strip().eq("")]
    return df2

# ===============================
# INTERFACE — SELEÇÃO
# ===============================

st.set_page_config(page_title="Relatórios Espelhos & Pfizer", layout="wide")
st.title("Relatórios — Espelhos (df_espelhos) e Pfizer (df_pfizer)")

with st.sidebar:
    st.header("Configurações")
    pasta_espelhos = st.text_input("Pasta dos Espelhos", value=PASTA_ESPELHOS)
    arq_pfizer = st.text_input("Arquivo Pfizer", value=ARQ_PFIZER)
    pasta_saida = st.text_input("Pasta de saída (HTML/PNG)", value=PASTA_SAIDA)
    if st.button("Recarregar lista de clientes"):
        st.session_state.pop('cache_lista', None)

# Listar arquivos de clientes
if 'cache_lista' not in st.session_state:
    arquivos = listar_arquivos_espelhos(pasta_espelhos)
    # Inferir nome do cliente pelo nome do arquivo/pasta
    clientes_map = {}
    for a in arquivos:
        nome = Path(a).stem
        cliente = re.sub(r"(?i)(espelho|relatorio|report|base|dados|_+|\s+)", " ", nome).strip()
        clientes_map.setdefault(cliente, []).append(a)
    st.session_state['cache_lista'] = clientes_map
else:
    clientes_map = st.session_state['cache_lista']

st.subheader("Seleção de Clientes")
clientes = sorted(clientes_map.keys())
sel_clientes = st.multiselect("Escolha os clientes para analisar (ESPelhos)", options=clientes)

# Para cada cliente escolhido, abrir e descobrir lojas para permitir selecionar
selecoes_lojas_por_cliente = {}

cols_sel = st.columns(min(3, max(1, len(sel_clientes))))
for i, cliente in enumerate(sel_clientes):
    arqs = clientes_map.get(cliente, [])
    # Carrega rapidamente para identificar lojas únicas
    dflist = []
    for p in arqs:
        try:
            dftmp = carregar_espelho(p)
            col_loja = detectar_coluna_loja(dftmp)
            if col_loja:
                dflist.append(dftmp[[col_loja]].dropna())
        except Exception:
            pass
    if dflist:
        dfl = pd.concat(dflist, ignore_index=True)
        col_loja = detectar_coluna_loja(dfl)
        lojas = sorted(dfl[col_loja].astype(str).str.strip().unique()) if col_loja else []
    else:
        lojas = []

    with cols_sel[i % len(cols_sel)]:
        st.markdown(f"**{cliente}**")
        if len(lojas) > 1:
            sel_lojas = st.multiselect(f"Lojas de {cliente}", options=lojas, key=f"{cliente}_lojas")
            selecoes_lojas_por_cliente[cliente] = sel_lojas
        else:
            selecoes_lojas_por_cliente[cliente] = lojas  # 0 ou 1: não precisa checkbox

# Opção para incluir Pfizer
st.subheader("Arquivo Pfizer")
incluir_pfizer = st.checkbox("Importar Pfizer (df_pfizer separado)", value=True if os.path.exists(ARQ_PFIZER) else False)

# ===============================
# IMPORTAÇÃO EFETIVA
# ===============================

if st.button("Importar e Unificar"):
    st.session_state['df_espelhos'] = None
    st.session_state['df_pfizer'] = None

    # 1) Unificar ESPELHOS
    dfs = []
    for cliente in sel_clientes:
        arqs = clientes_map.get(cliente, [])
        lojas_escolhidas = selecoes_lojas_por_cliente.get(cliente, [])
        for p in arqs:
            try:
                dfc = carregar_espelho(p)
                col_loja = detectar_coluna_loja(dfc)
                if col_loja and lojas_escolhidas:
                    dfc = dfc[dfc[col_loja].astype(str).str.strip().isin([str(x).strip() for x in lojas_escolhidas])]
                # Tag do cliente, se não houver coluna cliente
                col_cli = detectar_coluna_cliente(dfc)
                if not col_cli:
                    dfc["CLIENTE"] = cliente
                st.write(f"Arquivo carregado: {Path(p).name} — {len(dfc)} linhas")
                dfs.append(dfc)
            except Exception as e:
                st.warning(f"Falha ao carregar {p}: {e}")
    if dfs:
        df_espelhos = pd.concat(dfs, ignore_index=True)
        df_espelhos = remover_linhas_em_branco(df_espelhos)
        # Tipagens & datas
        col_aprov = detectar_coluna_data_aprov(df_espelhos)
        col_exec = detectar_coluna_data_exec(df_espelhos)
        if col_aprov: df_espelhos[col_aprov] = to_datetime_safe(df_espelhos[col_aprov])
        if col_exec:  df_espelhos[col_exec]  = to_datetime_safe(df_espelhos[col_exec])
        # Valor
        col_valor = detectar_coluna_valor(df_espelhos)
        if col_valor:
            df_espelhos[col_valor] = pd.to_numeric(df_espelhos[col_valor], errors='coerce').fillna(0.0)
        st.session_state['df_espelhos'] = df_espelhos
        st.success(f"df_espelhos criado com {len(df_espelhos)} linhas")
    else:
        st.info("Nenhum ESPelho selecionado ou carregado.")

    # 2) PFIZER
    if incluir_pfizer and os.path.exists(arq_pfizer):
        try:
            df_pfizer = carregar_pfizer(arq_pfizer)
            df_pfizer = remover_linhas_em_branco(df_pfizer)
            # Datas e valor, se existirem
            col_aprov_p = detectar_coluna_data_aprov(df_pfizer)
            col_exec_p  = detectar_coluna_data_exec(df_pfizer)
            if col_aprov_p: df_pfizer[col_aprov_p] = to_datetime_safe(df_pfizer[col_aprov_p])
            if col_exec_p:  df_pfizer[col_exec_p]  = to_datetime_safe(df_pfizer[col_exec_p])
            col_valor_p  = detectar_coluna_valor(df_pfizer)
            if col_valor_p:
                df_pfizer[col_valor_p] = pd.to_numeric(df_pfizer[col_valor_p], errors='coerce').fillna(0.0)
            st.session_state['df_pfizer'] = df_pfizer
            st.success(f"df_pfizer criado com {len(df_pfizer)} linhas")
        except Exception as e:
            st.warning(f"Falha ao carregar Pfizer: {e}")

# ===============================
# BLOCO DE ANÁLISES — FUNÇÕES
# ===============================

def bloco_1_gastos_mes(df: pd.DataFrame, nome_prefixo: str):
    """Gastos no mês atual por CLIENTE/LOJA. Tabela + Gráfico de barras."""
    if df is None or df.empty:
        return
    df = df.copy()
    col_valor = detectar_coluna_valor(df)
    col_cli   = detectar_coluna_cliente(df) or 'CLIENTE'
    col_loja  = detectar_coluna_loja(df)
    col_aprov = detectar_coluna_data_aprov(df) or detectar_coluna_data_exec(df)
    if not col_valor or not col_aprov:
        st.info("Colunas de VALOR/Data não encontradas para 'Gastos no mês'.")
        return
    df_mes = df[(df[col_aprov] >= INICIO_MES) & (df[col_aprov] < (INICIO_MES + relativedelta(months=1)))]
    if df_mes.empty:
        st.info("Sem dados no mês atual.")
        return
    grp_cols = [c for c in [col_cli, col_loja] if c]
    tab = df_mes.groupby(grp_cols, dropna=False)[col_valor].sum().reset_index()
    tab = tab.sort_values(col_valor, ascending=False)
    tab['VALOR (R$)'] = formatar_reais_serie(tab[col_valor])
    tab_fmt = tab[[c for c in grp_cols] + ['VALOR (R$)']]

    # Salvar tabela e gráfico
    nome_base = f"{nome_prefixo}_01_gastos_mes"
    caminho_html = os.path.join(pasta_saida, f"{nome_base}.html")
    salvar_tabela_html(tab_fmt, caminho_html)

    fig, ax = plt.subplots(figsize=(10, 5))
    xlabels = tab_fmt.apply(lambda r: " - ".join([str(r[c]) for c in grp_cols if c in r.index and pd.notna(r[c])]), axis=1)
    ax.bar(xlabels, tab[col_valor].values)
    ax.set_title("Gastos no mês atual")
    ax.set_ylabel("R$")
    ax.set_xticklabels(xlabels, rotation=45, ha='right')
    caminho_png = os.path.join(pasta_saida, f"{nome_base}.png")
    salvar_fig(fig, caminho_png)

    st.subheader("Gastos no mês atual")
    st.components.v1.html(Path(caminho_html).read_text(encoding='utf-8'), height=350, scrolling=True)
    st.image(caminho_png, use_column_width=True)


def bloco_2_gastos_ano(df: pd.DataFrame, nome_prefixo: str):
    """Gastos no ano atual por CLIENTE/LOJA."""
    if df is None or df.empty:
        return
    df = df.copy()
    col_valor = detectar_coluna_valor(df)
    col_cli   = detectar_coluna_cliente(df) or 'CLIENTE'
    col_loja  = detectar_coluna_loja(df)
    col_aprov = detectar_coluna_data_aprov(df) or detectar_coluna_data_exec(df)
    if not col_valor or not col_aprov:
        st.info("Colunas de VALOR/Data não encontradas para 'Gastos no ano'.")
        return
    df_ano = df[(df[col_aprov] >= INICIO_ANO) & (df[col_aprov] <= HOJE)]
    grp_cols = [c for c in [col_cli, col_loja] if c]
    tab = df_ano.groupby(grp_cols, dropna=False)[col_valor].sum().reset_index()
    tab = tab.sort_values(col_valor, ascending=False)
    tab['VALOR (R$)'] = formatar_reais_serie(tab[col_valor])
    tab_fmt = tab[[c for c in grp_cols] + ['VALOR (R$)']]

    nome_base = f"{nome_prefixo}_02_gastos_ano"
    caminho_html = os.path.join(pasta_saida, f"{nome_base}.html")
    salvar_tabela_html(tab_fmt, caminho_html)

    fig, ax = plt.subplots(figsize=(10, 5))
    xlabels = tab_fmt.apply(lambda r: " - ".join([str(r[c]) for c in grp_cols if c in r.index and pd.notna(r[c])]), axis=1)
    ax.bar(xlabels, tab[col_valor].values)
    ax.set_title("Gastos no ano atual")
    ax.set_ylabel("R$")
    ax.set_xticklabels(xlabels, rotation=45, ha='right')
    caminho_png = os.path.join(pasta_saida, f"{nome_base}.png")
    salvar_fig(fig, caminho_png)

    st.subheader("Gastos no ano atual")
    st.components.v1.html(Path(caminho_html).read_text(encoding='utf-8'), height=350, scrolling=True)
    st.image(caminho_png, use_column_width=True)


def bloco_3_motivos(df: pd.DataFrame, nome_prefixo: str):
    """Análise de motivos (MOTIVO / CORRETIVA / MAU USO / GARANTIA) — mês atual."""
    if df is None or df.empty:
        return
    df = df.copy()
    col_motivo = detectar_coluna_motivo(df)
    col_valor  = detectar_coluna_valor(df)
    col_aprov  = detectar_coluna_data_aprov(df) or detectar_coluna_data_exec(df)
    if not col_aprov:
        st.info("Coluna de data não encontrada para 'Motivos (mês)'.")
        return

    df_mes = df[(df[col_aprov] >= INICIO_MES) & (df[col_aprov] < (INICIO_MES + relativedelta(months=1)))]

    # Extrair motivo final quando padrão "M.O.V.C - ... / SUBMOTIVO"
    def extrair_motivo_final(row):
        texto = str(row.get(col_motivo, "")) if col_motivo else ""
        texto_norm = unidecode(texto).lower()
        # tenta pegar parte apos ' - '
        if "-" in texto:
            parte = texto.split("-", 1)[1].strip()
        else:
            parte = texto
        # Se houver '/', pega submotivo (depois de '/')
        if "/" in parte:
            sub = parte.split("/")[-1].strip()
        else:
            sub = parte
        # Valida em colunas específicas se existirem
        for espec in ["MOTIVO CORRETIVA", "MOTIVO MAU USO", "MOTIVO GARANTIA"]:
            if espec in df.columns and pd.notna(row.get(espec)) and str(row.get(espec)).strip():
                return f"{espec}: {str(row.get(espec)).strip()}"
        return sub if sub else (texto if texto else "N/D")

    df_mes["MOTIVO_FINAL"] = df_mes.apply(extrair_motivo_final, axis=1)

    # Contagem e somatório
    if col_valor and col_valor in df_mes.columns:
        tab = df_mes.groupby("MOTIVO_FINAL").agg(
            QUANTIDADE=("MOTIVO_FINAL", "size"),
            VALOR=(col_valor, "sum")
        ).reset_index().sort_values(["VALOR", "QUANTIDADE"], ascending=False)
        tab["VALOR (R$)"] = formatar_reais_serie(tab["VALOR"])
        tab_fmt = tab[["MOTIVO_FINAL", "QUANTIDADE", "VALOR (R$)"]]
    else:
        tab = df_mes.groupby("MOTIVO_FINAL").size().reset_index(name="QUANTIDADE").sort_values("QUANTIDADE", ascending=False)
        tab_fmt = tab.copy()

    nome_base = f"{nome_prefixo}_03_motivos_mes"
    caminho_html = os.path.join(pasta_saida, f"{nome_base}.html")
    salvar_tabela_html(tab_fmt, caminho_html)

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.barh(tab_fmt["MOTIVO_FINAL"].astype(str), tab["QUANTIDADE"].values)
    ax.set_title("Motivos — mês atual")
    ax.set_xlabel("Quantidade")
    caminho_png = os.path.join(pasta_saida, f"{nome_base}.png")
    salvar_fig(fig, caminho_png)

    st.subheader("Motivos — mês atual")
    st.components.v1.html(Path(caminho_html).read_text(encoding='utf-8'), height=380, scrolling=True)
    st.image(caminho_png, use_column_width=True)


def bloco_4_os_aprovadas_exec(df: pd.DataFrame, nome_prefixo: str):
    """OS aprovadas, aprovadas+executadas no mês; tempo médio aprovação→execução (mês e meses passados do ano)."""
    if df is None or df.empty:
        return
    df = df.copy()
    col_aprov = detectar_coluna_data_aprov(df)
    col_exec  = detectar_coluna_data_exec(df)
    col_os    = detectar_coluna_os(df)
    if not col_aprov or not col_os:
        st.info("Colunas de aprovação/OS não encontradas para 'OS e tempos'.")
        return

    df[col_aprov] = to_datetime_safe(df[col_aprov])
    if col_exec:
        df[col_exec] = to_datetime_safe(df[col_exec])

    # Mês atual
    dm = df[(df[col_aprov] >= INICIO_MES) & (df[col_aprov] < (INICIO_MES + relativedelta(months=1)))]
    os_aprov = dm[col_os].nunique()
    if col_exec:
        dm_exec = dm[dm[col_exec].notna()]
        os_aprov_exec = dm_exec[col_os].nunique()
        tempo_medio_mes = (dm_exec[col_exec] - dm_exec[col_aprov]).dt.days.dropna().mean()
    else:
        os_aprov_exec = np.nan
        tempo_medio_mes = np.nan

    # Meses passados do ano (até mês anterior)
    ate_mes_anterior = INICIO_MES - timedelta(days=1)
    da = df[(df[col_aprov] >= INICIO_ANO) & (df[col_aprov] <= ate_mes_anterior)]
    if col_exec:
        da_exec = da[da[col_exec].notna()]
        tempo_medio_passado = (da_exec[col_exec] - da_exec[col_aprov]).dt.days.dropna().mean()
    else:
        tempo_medio_passado = np.nan

    tab = pd.DataFrame({
        "Indicador": ["OS aprovadas (mês)", "OS aprovadas & executadas (mês)", "Tempo médio (dias) mês", "Tempo médio (dias) meses passados"],
        "Valor": [os_aprov, os_aprov_exec, round(tempo_medio_mes, 1) if pd.notna(tempo_medio_mes) else None, round(tempo_medio_passado, 1) if pd.notna(tempo_medio_passado) else None]
    })

    nome_base = f"{nome_prefixo}_04_os_tempos"
    caminho_html = os.path.join(pasta_saida, f"{nome_base}.html")
    salvar_tabela_html(tab, caminho_html)

    fig, ax = plt.subplots(figsize=(8, 4))
    ax.bar(tab["Indicador"], tab["Valor"].astype(float))
    ax.set_title("OS e Tempos")
    ax.set_ylabel("Valor")
    ax.set_xticklabels(tab["Indicador"], rotation=20, ha='right')
    caminho_png = os.path.join(pasta_saida, f"{nome_base}.png")
    salvar_fig(fig, caminho_png)

    st.subheader("OS aprovadas & tempos")
    st.components.v1.html(Path(caminho_html).read_text(encoding='utf-8'), height=220, scrolling=False)
    st.image(caminho_png, use_column_width=True)


def bloco_5_lista_equipamentos(df: pd.DataFrame, nome_prefixo: str):
    """Lista por cliente/loja: equipamentos atendidos no mês, não atendidos no mês, média de dias entre atendimentos."""
    if df is None or df.empty:
        return
    df = df.copy()
    col_cli   = detectar_coluna_cliente(df) or 'CLIENTE'
    col_loja  = detectar_coluna_loja(df)
    col_eq    = detectar_coluna_equip(df)
    col_aprov = detectar_coluna_data_aprov(df) or detectar_coluna_data_exec(df)
    if not col_eq or not col_aprov:
        st.info("Colunas de EQUIPAMENTO/Data não encontradas para 'Lista de equipamentos'.")
        return

    df[col_aprov] = to_datetime_safe(df[col_aprov])

    # Universo por cliente/loja
    grp = [c for c in [col_cli, col_loja, col_eq] if c]

    # Atendidos no mês
    dm = df[(df[col_aprov] >= INICIO_MES) & (df[col_aprov] < (INICIO_MES + relativedelta(months=1)))]
    atendidos_mes = dm.groupby(grp).size().reset_index(name='ATENDIMENTOS_MES')

    # Últimas datas por equipamento
    ult_datas = df.groupby(grp)[col_aprov].max().reset_index(name='ULT_ATEND')

    # Média de dias entre atendimentos por equipamento
    df_sorted = df.sort_values(grp + [col_aprov])
    df_sorted['_LAG'] = df_sorted.groupby(grp)[col_aprov].shift(1)
    df_sorted['DIAS_ENTRE'] = (df_sorted[col_aprov] - df_sorted['_LAG']).dt.days
    medias = df_sorted.groupby(grp)['DIAS_ENTRE'].mean().reset_index(name='MEDIA_DIAS_ENTRE')

    tab = ult_datas.merge(medias, on=grp, how='left').merge(atendidos_mes, on=grp, how='left')
    tab['ATENDIMENTOS_MES'] = tab['ATENDIMENTOS_MES'].fillna(0).astype(int)

    # Não atendidos no mês = ATENDIMENTOS_MES == 0
    tab['NAO_ATENDIDOS_MES'] = (tab['ATENDIMENTOS_MES'] == 0)

    # Equipamentos > 30 dias sem atendimento
    tab['DIAS_SEM_ATEND'] = (HOJE - tab['ULT_ATEND']).dt.days
    tab['>30_DIAS_SEM_ATEND'] = tab['DIAS_SEM_ATEND'] > 30

    # Formatar
    tab_fmt = tab.copy()
    tab_fmt['ULT_ATEND'] = tab_fmt['ULT_ATEND'].dt.strftime('%d/%m/%Y')
    tab_fmt['MEDIA_DIAS_ENTRE'] = tab_fmt['MEDIA_DIAS_ENTRE'].round(1)

    nome_base = f"{nome_prefixo}_05_equipamentos"
    caminho_html = os.path.join(pasta_saida, f"{nome_base}.html")
    salvar_tabela_html(tab_fmt, caminho_html)

    # Gráfico: barras de atendidos no mês por equipamento (top 20)
    top = tab.sort_values('ATENDIMENTOS_MES', ascending=False).head(20)
    labels = top[col_eq].astype(str)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(labels, top['ATENDIMENTOS_MES'])
    ax.invert_yaxis()
    ax.set_title("Top equipamentos por atendimentos no mês")
    ax.set_xlabel("Atendimentos")
    caminho_png = os.path.join(pasta_saida, f"{nome_base}.png")
    salvar_fig(fig, caminho_png)

    st.subheader("Equipamentos — atendimentos & intervalos")
    st.components.v1.html(Path(caminho_html).read_text(encoding='utf-8'), height=420, scrolling=True)
    st.image(caminho_png, use_column_width=True)


def bloco_6_paradas_motivos(df: pd.DataFrame, nome_prefixo: str):
    """Máquinas paradas: dias parados e motivo (checando colunas PARADA/LIBERADA e colunas de motivo específicas)."""
    if df is None or df.empty:
        return
    df = df.copy()
    col_eq   = detectar_coluna_equip(df)
    col_cli  = detectar_coluna_cliente(df) or 'CLIENTE'
    col_loja = detectar_coluna_loja(df)
    col_par  = detectar_coluna_parada(df)
    col_lib  = detectar_coluna_liberada(df)
    if not col_eq or not col_par:
        st.info("Colunas de EQUIPAMENTO/PARADA não encontradas para 'Máquinas paradas'.")
        return

    df[col_par] = to_datetime_safe(df[col_par])
    if col_lib:
        df[col_lib] = to_datetime_safe(df[col_lib])

    # Considera parada aberta quando LIBERADA é nula ou LIBERADA < PARADA?
    base = df[df[col_par].notna()].copy()
    base['DIAS_PARADA'] = np.where(base.get(col_lib, pd.Series(index=base.index)).notna(),
                                   (base[col_lib] - base[col_par]).dt.days,
                                   (HOJE - base[col_par]).dt.days)

    def pick_motivo(row):
        for espec in ["MOTIVO CORRETIVA", "MOTIVO MAU USO", "MOTIVO GARANTIA"]:
            if espec in df.columns and pd.notna(row.get(espec)) and str(row.get(espec)).strip():
                return f"{espec}: {str(row.get(espec)).strip()}"
        mot = detectar_coluna_motivo(df)
        return str(row.get(mot)) if mot and pd.notna(row.get(mot)) else "N/D"

    base['MOTIVO_PARADA'] = base.apply(pick_motivo, axis=1)

    grp = [c for c in [col_cli, col_loja, col_eq] if c]
    tab = base.groupby(grp, dropna=False).agg(
        OCORRENCIAS=(col_eq, 'size'),
        DIAS_PARADA_MEDIO=('DIAS_PARADA', 'mean'),
        DIAS_PARADA_MAX=('DIAS_PARADA', 'max'),
        MOTIVO_PREDOMINANTE=('MOTIVO_PARADA', lambda x: x.value_counts().idxmax() if len(x) else 'N/D')
    ).reset_index()
    tab['DIAS_PARADA_MEDIO'] = tab['DIAS_PARADA_MEDIO'].round(1)

    nome_base = f"{nome_prefixo}_06_paradas"
    caminho_html = os.path.join(pasta_saida, f"{nome_base}.html")
    salvar_tabela_html(tab, caminho_html)

    # Gráfico: top 20 por dias parada máximo
    top = tab.sort_values('DIAS_PARADA_MAX', ascending=False).head(20)
    label_eq = top[col_eq].astype(str)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(label_eq, top['DIAS_PARADA_MAX'])
    ax.invert_yaxis()
    ax.set_title("Top máquinas por maior tempo parado")
    ax.set_xlabel("Dias parados (máximo)")
    caminho_png = os.path.join(pasta_saida, f"{nome_base}.png")
    salvar_fig(fig, caminho_png)

    st.subheader("Máquinas paradas — dias e motivos")
    st.components.v1.html(Path(caminho_html).read_text(encoding='utf-8'), height=420, scrolling=True)
    st.image(caminho_png, use_column_width=True)


def bloco_7_gastos_por_equipamento(df: pd.DataFrame, nome_prefixo: str):
    """Quanto cada equipamento gastou no mês (com motivo detalhado via parsing de 'M.O.V.C - ... / ...')."""
    if df is None or df.empty:
        return
    df = df.copy()
    col_eq    = detectar_coluna_equip(df)
    col_valor = detectar_coluna_valor(df)
    col_aprov = detectar_coluna_data_aprov(df) or detectar_coluna_data_exec(df)
    mot_col   = detectar_coluna_motivo(df)
    if not col_eq or not col_valor or not col_aprov:
        st.info("Colunas de EQUIPAMENTO/VALOR/Data não encontradas para 'Gastos por equipamento'.")
        return

    df[col_aprov] = to_datetime_safe(df[col_aprov])
    df_mes = df[(df[col_aprov] >= INICIO_MES) & (df[col_aprov] < (INICIO_MES + relativedelta(months=1)))]

    def motivo_detalhe(row):
        texto = str(row.get(mot_col, "")) if mot_col else ""
        if "-" in texto:
            parte = texto.split("-", 1)[1].strip()
        else:
            parte = texto
        if "/" in parte:
            sub = parte.split("/")[-1].strip()
        else:
            sub = parte
        # Verifica colunas específicas
        for espec in ["MOTIVO CORRETIVA", "MOTIVO MAU USO", "MOTIVO GARANTIA"]:
            if espec in df.columns and pd.notna(row.get(espec)) and str(row.get(espec)).strip():
                return f"{espec}: {str(row.get(espec)).strip()}"
        return sub if sub else (texto if texto else "N/D")

    df_mes['MOTIVO_DETALHE'] = df_mes.apply(motivo_detalhe, axis=1)

    tab = df_mes.groupby([col_eq, 'MOTIVO_DETALHE'])[col_valor].sum().reset_index()
    tab = tab.sort_values(col_valor, ascending=False)
    tab['VALOR (R$)'] = formatar_reais_serie(tab[col_valor])
    tab_fmt = tab[[col_eq, 'MOTIVO_DETALHE', 'VALOR (R$)']]

    nome_base = f"{nome_prefixo}_07_gastos_por_equip_mes"
    caminho_html = os.path.join(pasta_saida, f"{nome_base}.html")
    salvar_tabela_html(tab_fmt, caminho_html)

    # Gráfico: barras por equipamento (top 20)
    top = tab.groupby(col_eq)[col_valor].sum().reset_index().sort_values(col_valor, ascending=False).head(20)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(top[col_eq].astype(str), top[col_valor])
    ax.invert_yaxis()
    ax.set_title("Gastos por equipamento — mês atual")
    ax.set_xlabel("R$")
    caminho_png = os.path.join(pasta_saida, f"{nome_prefixo}_07_gastos_por_equip_mes.png")
    salvar_fig(fig, caminho_png)

    st.subheader("Gastos por equipamento — mês atual (com motivo)")
    st.components.v1.html(Path(caminho_html).read_text(encoding='utf-8'), height=420, scrolling=True)
    st.image(caminho_png, use_column_width=True)


def bloco_8_tecnicos_por_loja(df: pd.DataFrame, nome_prefixo: str):
    """Técnicos que mais atendem por LOJA (mês atual)."""
    if df is None or df.empty:
        return
    df = df.copy()
    col_loja = detectar_coluna_loja(df)
    col_tec  = detectar_coluna_tecnico(df)
    col_aprov = detectar_coluna_data_aprov(df) or detectar_coluna_data_exec(df)
    if not col_loja or not col_tec or not col_aprov:
        st.info("Colunas de LOJA/TÉCNICO/Data não encontradas para 'Técnicos por loja'.")
        return

    df[col_aprov] = to_datetime_safe(df[col_aprov])
    dm = df[(df[col_aprov] >= INICIO_MES) & (df[col_aprov] < (INICIO_MES + relativedelta(months=1)))]

    tab = dm.groupby([col_loja, col_tec]).size().reset_index(name='ATENDIMENTOS')
    tab = tab.sort_values(['ATENDIMENTOS'], ascending=False)

    nome_base = f"{nome_prefixo}_08_tecnicos_por_loja"
    caminho_html = os.path.join(pasta_saida, f"{nome_base}.html")
    salvar_tabela_html(tab, caminho_html)

    # Gráfico: top técnicos por total atendimentos (todas lojas)
    top = tab.groupby(col_tec)['ATENDIMENTOS'].sum().reset_index().sort_values('ATENDIMENTOS', ascending=False).head(20)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(top[col_tec].astype(str), top['ATENDIMENTOS'])
    ax.invert_yaxis()
    ax.set_title("Top técnicos por atendimentos (mês)")
    ax.set_xlabel("Atendimentos")
    caminho_png = os.path.join(pasta_saida, f"{nome_prefixo}_08_tecnicos.png")
    salvar_fig(fig, caminho_png)

    st.subheader("Técnicos por loja — mês atual")
    st.components.v1.html(Path(caminho_html).read_text(encoding='utf-8'), height=420, scrolling=True)
    st.image(caminho_png, use_column_width=True)

# ===============================
# EXECUÇÃO DOS BLOCOS DE ANÁLISE
# ===============================

pasta_saida = PASTA_SAIDA  # usa o valor padrão; é atualizado pelos inputs quando importar

if 'df_espelhos' in st.session_state and st.session_state['df_espelhos'] is not None:
    st.markdown("---")
    st.header("Análises — df_espelhos (clientes agrupados)")
    nome_prefixo = "espelhos"
    df_base = st.session_state['df_espelhos']

    bloco_1_gastos_mes(df_base, nome_prefixo)
    bloco_2_gastos_ano(df_base, nome_prefixo)
    bloco_3_motivos(df_base, nome_prefixo)
    bloco_4_os_aprovadas_exec(df_base, nome_prefixo)
    bloco_5_lista_equipamentos(df_base, nome_prefixo)
    bloco_6_paradas_motivos(df_base, nome_prefixo)
    bloco_7_gastos_por_equipamento(df_base, nome_prefixo)
    bloco_8_tecnicos_por_loja(df_base, nome_prefixo)

if 'df_pfizer' in st.session_state and st.session_state['df_pfizer'] is not None:
    st.markdown("---")
    st.header("Análises — df_pfizer (separado)")
    nome_prefixo = "pfizer"
    df_base = st.session_state['df_pfizer']

    # Reaproveita alguns blocos que fazem sentido para Pfizer
    bloco_1_gastos_mes(df_base, nome_prefixo)
    bloco_2_gastos_ano(df_base, nome_prefixo)
    bloco_3_motivos(df_base, nome_prefixo)
    bloco_4_os_aprovadas_exec(df_base, nome_prefixo)
    bloco_5_lista_equipamentos(df_base, nome_prefixo)
    bloco_6_paradas_motivos(df_base, nome_prefixo)
    bloco_7_gastos_por_equipamento(df_base, nome_prefixo)
    bloco_8_tecnicos_por_loja(df_base, nome_prefixo)

st.markdown("---")
st.caption("Todas as tabelas foram salvas em HTML e os gráficos em PNG na pasta de saída configurada. Esses arquivos já estão prontos para inserção no corpo do e-mail.")

