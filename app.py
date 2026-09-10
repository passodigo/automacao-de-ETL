"""
================================================================================
 ETL de PDF -> Excel com interface visual (Streamlit)
================================================================================
Dois modos, na barra lateral:

  1) ETL Genérico — escolhe página, tabela, linha de cabeçalho, renomeia
     colunas livremente, adiciona ao conjunto final.

  2) Comitês x Programas (despivotar) — seção dedicada para tabelas em
     formato "matriz" (Comitê nas linhas, Programa nas colunas). Gera
     formato linear: Comitê | Programa | Ano | Valor.

PONTOS IMPORTANTES DESTA VERSÃO
--------------------------------
- Muitas tabelas de PDF têm CABEÇALHO EM MAIS DE UMA LINHA (ex: uma linha
  com "Programa de Investimento" mesclada em cima, e a linha de baixo com
  "Comitê", "P1", "P2"...). O app agora mostra a tabela crua (sem assumir
  qual linha é o cabeçalho) e deixa você escolher qual linha é o cabeçalho
  real. Isso evita o cabeçalho errado virar "dado" (e gerar lixo).
- No modo Comitê x Programa, dá pra excluir linhas específicas (ex: "Total
  por Programa") antes de despivotar, além de excluir tabelas inteiras na
  etapa de seleção.
- O campo Ano fica vinculado ao lote inteiro: editar depois de adicionado
  recalcula automaticamente todas as linhas daquele lote.

COMO RODAR
----------
    pip install streamlit pdfplumber pandas openpyxl Pillow
    streamlit run app.py
================================================================================
"""

import io
import re
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="PDF -> Excel", layout="wide")

# ==============================================================================
# ESTADO DA SESSÃO
# ==============================================================================
if "consolidado" not in st.session_state:
    st.session_state.consolidado = []

if "coluna_aliases" not in st.session_state:
    st.session_state.coluna_aliases = {}


def normalize(s: str) -> str:
    return " ".join(str(s).strip().lower().split())


def parse_valor_brl(v):
    """Converte '62.805.187' ou '218.426,50' -> float. '-' ou vazio -> None."""
    if v is None:
        return None
    s = str(v).strip()
    if s == "" or s in ("-", "—", "–", "nan", "None"):
        return None
    s = re.sub(r"[^\d,.\-]", "", s)
    if s in ("", "-"):
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        partes = s.split(".")
        if len(partes) > 1 and len(partes[-1]) == 3:
            s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def extract_programa_label(col_name: str) -> str:
    """Mantém o nome da coluna como está (ex: 'P1' continua 'P1')."""
    return str(col_name).strip()


def build_header_and_df(raw_rows: list, header_row_idx: int) -> pd.DataFrame:
    """Constrói o DataFrame a partir das linhas cruas, usando a linha
    header_row_idx como cabeçalho e descartando tudo antes/dela."""
    header = raw_rows[header_row_idx]
    header = [str(h).strip() if h else f"coluna_{i}" for i, h in enumerate(header)]
    # garante nomes únicos (colunas mescladas podem repetir texto)
    seen = {}
    header_unicos = []
    for h in header:
        if h in seen:
            seen[h] += 1
            header_unicos.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            header_unicos.append(h)

    data_rows = raw_rows[header_row_idx + 1:]
    df = pd.DataFrame(data_rows, columns=header_unicos).dropna(how="all")
    return df


def get_df(item: dict) -> pd.DataFrame:
    """Retorna o DataFrame final de um item do conjunto, recalculando na hora
    para o caso de 'comite_programa' (assim o Ano editado sempre reflete)."""
    if item["tipo"] == "generico":
        return item["df"]

    raw_df = item["raw_df"]
    comite_col = item["comite_col"]
    programa_cols = item["programa_cols"]
    ano = item["ano"]

    linhas = []
    for _, row in raw_df.iterrows():
        comite = row.get(comite_col)
        if comite is None or str(comite).strip() == "" or str(comite).strip() == "-":
            continue
        for col in programa_cols:
            valor = parse_valor_brl(row.get(col))
            if valor is None:
                continue
            linhas.append({
                "Comitê": str(comite).strip(),
                "Programa": extract_programa_label(col),
                "Ano": ano,
                "Valor": valor,
            })

    df = pd.DataFrame(linhas, columns=["Comitê", "Programa", "Ano", "Valor"])
    df["_arquivo_origem"] = item.get("arquivo_origem", "")
    df["_pagina_origem"] = item.get("pagina_origem", "")
    return df


# ==============================================================================
# SIDEBAR
# ==============================================================================
st.sidebar.header("Modo de extração")
modo = st.sidebar.radio(
    "Escolha o tipo de tabela que você vai extrair:",
    ["ETL Genérico (qualquer tabela)", "Comitês x Programas (despivotar)"],
)

st.sidebar.header("1. Upload do PDF")
uploaded_file = st.sidebar.file_uploader("Selecione um PDF", type=["pdf"])

st.title("📄 ETL de PDF para Excel")

if uploaded_file is None:
    st.info("Envie um PDF na barra lateral para começar.")
    if st.session_state.consolidado:
        st.subheader("Tabelas já adicionadas nesta sessão")
        for i, item in enumerate(st.session_state.consolidado):
            df_preview = get_df(item)
            st.write(f"**{i+1}. {item['origem']}** — {df_preview.shape[0]} linhas x {df_preview.shape[1]} colunas")
    st.stop()

pdf_bytes = uploaded_file.read()

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    total_paginas = len(pdf.pages)

    st.sidebar.header("2. Escolha a página")
    pagina_num = st.sidebar.number_input(
        f"Página (1 a {total_paginas})", min_value=1, max_value=total_paginas, value=1
    )

    page = pdf.pages[pagina_num - 1]

    ESTRATEGIAS = {
        "Automática (por linhas/bordas)": {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
        "Baseada em texto (sem bordas)": {"vertical_strategy": "text", "horizontal_strategy": "text"},
        "Híbrida (linhas verticais + texto horizontal)": {"vertical_strategy": "lines", "horizontal_strategy": "text"},
    }

    if "estrategia_deteccao" not in st.session_state:
        st.session_state.estrategia_deteccao = "Automática (por linhas/bordas)"

    estrategia_nome = st.sidebar.selectbox(
        "Estratégia de detecção de tabela",
        list(ESTRATEGIAS.keys()),
        index=list(ESTRATEGIAS.keys()).index(st.session_state.estrategia_deteccao),
        help="Se a tabela não tem bordas desenhadas (só cor de fundo), tente "
             "'Baseada em texto'. Se algumas colunas ficarem grudadas, tente 'Híbrida'.",
    )
    st.session_state.estrategia_deteccao = estrategia_nome

    found_tables = page.find_tables(table_settings=ESTRATEGIAS[estrategia_nome])

    # fallback automático: se a estratégia escolhida não achou nada, testa as outras
    if not found_tables:
        st.warning(f"Nenhuma tabela encontrada com a estratégia '{estrategia_nome}'.")
        for nome_alt, settings_alt in ESTRATEGIAS.items():
            if nome_alt == estrategia_nome:
                continue
            tentativa = page.find_tables(table_settings=settings_alt)
            if tentativa:
                st.info(
                    f"🔎 A estratégia **'{nome_alt}'** encontrou {len(tentativa)} tabela(s) "
                    f"nesta página. Clique abaixo para usar essa estratégia."
                )
                if st.button(f"Usar estratégia '{nome_alt}'", key=f"usar_{nome_alt}"):
                    st.session_state.estrategia_deteccao = nome_alt
                    st.rerun()

    raw_tables = []
    opcoes = []
    for idx, t in enumerate(found_tables):
        raw = t.extract()
        if not raw or len(raw) < 1:
            continue
        raw_tables.append(raw)
        n_cols = max(len(r) for r in raw)
        opcoes.append(f"Tabela {idx + 1} — {len(raw)} linhas x {n_cols} colunas (bruto)")

    df_selecionado = None
    escolha_idx = None

    col_preview, col_selecao = st.columns([1.3, 1])

    with col_preview:
        st.subheader(f"Página {pagina_num} de {total_paginas}")
        im = page.to_image(resolution=150)
        if found_tables:
            im.draw_rects([t.bbox for t in found_tables], stroke="red", stroke_width=3)
        st.image(im.original, use_container_width=True, caption="Tabelas detectadas destacadas em vermelho")

    with col_selecao:
        st.subheader("3. Selecione a tabela")

        if not opcoes:
            st.warning("Nenhuma tabela detectada automaticamente nesta página.")
        else:
            tabelas_consideradas = st.multiselect(
                "Tabelas consideradas (desmarque para excluir):",
                opcoes,
                default=opcoes,
            )

            if not tabelas_consideradas:
                st.warning("Todas as tabelas foram excluídas. Marque ao menos uma para continuar.")
            else:
                escolha = st.radio("Qual tabela extrair?", tabelas_consideradas, index=0)
                escolha_idx = opcoes.index(escolha)
                raw_selecionado = raw_tables[escolha_idx]

                st.caption("Prévia BRUTA (sem assumir cabeçalho ainda) — primeiras linhas:")
                st.dataframe(pd.DataFrame(raw_selecionado[:8]), use_container_width=True)

                max_idx = len(raw_selecionado) - 1
                header_row_idx = st.number_input(
                    "Qual linha (0 = primeira linha acima) contém os nomes REAIS das colunas?",
                    min_value=0, max_value=max_idx, value=0,
                    help="Se a tabela tem cabeçalho mesclado em duas linhas (ex: um título "
                         "geral em cima e os nomes das colunas embaixo), escolha o índice "
                         "da linha de baixo, que é a que tem os nomes de verdade.",
                )

                # ---- edição manual do cabeçalho, para corrigir desalinhamentos ----
                raw_header = raw_selecionado[header_row_idx]
                n_cols_dados = max(len(r) for r in raw_selecionado[header_row_idx + 1:]) if len(raw_selecionado) > header_row_idx + 1 else len(raw_header)
                n_cols = max(len(raw_header), n_cols_dados)
                raw_header = list(raw_header) + [""] * (n_cols - len(raw_header))

                st.caption(
                    "Confira/edite os nomes de coluna abaixo antes de continuar "
                    "(corrija aqui se algum nome vier errado ou desalinhado):"
                )
                header_editado = []
                cols_header_widget = st.columns(4)
                for i, h in enumerate(raw_header):
                    default_val = str(h).strip() if h else f"coluna_{i}"
                    with cols_header_widget[i % 4]:
                        val = st.text_input(
                            f"Coluna {i}", value=default_val,
                            key=f"header_{pagina_num}_{escolha_idx}_{header_row_idx}_{i}",
                        )
                    header_editado.append(val)

                # garante nomes únicos mesmo após edição manual
                seen = {}
                header_final = []
                for h in header_editado:
                    if h in seen:
                        seen[h] += 1
                        header_final.append(f"{h}_{seen[h]}")
                    else:
                        seen[h] = 0
                        header_final.append(h)

                data_rows = raw_selecionado[header_row_idx + 1:]
                data_rows_ajustadas = [
                    list(r) + [None] * (n_cols - len(r)) for r in data_rows
                ]
                df_selecionado = pd.DataFrame(data_rows_ajustadas, columns=header_final).dropna(how="all")

                st.caption("Prévia já com cabeçalho aplicado:")
                st.dataframe(df_selecionado.head(6), use_container_width=True)

    # ==========================================================================
    # MODO 1: ETL GENÉRICO
    # ==========================================================================
    if modo.startswith("ETL Genérico") and df_selecionado is not None:
        st.divider()
        st.subheader("4. Confira e ajuste as colunas (opcional)")
        st.caption(
            "Renomeie para nomes padronizados se quiser manter consistência entre "
            "anos diferentes. O app lembra renomeações já feitas e sugere de novo."
        )

        novos_nomes = {}
        cols_widget = st.columns(min(len(df_selecionado.columns), 4) or 1)
        for i, col_original in enumerate(df_selecionado.columns):
            sugestao = st.session_state.coluna_aliases.get(normalize(col_original), col_original)
            with cols_widget[i % len(cols_widget)]:
                novo_nome = st.text_input(
                    f"'{col_original}' ->", value=sugestao,
                    key=f"rename_{pagina_num}_{escolha_idx}_{header_row_idx}_{i}",
                )
                novos_nomes[col_original] = novo_nome

        if st.button("✅ Adicionar esta tabela ao conjunto final", type="primary"):
            df_final = df_selecionado.rename(columns=novos_nomes)

            for original, novo in novos_nomes.items():
                st.session_state.coluna_aliases[normalize(original)] = novo

            df_final["_arquivo_origem"] = uploaded_file.name
            df_final["_pagina_origem"] = pagina_num
            df_final["_tabela_origem"] = escolha_idx + 1

            st.session_state.consolidado.append({
                "tipo": "generico",
                "df": df_final,
                "origem": f"[Genérico] {uploaded_file.name} (pág. {pagina_num}, tabela {escolha_idx + 1})",
            })
            st.success("Tabela adicionada!")
            st.rerun()

    # ==========================================================================
    # MODO 2: COMITÊS X PROGRAMAS (DESPIVOTAR)
    # ==========================================================================
    elif modo.startswith("Comitês") and df_selecionado is not None:
        st.divider()
        st.subheader("4. Configure a despivotagem")

        colunas_disponiveis = list(df_selecionado.columns)

        comite_col = st.selectbox("Qual coluna é o Comitê?", colunas_disponiveis, index=0)

        sugestao_programa = [
            c for c in colunas_disponiveis
            if c != comite_col and re.match(r"^[A-Za-z]{0,4}\s*\d+$", str(c).strip())
        ]
        programa_cols = st.multiselect(
            "Quais colunas são Programas (serão despivotadas)?",
            [c for c in colunas_disponiveis if c != comite_col],
            default=sugestao_programa,
        )

        # ---- exclusão de linhas específicas (ex: "Total por Programa") ----
        comites_unicos = [
            str(v).strip() for v in df_selecionado[comite_col].dropna().unique()
            if str(v).strip() not in ("", "-")
        ]
        comites_incluidos = st.multiselect(
            "Comitês (linhas) a incluir — desmarque para excluir linhas de total/rodapé:",
            comites_unicos,
            default=comites_unicos,
        )

        ano_input = st.number_input(
            "Ano de referência desta tabela",
            min_value=1900, max_value=2100,
            value=datetime.now().year, step=1,
        )

        df_filtrado = df_selecionado[
            df_selecionado[comite_col].astype(str).str.strip().isin(comites_incluidos)
        ]

        if programa_cols and comites_incluidos:
            preview_item = {
                "tipo": "comite_programa",
                "raw_df": df_filtrado,
                "comite_col": comite_col,
                "programa_cols": programa_cols,
                "ano": ano_input,
            }
            df_preview = get_df(preview_item)
            st.caption(f"Prévia da transformação ({df_preview.shape[0]} linhas):")
            st.dataframe(
                df_preview.drop(columns=["_arquivo_origem", "_pagina_origem"]).head(15),
                use_container_width=True,
            )
        else:
            st.info("Selecione ao menos uma coluna de Programa e um Comitê para ver a prévia.")

        if st.button(
            "✅ Transformar e adicionar ao conjunto final",
            type="primary",
            disabled=not (programa_cols and comites_incluidos),
        ):
            st.session_state.consolidado.append({
                "tipo": "comite_programa",
                "raw_df": df_filtrado,
                "comite_col": comite_col,
                "programa_cols": programa_cols,
                "ano": ano_input,
                "arquivo_origem": uploaded_file.name,
                "pagina_origem": pagina_num,
                "origem": f"[Comitê x Programa] {uploaded_file.name} (pág. {pagina_num}, tabela {escolha_idx + 1}) — Ano {ano_input}",
            })
            st.success("Tabela despivotada e adicionada!")
            st.rerun()

# ==============================================================================
# CONJUNTO FINAL / EXPORTAÇÃO
# ==============================================================================
st.divider()
st.header("5. Conjunto final")

if not st.session_state.consolidado:
    st.info("Nenhuma tabela adicionada ainda.")
else:
    for i, item in enumerate(st.session_state.consolidado):
        c1, c2, c3 = st.columns([4, 1.3, 1])
        with c1:
            st.write(f"**{i + 1}. {item['origem']}**")
        with c2:
            if item["tipo"] == "comite_programa":
                novo_ano = st.number_input(
                    "Ano do lote", min_value=1900, max_value=2100,
                    value=item["ano"], step=1, key=f"ano_lote_{i}",
                    label_visibility="collapsed",
                )
                item["ano"] = novo_ano
        with c3:
            if st.button("Remover", key=f"remove_{i}"):
                st.session_state.consolidado.pop(i)
                st.rerun()

        df_item = get_df(item)
        st.caption(f"{df_item.shape[0]} linhas x {df_item.shape[1]} colunas")

    with st.expander("Ver prévia consolidada"):
        try:
            consolidado_df = pd.concat(
                [get_df(item) for item in st.session_state.consolidado], ignore_index=True
            )
            st.dataframe(consolidado_df, use_container_width=True)
        except Exception as e:
            st.warning(f"Não foi possível consolidar automaticamente: {e}")

    def gerar_excel() -> bytes:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            for i, item in enumerate(st.session_state.consolidado):
                sheet_name = f"Tabela_{i+1}"[:31]
                get_df(item).to_excel(writer, sheet_name=sheet_name, index=False)
            try:
                consolidado_df = pd.concat(
                    [get_df(item) for item in st.session_state.consolidado], ignore_index=True
                )
                consolidado_df.to_excel(writer, sheet_name="Consolidado", index=False)
            except Exception:
                pass
        return buffer.getvalue()

    excel_bytes = gerar_excel()
    st.download_button(
        "⬇️ Baixar Excel consolidado",
        data=excel_bytes,
        file_name=f"dados_consolidados_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if st.button("🗑️ Limpar tudo"):
        st.session_state.consolidado = []
        st.rerun()