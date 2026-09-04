"""
================================================================================
 ETL de PDF -> Excel com interface visual (Streamlit)
================================================================================
O usuário:
  1. Faz upload de um ou mais PDFs (um por ano, por exemplo).
  2. Para cada PDF, navega pelas páginas e VÊ a tabela destacada em vermelho.
  3. Se houver mais de uma tabela na página, escolhe qual quer extrair.
  4. Confere/edita o preview da tabela e opcionalmente renomeia colunas para
     nomes padronizados (mapeamento reaproveitado entre documentos).
  5. Adiciona a tabela escolhida ao conjunto final (pode repetir para vários
     PDFs/páginas -> consolida tudo).
  6. Baixa o Excel final consolidado, com uma aba por "lote" e uma aba
     "Consolidado" com tudo junto.

COMO RODAR
----------
    pip install streamlit pdfplumber pandas openpyxl Pillow
    streamlit run app.py
================================================================================
"""

import io
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st
from PIL import Image

st.set_page_config(page_title="PDF -> Excel", layout="wide")

# ==============================================================================
# ESTADO DA SESSÃO
# ==============================================================================
if "consolidado" not in st.session_state:
    # lista de dicts: {"df": DataFrame, "origem": str}
    st.session_state.consolidado = []

if "coluna_aliases" not in st.session_state:
    # lembra renomeações já feitas, para sugerir de novo em próximos uploads
    # formato: {nome_original_normalizado: nome_final_escolhido}
    st.session_state.coluna_aliases = {}


def normalize(s: str) -> str:
    return " ".join(str(s).strip().lower().split())


# ==============================================================================
# SIDEBAR: upload e navegação
# ==============================================================================
st.sidebar.header("1. Upload do PDF")
uploaded_file = st.sidebar.file_uploader("Selecione um PDF", type=["pdf"])

if uploaded_file is None:
    st.title("📄 ETL de PDF para Excel")
    st.info("Envie um PDF na barra lateral para começar.")
    if st.session_state.consolidado:
        st.subheader("Tabelas já adicionadas nesta sessão")
        for i, item in enumerate(st.session_state.consolidado):
            st.write(f"**{i+1}. {item['origem']}** — {item['df'].shape[0]} linhas x {item['df'].shape[1]} colunas")
    st.stop()

pdf_bytes = uploaded_file.read()

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    total_paginas = len(pdf.pages)

    st.sidebar.header("2. Escolha a página")
    pagina_num = st.sidebar.number_input(
        f"Página (1 a {total_paginas})", min_value=1, max_value=total_paginas, value=1
    )

    page = pdf.pages[pagina_num - 1]

    # ------------------------------------------------------------------------
    # Extrai tabelas da página e monta a imagem com destaque (retângulos)
    # ------------------------------------------------------------------------
    found_tables = page.find_tables()

    im = page.to_image(resolution=150)
    if found_tables:
        im.draw_rects([t.bbox for t in found_tables], stroke="red", stroke_width=3)
    pil_image: Image.Image = im.original

    col_preview, col_selecao = st.columns([1.3, 1])

    with col_preview:
        st.subheader(f"Página {pagina_num} de {total_paginas}")
        st.image(pil_image, use_container_width=True, caption="Tabelas detectadas destacadas em vermelho")

    with col_selecao:
        st.subheader("3. Selecione a tabela")

        if not found_tables:
            st.warning("Nenhuma tabela detectada automaticamente nesta página.")
            st.caption(
                "Isso pode acontecer se a tabela usa linhas/bordas muito finas ou "
                "espaçamento por texto em vez de grade. Tente outra página ou "
                "ajuste manualmente depois."
            )
        else:
            opcoes = []
            dfs_extraidos = []
            for idx, t in enumerate(found_tables):
                raw = t.extract()
                if not raw or len(raw) < 1:
                    continue
                header, *rows = raw
                header = [h.strip() if h else f"coluna_{i}" for i, h in enumerate(header)]
                df_tmp = pd.DataFrame(rows, columns=header).dropna(how="all")
                dfs_extraidos.append(df_tmp)
                opcoes.append(f"Tabela {idx + 1} — {df_tmp.shape[0]} linhas x {df_tmp.shape[1]} colunas")

            escolha = st.radio("Tabelas encontradas nesta página:", opcoes, index=0)
            escolha_idx = opcoes.index(escolha)
            df_selecionado = dfs_extraidos[escolha_idx].copy()

            st.caption("Prévia (primeiras linhas):")
            st.dataframe(df_selecionado.head(5), use_container_width=True)

    # ------------------------------------------------------------------------
    # Mapeamento / renomeação de colunas antes de consolidar
    # ------------------------------------------------------------------------
    if found_tables and 'df_selecionado' in dir():
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
                    f"'{col_original}' ->", value=sugestao, key=f"rename_{pagina_num}_{escolha_idx}_{i}"
                )
                novos_nomes[col_original] = novo_nome

        if st.button("✅ Adicionar esta tabela ao conjunto final", type="primary"):
            df_final = df_selecionado.rename(columns=novos_nomes)

            # memoriza os mapeamentos para reaproveitar em próximos PDFs
            for original, novo in novos_nomes.items():
                st.session_state.coluna_aliases[normalize(original)] = novo

            df_final["_arquivo_origem"] = uploaded_file.name
            df_final["_pagina_origem"] = pagina_num
            df_final["_tabela_origem"] = escolha_idx + 1

            st.session_state.consolidado.append({
                "df": df_final,
                "origem": f"{uploaded_file.name} (pág. {pagina_num}, tabela {escolha_idx + 1})",
            })
            st.success("Tabela adicionada! Você pode navegar para outra página/PDF e repetir.")
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
        c1, c2 = st.columns([5, 1])
        with c1:
            st.write(f"**{i + 1}. {item['origem']}** — {item['df'].shape[0]} linhas x {item['df'].shape[1]} colunas")
        with c2:
            if st.button("Remover", key=f"remove_{i}"):
                st.session_state.consolidado.pop(i)
                st.rerun()

    with st.expander("Ver prévia consolidada"):
        try:
            consolidado_df = pd.concat([item["df"] for item in st.session_state.consolidado], ignore_index=True)
            st.dataframe(consolidado_df, use_container_width=True)
        except Exception as e:
            st.warning(f"As tabelas têm colunas muito diferentes para consolidar automaticamente: {e}")
            consolidado_df = None

    def gerar_excel() -> bytes:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            for i, item in enumerate(st.session_state.consolidado):
                sheet_name = f"Tabela_{i+1}"[:31]
                item["df"].to_excel(writer, sheet_name=sheet_name, index=False)
            try:
                consolidado_df = pd.concat([item["df"] for item in st.session_state.consolidado], ignore_index=True)
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