"""
================================================================================
 ETL de PDF -> Excel com interface visual (Streamlit)
================================================================================
Dois modos, na barra lateral:

  1) ETL Genérico — escolhe página, tabela, linha de cabeçalho, renomeia
     colunas livremente, adiciona ao conjunto final.

  2) Comitês x Programas (despivotar) — seção dedicada para tabelas em
     formato "matriz" (Comitê nas linhas, Programa nas colunas). Gera
     formato linear: Comitê | Programa | Ano | Valor. E salva no SQLite!

COMO RODAR
----------
    pip install streamlit docling pdfplumber pypdf streamlit-cropper pandas openpyxl Pillow
    streamlit run app.py
================================================================================
"""

import io
import re
from datetime import datetime
from db_setup import inicializar_banco, get_connection, DB_PATH
import pandas as pd
import pdfplumber
import streamlit as st
from pypdf import PdfReader, PdfWriter
from streamlit_cropper import st_cropper
from docling.document_converter import DocumentConverter, PdfFormatOption
from docling.datamodel.base_models import DocumentStream, InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode

inicializar_banco()  # garante que o banco e as tabelas existem
st.set_page_config(page_title="PDF -> Excel", layout="wide")

# ==============================================================================
# ESTADO DA SESSÃO
# ==============================================================================
if "consolidado" not in st.session_state:
    st.session_state.consolidado = []

if "coluna_aliases" not in st.session_state:
    st.session_state.coluna_aliases = {}

if "docling_doc" not in st.session_state:
    st.session_state.docling_doc = None

if "docling_file_id" not in st.session_state:
    st.session_state.docling_file_id = None


def normalize(s: str) -> str:
    return " ".join(str(s).strip().lower().split())


def parse_valor_brl(v):
    """Converte valores em texto para float, suportando vários padrões numéricos."""
    if pd.isna(v) or v is None:
        return 0.0
    
    s = str(v).strip()
    if not s or s.lower() in ("-", "—", "–", "nan", "none", "null"):
        return 0.0
    
    is_negative = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    
    # Mantém apenas números, ponto e vírgula
    s = re.sub(r"[^\d,.]", "", s)
    if not s:
        return 0.0
        
    # Padroniza formatos com ponto e vírgula
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        partes = s.split(",")
        if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3):
            s = s.replace(",", "") # Trata como milhar (US)
        else:
            s = s.replace(",", ".") # Trata como decimal (BR)
    elif "." in s:
        partes = s.split(".")
        if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3):
            s = s.replace(".", "") # Trata como milhar (BR)
            
    try:
        val = float(s)
        return -val if is_negative else val
    except ValueError:
        return 0.0


def aplicar_despivotagem(raw_df, comite_col, programa_cols, comites_incluidos, ano, remover_zerados):
    """
    Despivota a matriz perfeitamente usando pd.melt, agrupando por comitê e limpando vazios.
    """
    df_filtrado = raw_df[raw_df[comite_col].astype(str).str.strip().isin(comites_incluidos)].copy()
    df_filtrado = df_filtrado[[comite_col] + programa_cols]
    
    df_linear = df_filtrado.melt(
        id_vars=[comite_col],
        value_vars=programa_cols,
        var_name='Programa',
        value_name='Valor_Cru'
    )
    
    df_linear['Comitê'] = df_linear[comite_col].astype(str).str.strip()
    df_linear['Programa'] = df_linear['Programa'].astype(str).str.strip()
    df_linear['Ano'] = ano
    df_linear['Valor'] = df_linear['Valor_Cru'].apply(parse_valor_brl)
    
    # Ordena bonitinho para o Excel: Primeiro o Comitê, depois os programas dele
    df_linear = df_linear.sort_values(by=['Comitê', 'Programa']).reset_index(drop=True)
    
    # Mágica que limpa o lixo visual e deixa só quem recebeu investimento
    if remover_zerados:
        df_linear = df_linear[df_linear['Valor'] != 0.0].reset_index(drop=True)
        
    return df_linear[['Comitê', 'Programa', 'Ano', 'Valor']]


# ==============================================================================
# FUNÇÕES DE BANCO DE DADOS
# ==============================================================================
def get_ou_criar_comite(conn, nome_comite: str) -> int:
    nome = str(nome_comite).strip().upper()
    cur = conn.cursor()
    cur.execute("SELECT id FROM comite WHERE nome = ?", (nome,))
    row = cur.fetchone()
    if row: return row[0]
    cur.execute("INSERT INTO comite (nome) VALUES (?)", (nome,))
    return cur.lastrowid

def get_ou_criar_programa(conn, codigo_programa: str) -> int:
    codigo = str(codigo_programa).strip().upper()
    cur = conn.cursor()
    cur.execute("SELECT id FROM programa WHERE codigo = ?", (codigo,))
    row = cur.fetchone()
    if row: return row[0]
    cur.execute("INSERT INTO programa (codigo) VALUES (?)", (codigo,))
    return cur.lastrowid

def get_ou_criar_exercicio(conn, ano: int) -> int:
    cur = conn.cursor()
    cur.execute("SELECT id FROM exercicio WHERE ano = ?", (ano,))
    row = cur.fetchone()
    if row: return row[0]
    cur.execute("INSERT INTO exercicio (ano) VALUES (?)", (ano,))
    return cur.lastrowid


# ==============================================================================
# DOCLING: CONVERSÃO DO PDF E RECONSTRUÇÃO
# ==============================================================================
@st.cache_resource
def get_docling_converter() -> DocumentConverter:
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_table_structure = True
    pipeline_options.table_structure_options.mode = TableFormerMode.ACCURATE
    return DocumentConverter(
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)}
    )

def converter_pdf_com_docling(pdf_bytes: bytes, nome_arquivo: str):
    converter = get_docling_converter()
    source = DocumentStream(name=nome_arquivo, stream=io.BytesIO(pdf_bytes))
    return converter.convert(source).document

def cropar_pagina_pdf(pdf_bytes: bytes, pagina_num: int, bbox_pdf: tuple) -> bytes:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    page = reader.pages[pagina_num - 1]

    x0, y0, x1, y1 = bbox_pdf
    MARGEM = 30.0 
    
    lim_x0 = float(page.mediabox.left)
    lim_y0 = float(page.mediabox.bottom)
    lim_x1 = float(page.mediabox.right)
    lim_y1 = float(page.mediabox.top)

    novo_x0 = max(lim_x0, x0 - MARGEM)
    novo_y0 = max(lim_y0, y0 - MARGEM)
    novo_x1 = min(lim_x1, x1 + MARGEM)
    novo_y1 = min(lim_y1, y1 + MARGEM)

    page.mediabox.lower_left = (novo_x0, novo_y0)
    page.mediabox.upper_right = (novo_x1, novo_y1)
    page.cropbox.lower_left = (novo_x0, novo_y0)
    page.cropbox.upper_right = (novo_x1, novo_y1)

    writer.add_page(page)
    buffer = io.BytesIO()
    writer.write(buffer)
    return buffer.getvalue()

def box_cropper_para_bbox_pdf(box: dict, resolucao_dpi: int, altura_pagina_pt: float) -> tuple:
    fator_escala = 72.0 / resolucao_dpi
    x0 = box["left"] * fator_escala
    x1 = (box["left"] + box["width"]) * fator_escala
    topo_pt = box["top"] * fator_escala
    base_pt = (box["top"] + box["height"]) * fator_escala
    y0 = altura_pagina_pt - base_pt
    y1 = altura_pagina_pt - topo_pt
    return (x0, y0, x1, y1)

def docling_table_para_matriz(table) -> list[list[str]]:
    data = table.data
    n_linhas = data.num_rows
    n_colunas = data.num_cols
    matriz = [["" for _ in range(n_colunas)] for _ in range(n_linhas)]
    for cell in data.table_cells:
        texto = (cell.text or "").strip()
        for r in range(cell.start_row_offset_idx, cell.end_row_offset_idx):
            for c in range(cell.start_col_offset_idx, cell.end_col_offset_idx):
                if r < n_linhas and c < n_colunas:
                    matriz[r][c] = texto
    return matriz

def pagina_da_tabela(table) -> int:
    if table.prov:
        return table.prov[0].page_no
    return 1


# ==============================================================================
# SIDEBAR E LEITURA DO PDF
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
            df_preview = item["df"]
            st.write(f"**{i+1}. {item['origem']}** — {df_preview.shape[0]} linhas x {df_preview.shape[1]} colunas")
    st.stop()

pdf_bytes = uploaded_file.read()
file_id = (uploaded_file.name, uploaded_file.size)

if st.session_state.docling_file_id != file_id:
    with st.spinner("Processando PDF com Docling (pode levar alguns segundos na primeira vez)..."):
        st.session_state.docling_doc = converter_pdf_com_docling(pdf_bytes, uploaded_file.name)
        st.session_state.docling_file_id = file_id

docling_doc = st.session_state.docling_doc
total_paginas = len(docling_doc.pages)

st.sidebar.header("2. Escolha a página")
pagina_num = st.sidebar.number_input(
    f"Página (1 a {total_paginas})", min_value=1, max_value=total_paginas, value=1
)

RESOLUCAO_PREVIEW_DPI = 150
if "tabelas_extra_por_pagina" not in st.session_state:
    st.session_state.tabelas_extra_por_pagina = {}

col_preview, col_selecao = st.columns([1.3, 1])

with col_preview:
    st.subheader(f"Página {pagina_num} de {total_paginas}")
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_preview:
        page_preview = pdf_preview.pages[pagina_num - 1]
        altura_pagina_pt = page_preview.height
        im = page_preview.to_image(resolution=RESOLUCAO_PREVIEW_DPI)
        pil_image = im.original
        st.image(pil_image, use_container_width=True, caption="Referência visual da página")

tabelas_docling = [
    docling_table_para_matriz(t) for t in docling_doc.tables if pagina_da_tabela(t) == pagina_num
]

chave_extra = (file_id, pagina_num)
tabelas_extra = st.session_state.tabelas_extra_por_pagina.get(chave_extra, [])

with col_selecao:
    st.subheader("3. Selecione a tabela")

    raw_tables = tabelas_docling + tabelas_extra
    origem_tabelas = (["docling"] * len(tabelas_docling)) + (["reforço manual"] * len(tabelas_extra))

    if not raw_tables:
        st.warning("O Docling não encontrou tabelas nesta página.")
        opcoes = []
    else:
        opcoes = [
            f"Tabela {idx + 1} — {len(raw)} linhas x {max(len(r) for r in raw)} colunas"
            + (" (reforço manual)" if origem_tabelas[idx] == "reforço manual" else "")
            for idx, raw in enumerate(raw_tables)
        ]

    df_selecionado = None
    escolha_idx = None
    header_row_idx = 0

    if opcoes:
        tabelas_consideradas = st.multiselect(
            "Tabelas consideradas (desmarque para excluir):", opcoes, default=opcoes
        )

        if not tabelas_consideradas:
            st.warning("Todas as tabelas foram excluídas. Marque ao menos uma para continuar.")
        else:
            escolha = st.radio("Qual tabela extrair?", tabelas_consideradas, index=0)
            escolha_idx = opcoes.index(escolha)
            raw_selecionado = raw_tables[escolha_idx]

            st.caption("Prévia BRUTA — primeiras linhas:")
            st.dataframe(pd.DataFrame(raw_selecionado[:8]), use_container_width=True)

            max_idx = len(raw_selecionado) - 1
            header_row_idx = st.number_input(
                "Qual linha contém os nomes REAIS das colunas?",
                min_value=0, max_value=max_idx, value=0
            )

            raw_header = raw_selecionado[header_row_idx]
            n_cols_dados = max(len(r) for r in raw_selecionado[header_row_idx + 1:]) if len(raw_selecionado) > header_row_idx + 1 else len(raw_header)
            n_cols = max(len(raw_header), n_cols_dados)
            raw_header = list(raw_header) + [""] * (n_cols - len(raw_header))

            st.caption("Confira/edite os nomes de coluna abaixo:")
            header_editado = []
            cols_header_widget = st.columns(4)
            for i, h in enumerate(raw_header):
                default_val = str(h).strip() if h else f"coluna_{i}"
                with cols_header_widget[i % 4]:
                    val = st.text_input(
                        f"Coluna {i}", value=default_val,
                        key=f"header_{pagina_num}_{escolha_idx}_{header_row_idx}_{i}"
                    )
                header_editado.append(val)

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
            data_rows_ajustadas = [list(r) + [None] * (n_cols - len(r)) for r in data_rows]
            df_selecionado = pd.DataFrame(data_rows_ajustadas, columns=header_final).dropna(how="all")

            st.caption("Prévia com cabeçalho aplicado:")
            st.dataframe(df_selecionado.head(6), use_container_width=True)

with st.expander("🔍 O Docling não achou todas as tabelas desta página? Recorte manualmente"):
    box = st_cropper(
        pil_image, realtime_update=True, box_color="#FF0000",
        aspect_ratio=None, return_type="box", key=f"cropper_{pagina_num}"
    )
    if st.button("↻ Reprocessar essa área com Docling", key=f"reprocessar_{pagina_num}"):
        if box["width"] < 10 or box["height"] < 10:
            st.warning("Selecione uma área maior antes de reprocessar.")
        else:
            bbox_pdf = box_cropper_para_bbox_pdf(box, RESOLUCAO_PREVIEW_DPI, altura_pagina_pt)
            pdf_recortado = cropar_pagina_pdf(pdf_bytes, pagina_num, bbox_pdf)
            with st.spinner("Reprocessando recorte com Docling..."):
                converter = get_docling_converter()
                source_recorte = DocumentStream(name=f"crop_p{pagina_num}", stream=io.BytesIO(pdf_recortado))
                doc_recortado = converter.convert(source_recorte).document

            if not doc_recortado.tables:
                st.warning("Nenhuma tabela encontrada no recorte.")
            else:
                novas_matrizes = [docling_table_para_matriz(t) for t in doc_recortado.tables]
                st.session_state.tabelas_extra_por_pagina.setdefault(chave_extra, []).extend(novas_matrizes)
                st.success(f"Tabela encontrada e adicionada!")
                st.rerun()

    if tabelas_extra:
        if st.button("🗑️ Limpar reforço manual", key=f"limpar_extra_{pagina_num}"):
            st.session_state.tabelas_extra_por_pagina.pop(chave_extra, None)
            st.rerun()

# ==============================================================================
# MODO 1: ETL GENÉRICO
# ==============================================================================
if modo.startswith("ETL Genérico") and df_selecionado is not None:
    st.divider()
    st.subheader("4. Ajuste e Adição")

    novos_nomes = {}
    cols_widget = st.columns(min(len(df_selecionado.columns), 4) or 1)
    for i, col_original in enumerate(df_selecionado.columns):
        sugestao = st.session_state.coluna_aliases.get(normalize(col_original), col_original)
        with cols_widget[i % len(cols_widget)]:
            novo_nome = st.text_input(f"'{col_original}' ->", value=sugestao, key=f"rename_{pagina_num}_{escolha_idx}_{i}")
            novos_nomes[col_original] = novo_nome

    if st.button("✅ Adicionar ao conjunto final", type="primary"):
        df_final = df_selecionado.rename(columns=novos_nomes)
        for original, novo in novos_nomes.items():
            st.session_state.coluna_aliases[normalize(original)] = novo

        df_final["_arquivo_origem"] = uploaded_file.name
        df_final["_pagina_origem"] = pagina_num

        st.session_state.consolidado.append({
            "tipo": "generico",
            "df": df_final,
            "origem": f"[Genérico] {uploaded_file.name} (pág. {pagina_num})",
        })
        st.success("Tabela adicionada!")
        st.rerun()

# ==============================================================================
# MODO 2: COMITÊS X PROGRAMAS (DESPIVOTAR)
# ==============================================================================
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

    comites_unicos = [
        str(v).strip() for v in df_selecionado[comite_col].dropna().unique()
        if str(v).strip() not in ("", "-")
    ]
    comites_incluidos = st.multiselect(
        "Comitês (linhas) a incluir:",
        comites_unicos,
        default=comites_unicos,
    )

    col1, col2 = st.columns(2)
    with col1:
        ano_input = st.number_input("Ano de referência", min_value=1900, max_value=2100, value=datetime.now().year)
    with col2:
        remover_zerados = st.checkbox("🧹 Remover linhas vazias (Oculta R$ 0,00)", value=True, help="Desmarque se você quiser salvar no Excel os comitês que não receberam investimento em determinados programas.")

    if programa_cols and comites_incluidos:
        df_linear_preview = aplicar_despivotagem(df_selecionado, comite_col, programa_cols, comites_incluidos, ano_input, remover_zerados)
        
        st.caption(f"Prévia da transformação ({df_linear_preview.shape[0]} linhas encontradas):")
        if df_linear_preview.empty and remover_zerados:
            st.warning("Todas as linhas retornaram R$ 0,00. Isso indica que a 'Linha de Cabeçalho' no Passo 3 foi escolhida errada e os valores se perderam. Volte no Passo 3 e mude o número da linha.")
        else:
            st.dataframe(
                df_linear_preview.head(30).style.format({"Valor": "{:,.2f}"}),
                use_container_width=True,
            )
    else:
        st.info("Selecione ao menos um Programa e um Comitê para ver a prévia.")

    if st.button("✅ Transformar e salvar", type="primary", disabled=not (programa_cols and comites_incluidos)):
        try:
            df_para_salvar = aplicar_despivotagem(df_selecionado, comite_col, programa_cols, comites_incluidos, ano_input, remover_zerados)
            
            conn = get_connection()
            exercicio_id = get_ou_criar_exercicio(conn, ano_input)
            cur = conn.cursor()
            for _, row in df_para_salvar.iterrows():
                c_id = get_ou_criar_comite(conn, row["Comitê"])
                p_id = get_ou_criar_programa(conn, row["Programa"])
                cur.execute("""
                    INSERT INTO investimento (comite_id, programa_id, exercicio_id, valor)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(comite_id, programa_id, exercicio_id) DO UPDATE SET valor = excluded.valor
                """, (c_id, p_id, exercicio_id, float(row["Valor"])))
            conn.commit()
            st.toast("✅ Salvo no banco de dados!")
        except Exception as e:
            st.error(f"Erro ao salvar no banco: {e}")
        finally:
            if 'conn' in locals(): conn.close()

        df_para_salvar["_arquivo_origem"] = uploaded_file.name
        df_para_salvar["_pagina_origem"] = pagina_num
        st.session_state.consolidado.append({
            "tipo": "comite_programa",
            "df": df_para_salvar,
            "origem": f"[Comitê x Programa] {uploaded_file.name} (Ano {ano_input})",
        })
        st.success("Tabela despivotada adicionada!")
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
                ano_atual = int(item["df"]["Ano"].iloc[0]) if not item["df"].empty else 2026
                novo_ano = st.number_input(
                    "Ano do lote", min_value=1900, max_value=2100,
                    value=ano_atual, step=1, key=f"ano_lote_{i}",
                    label_visibility="collapsed"
                )
                if novo_ano != ano_atual:
                    item["df"]["Ano"] = novo_ano 
        with c3:
            if st.button("Remover", key=f"remove_{i}"):
                st.session_state.consolidado.pop(i)
                st.rerun()

        st.caption(f"{item['df'].shape[0]} linhas x {item['df'].shape[1]} colunas")

    with st.expander("Ver prévia consolidada"):
        try:
            consolidado_df = pd.concat([item["df"] for item in st.session_state.consolidado], ignore_index=True)
            if "Valor" in consolidado_df.columns:
                st.dataframe(consolidado_df.style.format({"Valor": "{:,.2f}"}), use_container_width=True)
            else:
                st.dataframe(consolidado_df, use_container_width=True)
        except Exception as e:
            st.warning(f"Não foi possível consolidar automaticamente: {e}")

    def gerar_excel() -> bytes:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            for i, item in enumerate(st.session_state.consolidado):
                sheet_name = f"Tabela_{i+1}"[:31]
                item["df"].to_excel(writer, sheet_name=sheet_name, index=False)
            try:
                pd.concat([item["df"] for item in st.session_state.consolidado], ignore_index=True).to_excel(writer, sheet_name="Consolidado", index=False)
            except Exception:
                pass
        return buffer.getvalue()

    st.download_button(
        "⬇️ Baixar Excel consolidado",
        data=gerar_excel(),
        file_name=f"dados_consolidados_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if st.button("🗑️ Limpar tudo"):
        st.session_state.consolidado = []
        st.rerun()