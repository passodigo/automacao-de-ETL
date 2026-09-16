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

PONTOS IMPORTANTES DESTA VERSÃO
--------------------------------
- EXTRAÇÃO VIA DOCLING: a detecção e reconstrução da estrutura da tabela
  agora usa o Docling (modelo TableFormer), em vez do pdfplumber. Isso
  resolve boa parte dos problemas de células mescladas, cabeçalhos em
  duas linhas e tabelas sem bordas desenhadas (só cor de fundo).
  O pdfplumber continua sendo usado só para renderizar a IMAGEM da
  página como referência visual (não participa mais da extração).
- MODO ACCURATE: o Docling roda com TableFormerMode.ACCURATE (em vez do
  padrão FAST). É mais lento, mas detecta bem melhor páginas com VÁRIAS
  tabelas próximas umas das outras (o modo FAST às vezes funde ou
  descarta tabelas menores nesse cenário).
- REFORÇO MANUAL POR CROP: se ainda assim o Docling não pegar alguma
  tabela da página, dá pra desenhar um recorte manual só naquela área —
  o recorte também é processado pelo Docling (não pelo pdfplumber), então
  o motor de extração continua sendo 100% Docling; o crop só serve pra
  apontar onde ele deve reanalisar.
- COMITÊS SEM INVESTIMENTO: valores vazios ou com tracinhos (-) são
  registrados como 0.0 em vez de serem ignorados. Assim, todos os
  comitês aparecem no banco de dados e no Excel.
- O campo Ano fica vinculado ao lote inteiro e atualiza automaticamente.
- INTEGRAÇÃO COM BANCO DE DADOS (SQLite) ativa para a despivotagem.

COMO RODAR
----------
    pip install streamlit docling pdfplumber pypdf streamlit-cropper pandas openpyxl Pillow
    streamlit run app.py

OBS: a primeira conversão de cada PDF com o Docling pode demorar alguns
segundos (ele baixa/carrega os modelos de layout e de estrutura de
tabela na primeira execução). Resultado fica em cache na sessão.
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
    """Converte valores em texto para float. Tracinhos ou vazio viram 0.0."""
    if v is None:
        return 0.0
    s = str(v).strip()
    if s == "" or s in ("-", "—", "–", "nan", "None"):
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
    if s in ("", "-"):
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        partes = s.split(".")
        if len(partes) > 1 and len(partes[-1]) == 3:
            s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def extract_programa_label(col_name: str) -> str:
    return str(col_name).strip()


def get_df(item: dict) -> pd.DataFrame:
    if item["tipo"] == "generico":
        return item["df"]

    raw_df = item["raw_df"]
    comite_col = item["comite_col"]
    programa_cols = item["programa_cols"]
    ano = item["ano"]

    linhas = []
    for _, row in raw_df.iterrows():
        comite = row.get(comite_col)
        # Ignora a linha toda apenas se o nome do comitê estiver vazio
        if comite is None or str(comite).strip() == "" or str(comite).strip() == "-":
            continue
        for col in programa_cols:
            valor = parse_valor_brl(row.get(col))
            # Não ignoramos mais os valores nulos; eles entram como 0.0
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
# FUNÇÕES DE BANCO DE DADOS
# ==============================================================================
def get_ou_criar_comite(conn, nome_comite: str) -> int:
    nome = str(nome_comite).strip().upper()
    cur = conn.cursor()
    cur.execute("SELECT id FROM comite WHERE nome = ?", (nome,))
    row = cur.fetchone()
    if row:
        return row[0]

    cur.execute("INSERT INTO comite (nome) VALUES (?)", (nome,))
    return cur.lastrowid


def get_ou_criar_programa(conn, codigo_programa: str) -> int:
    codigo = str(codigo_programa).strip().upper()
    cur = conn.cursor()
    cur.execute("SELECT id FROM programa WHERE codigo = ?", (codigo,))
    row = cur.fetchone()
    if row:
        return row[0]

    cur.execute("INSERT INTO programa (codigo) VALUES (?)", (codigo,))
    return cur.lastrowid


def get_ou_criar_exercicio(conn, ano: int) -> int:
    cur = conn.cursor()
    cur.execute("SELECT id FROM exercicio WHERE ano = ?", (ano,))
    row = cur.fetchone()
    if row:
        return row[0]

    cur.execute("INSERT INTO exercicio (ano) VALUES (?)", (ano,))
    return cur.lastrowid


# ==============================================================================
# DOCLING: CONVERSÃO DO PDF E RECONSTRUÇÃO DA TABELA COMO MATRIZ CRUA
# ==============================================================================

@st.cache_resource
def get_docling_converter() -> DocumentConverter:
    """O conversor carrega modelos pesados (layout + estrutura de tabela).
    Cacheado como 'resource' pra não recarregar isso a cada interação.

    Usa TableFormerMode.ACCURATE em vez do padrão FAST: é mais lento, mas
    detecta bem melhor páginas com várias tabelas próximas (o modo FAST
    tende a fundir ou descartar as menores nesse cenário)."""
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_table_structure = True
    pipeline_options.table_structure_options.mode = TableFormerMode.ACCURATE

    return DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
        }
    )


def converter_pdf_com_docling(pdf_bytes: bytes, nome_arquivo: str):
    """Roda o Docling sobre o PDF inteiro uma única vez por arquivo e guarda
    o resultado na sessão (a conversão é mais pesada que abrir com pdfplumber,
    então vale evitar refazer isso a cada rerun do Streamlit)."""
    converter = get_docling_converter()
    source = DocumentStream(name=nome_arquivo, stream=io.BytesIO(pdf_bytes))
    resultado = converter.convert(source)
    return resultado.document


def cropar_pagina_pdf(pdf_bytes: bytes, pagina_num: int, bbox_pdf: tuple) -> bytes:
    """Gera um PDF de 1 página só com a área recortada (em pontos PDF,
    origem no canto inferior esquerdo). Usado para reprocessar com o
    Docling apenas a região que o usuário apontou manualmente."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    writer.add_page(reader.pages[pagina_num - 1])

    x0, y0, x1, y1 = bbox_pdf
    nova_pagina = writer.pages[0]
    nova_pagina.mediabox.lower_left = (x0, y0)
    nova_pagina.mediabox.upper_right = (x1, y1)

    buffer = io.BytesIO()
    writer.write(buffer)
    return buffer.getvalue()


def box_cropper_para_bbox_pdf(box: dict, resolucao_dpi: int, altura_pagina_pt: float) -> tuple:
    """Converte a seleção do st_cropper (pixels, origem no topo-esquerda,
    na resolução usada pra renderizar a imagem) para coordenadas PDF em
    pontos (origem embaixo-à-esquerda, 72 pontos por polegada)."""
    fator_escala = 72.0 / resolucao_dpi
    x0 = box["left"] * fator_escala
    x1 = (box["left"] + box["width"]) * fator_escala
    topo_pt = box["top"] * fator_escala
    base_pt = (box["top"] + box["height"]) * fator_escala
    # inverte o eixo Y: PDF mede a partir de baixo, a imagem mede a partir de cima
    y0 = altura_pagina_pt - base_pt
    y1 = altura_pagina_pt - topo_pt
    return (x0, y0, x1, y1)


def docling_table_para_matriz(table) -> list[list[str]]:
    """Reconstrói a tabela como uma matriz de texto (lista de listas),
    propagando o texto de células mescladas para todas as posições que elas
    cobrem — o mesmo problema que tínhamos antes com cabeçalhos mesclados
    no pdfplumber/openpyxl, mas aqui resolvido com a estrutura que o próprio
    Docling já identificou (row_span/col_span), sem precisar de heurística
    manual de merge."""
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
    """Docling numera páginas a partir de 1, igual ao seletor da sidebar."""
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
            df_preview = get_df(item)
            st.write(f"**{i+1}. {item['origem']}** — {df_preview.shape[0]} linhas x {df_preview.shape[1]} colunas")
    st.stop()

pdf_bytes = uploaded_file.read()
file_id = (uploaded_file.name, uploaded_file.size)

# Reconverte com Docling só se for um arquivo novo (evita reprocessar a cada rerun)
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
    # {(file_id, pagina_num): [matriz1, matriz2, ...]} — tabelas achadas via
    # reforço manual (crop), fora do que o Docling detectou sozinho.
    st.session_state.tabelas_extra_por_pagina = {}

col_preview, col_selecao = st.columns([1.3, 1])

with col_preview:
    st.subheader(f"Página {pagina_num} de {total_paginas}")
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_preview:
        page_preview = pdf_preview.pages[pagina_num - 1]
        altura_pagina_pt = page_preview.height
        im = page_preview.to_image(resolution=RESOLUCAO_PREVIEW_DPI)
        pil_image = im.original
        st.image(pil_image, use_container_width=True, caption="Referência visual da página (extração é feita pelo Docling)")

# tabelas que o Docling encontrou sozinho, nessa página
tabelas_docling = [
    docling_table_para_matriz(t) for t in docling_doc.tables if pagina_da_tabela(t) == pagina_num
]

# tabelas encontradas via reforço manual (crop) em execuções anteriores desta sessão
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
                help="O Docling já reconstrói células mescladas automaticamente, mas ainda "
                     "assim confira aqui qual linha tem os nomes de coluna de verdade — "
                     "algumas tabelas têm uma linha de título geral acima do cabeçalho.",
            )

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

# ==============================================================================
# REFORÇO MANUAL: crop de uma área da página, reprocessada pelo próprio Docling
# ==============================================================================
with st.expander("🔍 O Docling não achou todas as tabelas desta página? Recorte manualmente"):
    st.caption(
        "Desenhe um retângulo envolvendo APENAS a tabela que ficou faltando. Esse recorte "
        "não é extraído por outra biblioteca — ele também é processado pelo Docling, só que "
        "isolado, o que costuma resolver os casos de tabelas próximas demais umas das outras."
    )

    box = st_cropper(
        pil_image,
        realtime_update=True,
        box_color="#FF0000",
        aspect_ratio=None,
        return_type="box",
        key=f"cropper_{pagina_num}",
    )

    if st.button("↻ Reprocessar essa área com Docling", key=f"reprocessar_{pagina_num}"):
        if box["width"] < 10 or box["height"] < 10:
            st.warning("Selecione uma área maior antes de reprocessar.")
        else:
            bbox_pdf = box_cropper_para_bbox_pdf(box, RESOLUCAO_PREVIEW_DPI, altura_pagina_pt)
            pdf_recortado = cropar_pagina_pdf(pdf_bytes, pagina_num, bbox_pdf)

            with st.spinner("Reprocessando o recorte com Docling..."):
                converter = get_docling_converter()
                source_recorte = DocumentStream(
                    name=f"{uploaded_file.name}_crop_p{pagina_num}", stream=io.BytesIO(pdf_recortado)
                )
                resultado_recorte = converter.convert(source_recorte)
                doc_recortado = resultado_recorte.document

            if not doc_recortado.tables:
                st.warning("O Docling não encontrou nenhuma tabela dentro da área recortada.")
            else:
                novas_matrizes = [docling_table_para_matriz(t) for t in doc_recortado.tables]
                st.session_state.tabelas_extra_por_pagina.setdefault(chave_extra, []).extend(novas_matrizes)
                st.success(f"{len(novas_matrizes)} tabela(s) encontrada(s) no recorte e adicionada(s) à lista acima!")
                st.rerun()

    if tabelas_extra:
        if st.button("🗑️ Limpar tabelas de reforço manual desta página", key=f"limpar_extra_{pagina_num}"):
            st.session_state.tabelas_extra_por_pagina.pop(chave_extra, None)
            st.rerun()

# ==============================================================================
# MODO 1: ETL GENÉRICO
# ==============================================================================
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
        try:
            conn = get_connection()
            exercicio_id = get_ou_criar_exercicio(conn, ano_input)
            df_para_salvar = get_df(preview_item)

            cur = conn.cursor()
            for _, row in df_para_salvar.iterrows():
                comite_nome = row["Comitê"]
                programa_codigo = row["Programa"]
                valor = row["Valor"]

                comite_id = get_ou_criar_comite(conn, comite_nome)
                programa_id = get_ou_criar_programa(conn, programa_codigo)

                cur.execute("""
                    INSERT INTO investimento (comite_id, programa_id, exercicio_id, valor)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(comite_id, programa_id, exercicio_id)
                    DO UPDATE SET valor = excluded.valor
                """, (comite_id, programa_id, exercicio_id, float(valor)))

            conn.commit()
            st.toast("✅ Dados salvos com sucesso no banco de dados!")

        except Exception as e:
            st.error(f"Erro ao salvar no banco: {e}")
        finally:
            if 'conn' in locals():
                conn.close()

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