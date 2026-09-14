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
- SELEÇÃO VISUAL (CROP): Agora você delimita a tabela desenhando um 
  quadrado na tela com o mouse!
- COMITÊS SEM INVESTIMENTO: Valores vazios ou com tracinhos (-) agora
  são registrados como 0.0 em vez de serem ignorados. Assim, todos os 
  comitês aparecem no banco de dados e no Excel.
- O campo Ano fica vinculado ao lote inteiro e atualiza automaticamente.
- INTEGRAÇÃO COM BANCO DE DADOS (SQLite) ativa para a despivotagem.

COMO RODAR
----------
    pip install streamlit pdfplumber pandas openpyxl Pillow streamlit-cropper
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
from streamlit_cropper import st_cropper

inicializar_banco()  # garante que o banco e as tabelas existem
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


def build_header_and_df(raw_rows: list, header_row_idx: int) -> pd.DataFrame:
    header = raw_rows[header_row_idx]
    header = [str(h).strip() if h else f"coluna_{i}" for i, h in enumerate(header)]
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

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    total_paginas = len(pdf.pages)

    st.sidebar.header("2. Escolha a página")
    pagina_num = st.sidebar.number_input(
        f"Página (1 a {total_paginas})", min_value=1, max_value=total_paginas, value=1
    )

    page = pdf.pages[pagina_num - 1]

    # ==========================================================================
    # CROP VISUAL DA TABELA
    # ==========================================================================
    st.markdown("---")
    st.markdown("### 3. Delimite a Tabela na Página")
    st.caption(
        "Desenhe um retângulo no quadro abaixo com o mouse, envolvendo **apenas** a tabela "
        "que você quer extrair (ignore títulos, textos soltos e cabeçalhos da página)."
    )
    
    col_crop, col_selecao = st.columns([1.3, 1])

    with col_crop:
        resolucao_dpi = 150
        pil_image = page.to_image(resolution=resolucao_dpi).original
        
        box = st_cropper(pil_image, realtime_update=True, box_color='#FF0000', aspect_ratio=None, return_type='box')
        
        fator_escala = 72.0 / resolucao_dpi
        left = box['left'] * fator_escala
        top = box['top'] * fator_escala
        width = box['width'] * fator_escala
        height = box['height'] * fator_escala

        if width > 10 and height > 10:
            crop_bbox = (left, top, left + width, top + height)
            try:
                page_cropped = page.crop(crop_bbox)
            except Exception as e:
                st.warning(f"Erro ao cortar a página: {e}")
                page_cropped = page
        else:
            page_cropped = page

    with col_selecao:
        st.subheader("4. Seleção e Extração")
        
        estrategia_nome = st.selectbox(
            "Estratégia de detecção",
            ["Automática (por linhas/bordas)", "Baseada em texto (sem bordas)", "Híbrida (linhas verticais + texto horizontal)"],
            help="Como você já cortou a tabela, 'Baseada em texto' ou 'Híbrida' devem funcionar bem."
        )

        tolerancia = 3.0
        if "texto" in estrategia_nome.lower() or "híbrida" in estrategia_nome.lower():
            tolerancia = st.slider(
                "Juntar letras separadas (Tolerância X)", 
                min_value=1.0, max_value=15.0, value=3.0, step=0.5,
                help="Se as palavras do cabeçalho estiverem quebrando em várias colunas (ex: P R O G R A), aumente esse número."
            )

        ESTRATEGIAS = {
            "Automática (por linhas/bordas)": {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
            "Baseada em texto (sem bordas)": {"vertical_strategy": "text", "horizontal_strategy": "text", "text_x_tolerance": tolerancia},
            "Híbrida (linhas verticais + texto horizontal)": {"vertical_strategy": "lines", "horizontal_strategy": "text", "text_x_tolerance": tolerancia},
        }

        found_tables = page_cropped.find_tables(table_settings=ESTRATEGIAS[estrategia_nome])

        if not found_tables:
            st.warning("Nenhuma tabela encontrada na área selecionada.")
        
        raw_tables = []
        opcoes = []
        for idx, t in enumerate(found_tables):
            raw = t.extract()
            if not raw or len(raw) < 1:
                continue
            raw_tables.append(raw)
            n_cols = max(len(r) for r in raw)
            opcoes.append(f"Tabela {idx + 1} — {len(raw)} linhas x {n_cols} colunas")

        df_selecionado = None
        escolha_idx = None

        if opcoes:
            tabelas_consideradas = st.multiselect("Tabelas encontradas na área:", opcoes, default=opcoes)

            if tabelas_consideradas:
                escolha = st.radio("Qual extrair?", tabelas_consideradas, index=0, label_visibility="collapsed")
                escolha_idx = opcoes.index(escolha)
                raw_selecionado = raw_tables[escolha_idx]

                st.caption("Prévia BRUTA (todas as linhas):")
                st.dataframe(pd.DataFrame(raw_selecionado), use_container_width=True)

                max_idx = len(raw_selecionado) - 1
                header_row_idx = st.number_input(
                    "Qual linha (0 = primeira linha) contém os nomes REAIS das colunas?",
                    min_value=0, max_value=max_idx, value=0,
                )

                raw_header = raw_selecionado[header_row_idx]
                n_cols_dados = max(len(r) for r in raw_selecionado[header_row_idx + 1:]) if len(raw_selecionado) > header_row_idx + 1 else len(raw_header)
                n_cols = max(len(raw_header), n_cols_dados)
                raw_header = list(raw_header) + [""] * (n_cols - len(raw_header))

                st.caption("Confira/edite os nomes de coluna (corrija se houver desalinhamento):")
                header_editado = []
                cols_header_widget = st.columns(3)
                for i, h in enumerate(raw_header):
                    default_val = str(h).strip() if h else f"coluna_{i}"
                    with cols_header_widget[i % 3]:
                        val = st.text_input(
                            f"Coluna {i}", value=default_val,
                            key=f"header_{pagina_num}_{escolha_idx}_{header_row_idx}_{i}",
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

                st.caption("Prévia já com cabeçalho aplicado:")
                st.dataframe(df_selecionado.head(10), use_container_width=True)

    # ==========================================================================
    # MODO 1: ETL GENÉRICO
    # ==========================================================================
    if modo.startswith("ETL Genérico") and df_selecionado is not None:
        st.divider()
        st.subheader("5. Confira e ajuste as colunas (opcional)")

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
        st.subheader("5. Configure a despivotagem")

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
st.header("6. Conjunto final")

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