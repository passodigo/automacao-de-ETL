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
