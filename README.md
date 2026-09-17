# Automação de ETL: PDF para Excel 

Uma aplicação web interativa construída com **Streamlit** para extração avançada de tabelas em arquivos PDF. Utiliza a inteligência artificial do **Docling (TableFormer)** para ler tabelas complexas (com células mescladas ou sem bordas) e convertê-las em conjuntos de dados estruturados e limpos.

O sistema possui dois fluxos de trabalho principais: um extrator genérico para qualquer tipo de tabela e um transformador específico para dados orçamentários (Matriz de Comitês x Programas), que realiza a despivotagem (unpivot) automática e salva os resultados em um banco de dados SQLite.

## 🚀 Principais Funcionalidades

* **Motor de Extração Inteligente (Docling):** Supera bibliotecas tradicionais ao reconhecer estruturas de tabelas invisíveis, cabeçalhos de múltiplas linhas e células mescladas sem perder o alinhamento.
* **Recorte Manual de Resgate:** Se a IA não detectar uma tabela em uma página complexa, o usuário pode desenhar um retângulo vermelho na tela. O sistema reprocessa cirurgicamente a área recortada, garantindo 100% de captura.
* **Modo 1: ETL Genérico:** 
  * Permite a seleção manual da "linha de cabeçalho" correta.
  * Renomeação customizada de colunas.
  * Consolidação de múltiplas tabelas (de páginas ou arquivos diferentes) em um único Excel final.
* **Modo 2: Despivotagem (Comitês x Programas):**
  * Converte matrizes cruzadas em formato tabular relacional (`Comitê | Programa | Ano | Valor`).
  * Tratamento automatizado de valores financeiros em BRL (limpa traços, pontos, formatações contábeis e textos nulos).
  * Filtro inteligente que remove linhas vazias (`R$ 0,00`) para não poluir o arquivo final.
  * Integração direta com Banco de Dados SQLite.

## 🛠️ Tecnologias Utilizadas

* **[Streamlit](https://streamlit.io/):** Interface gráfica e gerenciamento de estado.
* **[Docling](https://github.com/DS4SD/docling):** Extração de dados e conversão de documentos baseada em IA.
* **[Pandas](https://pandas.pydata.org/):** Manipulação, limpeza e despivotagem (`pd.melt`) dos dados.
* **[PdfPlumber](https://github.com/jsvine/pdfplumber) / PyPDF:** Manipulação de caixas de PDF e renderização de prévias visuais.
* **SQLite3:** Armazenamento persistente estruturado.

## ⚙️ Pré-requisitos e Instalação

Certifique-se de ter o Python 3.9+ instalado. Para instalar todas as dependências necessárias, abra o terminal e execute o comando abaixo:

```bash
pip install streamlit docling pdfplumber pypdf streamlit-cropper pandas openpyxl Pillow



Nota: A primeira execução do Docling fará o download automático dos modelos de IA (Layout e TableStructure). Isso pode levar alguns segundos, mas os arquivos ficarão salvos em cache para as próximas execuções.

🏃 Como Executar

Para iniciar a aplicação, navegue até a pasta do projeto no seu terminal e execute:
Bash

streamlit run app.py

O seu navegador padrão abrirá automaticamente na porta local da aplicação (geralmente http://localhost:8501).
📖 Guia de Uso
Passo 1: Seleção do Modo e Upload

Na barra lateral esquerda, escolha se deseja fazer um ETL Genérico ou extrair a Matriz Orçamentária (Comitês x Programas). Em seguida, faça o upload do arquivo .pdf.
Passo 2: Navegação e Otimização

    Selecione a página desejada. O sistema processa o PDF página por página sob demanda (Lazy Loading) para economizar memória.

    Dicas de Performance:

        Mantenha o modo "Rápido" ativado por padrão.

        Deixe o OCR desligado para relatórios digitais nativos (gerados por Word/Excel). Ative o OCR apenas se o PDF for uma imagem/documento escaneado.

Passo 3: Extração e Tratamento

    Selecione a tabela identificada pelo sistema na lista de opções.

    Aponte qual linha representa os nomes reais das colunas (índice 0, 1, 2...).

    Se a tabela não for encontrada, abra o menu suspensivo de Reforço Manual, desenhe um retângulo em volta da tabela na imagem de referência visual e clique em reprocessar.

Passo 4: Transformação (Modo Comitês)

    Selecione qual coluna atua como eixo Y (Comitês) e quais colunas atuam como eixo X (Programas).

    Selecione o ano de referência.

    Verifique a prévia financeira gerada. Se os dados estiverem corretos, clique em Transformar e salvar. Os dados irão para a memória da aplicação e serão registrados no banco local SQLite.

Passo 5: Exportação

    Revise os lotes processados no painel inferior.

    Ajuste o ano de um lote inteiro se necessário.

    Clique em Baixar Excel Consolidado para gerar o arquivo .xlsx limpo e formatado.