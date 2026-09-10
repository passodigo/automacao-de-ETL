# PDF → Excel ETL

Ferramenta em Python para extrair tabelas de documentos PDF (que mudam de layout ano a ano) e consolidar em uma planilha Excel — com interface visual para conferir e escolher exatamente qual tabela/página extrair, em vez de depender 100% de detecção automática.

Tem dois módulos:

| Arquivo | O que é | Quando usar |
|---|---|---|
| `app.py` | App visual (Streamlit): upload de PDF, preview de página, seleção manual de tabela/cabeçalho, exportação para Excel | Uso interativo, dia a dia, quando o layout muda ou não é confiável o suficiente para automação 100% |
| `etl_pipeline.py` | Script batch (linha de comando): processa uma pasta inteira de PDFs de uma vez, usando fuzzy matching para achar a tabela certa e mapear colunas | Quando o layout já está mapeado/estável e você quer rodar em lote, sem interface |

## Índice

- [Funcionalidades](#funcionalidades)
- [Instalação](#instalação)
- [Como usar o app visual (`app.py`)](#como-usar-o-app-visual-apppy)
- [Como usar o script em lote (`etl_pipeline.py`)](#como-usar-o-script-em-lote-etl_pipelinepy)
- [Limitações conhecidas](#limitações-conhecidas)
- [Estrutura do projeto](#estrutura-do-projeto)

## Funcionalidades

**App visual (`app.py`)**
- Upload de PDF com preview de cada página (imagem renderizada)
- Detecção automática de tabelas, com destaque visual (retângulo vermelho) sobre a tabela encontrada
- 3 estratégias de detecção configuráveis (por linhas/bordas, por texto, híbrida) — com fallback automático quando a estratégia atual não encontra nada
- Exclusão de tabelas detectadas antes de escolher qual extrair
- Seleção manual de qual linha é o cabeçalho real (resolve tabelas com cabeçalho mesclado em duas linhas)
- Edição manual dos nomes de coluna, para corrigir qualquer desalinhamento da extração automática
- Renomeação de colunas com memória entre uploads (sugere o mesmo de-para usado antes)
- **Modo especializado "Comitês x Programas"**: despivota automaticamente tabelas em formato matriz (Comitê nas linhas, Programa nas colunas) para formato linear `Comitê | Programa | Ano | Valor`, com exclusão de linhas de total/rodapé e campo Ano editável (que recalcula todas as linhas do lote automaticamente)
- Consolidação de múltiplas tabelas/PDFs na mesma sessão, com exportação para um único Excel (uma aba por tabela + uma aba consolidada)

**Script em lote (`etl_pipeline.py`)**
- Varre uma pasta de PDFs e identifica a tabela certa em cada um via fuzzy matching no título (mesmo que o título mude de ano para ano)
- Normaliza nomes de coluna variáveis para um schema canônico configurável, com fallback por fuzzy matching
- Registra tudo que não foi reconhecido com confiança em uma aba separada "Revisar", em vez de descartar silenciosamente
- Gera um Excel consolidado com abas "Dados", "Revisar" e "Log"

## Instalação

Requer Python 3.10+.

```bash
git clone <url-deste-repositorio>
cd <pasta-do-repositorio>
pip install -r requirements.txt
```

## Como usar o app visual (`app.py`)

```bash
streamlit run app.py
```

Isso abre automaticamente `http://localhost:8501` no navegador.

**Fluxo:**

1. Escolha o modo na barra lateral: **ETL Genérico** ou **Comitês x Programas**
2. Faça upload do PDF
3. Navegue até a página com a tabela desejada
4. Se nenhuma tabela for detectada, troque a **estratégia de detecção** (barra lateral) — tabelas sem bordas desenhadas costumam precisar da estratégia "Baseada em texto"
5. Marque/desmarque quais tabelas detectadas você quer considerar
6. Confira a prévia bruta e informe qual linha é o cabeçalho real (importante em tabelas com cabeçalho mesclado em duas linhas)
7. Ajuste os nomes de coluna se necessário
8. **Modo Genérico:** renomeie colunas à vontade e clique em "Adicionar ao conjunto final"
   **Modo Comitês x Programas:** selecione a coluna de Comitê, quais colunas são Programas, quais linhas incluir, informe o Ano, e clique em "Transformar e adicionar"
9. Repita para outras páginas/PDFs quantas vezes precisar
10. Baixe o Excel consolidado na seção "Conjunto final"

## Como usar o script em lote (`etl_pipeline.py`)

1. Coloque os PDFs em uma pasta (`./pdfs_entrada` por padrão)
2. Edite as configurações no topo do arquivo:
   - `TABLE_TITLE_KEYWORDS`: termos que identificam a tabela certa (ex: `"receita por municipio"`)
   - `COLUMN_ALIASES`: para cada coluna final, liste as variações de nome já vistas nos documentos
   - `FUZZY_THRESHOLD_TITLE` / `FUZZY_THRESHOLD_COLUMN`: sensibilidade do reconhecimento (0–100)
3. Rode:
   ```bash
   python etl_pipeline.py
   ```
4. O Excel gerado (`saida_consolidada.xlsx` por padrão) terá:
   - **Dados**: tudo consolidado, com colunas padronizadas
   - **Revisar**: tabelas ou colunas que não foram reconhecidas com confiança
   - **Log**: resumo da execução

## Limitações conhecidas

- A detecção de tabelas depende da estrutura do PDF. PDFs escaneados (imagem, sem texto selecionável) não são suportados sem OCR — seria necessário adicionar `pytesseract` para esse caso.
- Tabelas com células mescladas de forma complexa podem exigir ajuste manual do cabeçalho e/ou dos nomes de coluna no app visual.
- O app visual (`app.py`) roda localmente na máquina de quem usa; para acesso remoto por várias pessoas, seria necessário publicar (ex: Streamlit Community Cloud ou servidor interno).
- O `etl_pipeline.py` assume que existe exatamente uma tabela "certa" por PDF; para PDFs com múltiplas tabelas candidatas ambíguas, prefira o app visual.

## Estrutura do projeto

```
.
├── app.py              # App visual (Streamlit)
├── etl_pipeline.py      # Script de ETL em lote
├── requirements.txt      # Dependências Python
└── README.md
```
