# PDF → Excel ETL

App em Python (Streamlit) para extrair tabelas de documentos PDF — que mudam de layout ano a ano — e consolidar em uma planilha Excel. Em vez de depender 100% de detecção automática, o app deixa o usuário final conferir visualmente e escolher exatamente qual página, tabela e cabeçalho extrair.

## Funcionalidades

- Upload de PDF com preview de cada página (imagem renderizada)
- Detecção automática de tabelas, com destaque visual (retângulo vermelho) sobre a tabela encontrada
- 3 estratégias de detecção configuráveis (por linhas/bordas, por texto, híbrida) — com fallback automático quando a estratégia atual não encontra nada
- Exclusão de tabelas detectadas antes de escolher qual extrair
- Seleção manual de qual linha é o cabeçalho real (resolve tabelas com cabeçalho mesclado em duas linhas)
- Edição manual dos nomes de coluna, para corrigir qualquer desalinhamento da extração automática
- Renomeação de colunas com memória entre uploads (sugere o mesmo de-para usado antes)
- **Modo especializado "Comitês x Programas"**: despivota automaticamente tabelas em formato matriz (Comitê nas linhas, Programa nas colunas) para formato linear `Comitê | Programa | Ano | Valor`, com exclusão de linhas de total/rodapé e campo Ano editável (que recalcula todas as linhas do lote automaticamente)
- Consolidação de múltiplas tabelas/PDFs na mesma sessão, com exportação para um único Excel (uma aba por tabela + uma aba consolidada)

## Instalação

Requer Python 3.10+.

```bash
git clone <url-deste-repositorio>
cd <pasta-do-repositorio>
pip install -r requirements.txt
```

## Como usar

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

## Limitações conhecidas

- A detecção de tabelas depende da estrutura do PDF. PDFs escaneados (imagem, sem texto selecionável) não são suportados sem OCR — seria necessário adicionar `pytesseract` para esse caso.
- Tabelas com células mescladas de forma complexa podem exigir ajuste manual do cabeçalho e/ou dos nomes de coluna.
- O app roda localmente na máquina de quem usa; para acesso remoto por várias pessoas, seria necessário publicar (ex: Streamlit Community Cloud ou servidor interno).

## Estrutura do projeto

```
.
├── app.py              # App principal (Streamlit)
├── requirements.txt     # Dependências Python
└── README.md
```
