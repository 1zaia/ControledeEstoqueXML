**Controle de Estoque — NFe XML (KG)**

Ferramenta em Python para ler arquivos XML de NFe e gerar relatórios de movimentação e estoque em Excel.

**Funcionalidades**
- Lê arquivos XML de NFe (entrada/saída)
- Extrai código do produto, descrição e quantidade (qCom)
- Classifica movimento como Entrada (+) ou Saída (−)
- Gera relatório diário de movimentos em Excel
- Mantém arquivo de estoque acumulado
- Evita processamento duplicado de notas (por chNFe)
- Armazena XMLs processados e logs de erro

**Pré-requisitos**
- Python 3.8+
- Instalar dependências:

  pip install pandas openpyxl

Tkinter costuma vir com a instalação padrão do Python.

**Como usar**
1. Coloque os arquivos XML nas pastas apropriadas:
   - entrada/ — arquivos de entrada
   - saida/   — arquivos de saída
2. Execute o script principal:

   python ControleEstoqueXML.py

3. Os arquivos processados serão movidos para a pasta de `processados/` e os resultados ficarão em `resultado/`.

**Estrutura de pastas (auto-criada)**
- entrada/ — arquivos XML de entrada
- saida/ — arquivos XML de saída
- processados/
  - entrada/ — XMLs de entrada processados
  - saida/   — XMLs de saída processados
- resultado/ — arquivos Excel gerados
- erros/ — logs de erros

**Saídas geradas**
- `resultado/movimento_TIMESTAMP.xlsx` — relatório de movimentação do dia
- `resultado/estoque_geral.xlsx` — saldo acumulado por produto
- `resultado/processadas.xlsx` — registro de chNFe processadas (para idempotência)

Colunas típicas do relatório de movimento:
- Data — data do processamento
- Tipo — Entrada / Saída
- Codigo — código do produto
- Produto — descrição
- Quantidade_KG — quantidade com sinal (+/-)
- chNFe — chave da nota fiscal
- Arquivo_XML — arquivo-fonte

**Regras de processamento**
- Para cada XML lido: extrai `chNFe`. Se já estiver em `processadas`, o arquivo é ignorado.
- Percorre os itens em `det/prod` e extrai `cProd`, `xProd` e `qCom`.
- Aplica fator de movimento: Entrada = +1, Saída = −1. Quantidade final = qCom × fator.

**Tratamento de erros**
- Se um arquivo falhar ao ser parseado ou estiver faltando tags, o processamento continua.
- É criado um log em `erros/` com o nome do arquivo e a mensagem de exceção.

**Observações**
- Garanta que os XMLs estejam bem formados e na codificação correta.
- Ajuste unidades e conversões se necessário (este projeto considera KG por padrão).

Se quiser, eu adapto o texto (ex.: inglês, instruções de instalação com `requirements.txt`, ou exemplos de execução). 