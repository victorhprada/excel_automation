# Ferramenta de Validação de Faturamento Excel

Uma aplicação Streamlit para validação de faturamento através da comparação de arquivos Excel.

## Descrição

Esta ferramenta permite:
- Selecionar o período de análise (mês e ano)
- Fazer upload de arquivos Excel (PARCEIRO e BASE)
- Processar e validar dados de faturamento
- Visualizar dados carregados

## Instalação

1. Clone o repositório
2. Instale as dependências:

```bash
pip install -r requirements.txt
```

## Como Executar

Execute o comando:

```bash
streamlit run app.py
```

A aplicação abrirá automaticamente no seu navegador em `http://localhost:8501`

## Uso

1. **Selecionar Período**: Na sidebar, escolha o mês e ano desejados
2. **Upload de Arquivos**:
   - Arquivo PARCEIRO: arquivo `.xlsx` com dados do parceiro
   - Arquivo BASE: arquivo `.xlsx` ou `.xlsm` com dados base (preserva fórmulas)
3. **Processar**: Clique em "Iniciar Processamento" para carregar os dados
4. **Visualizar**: Veja o preview dos dados carregados

## Estrutura dos Arquivos

- `app.py` - Aplicação principal Streamlit
- `requirements.txt` - Dependências do projeto
- `README.md` - Esta documentação

## Tecnologias

- **Streamlit**: Framework web para interface
- **Pandas**: Manipulação de dados
- **Openpyxl**: Leitura de arquivos Excel com preservação de fórmulas
