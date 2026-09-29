# Análise de Contratos Vencidos

Aplicativo Streamlit para analisar:

- contratos temporários vencidos, considerando mais de dois anos desde o início do vínculo;
- registros em que `NUMFUNC` é igual ao `CPF`, ignorando formatação e zeros à esquerda.

## Como executar

```bash
pip install -r requirements.txt
streamlit run app.py
```

A planilha pode ser enviada nos formatos CSV, XLS, XLSX ou XLSM. A data de referência dos contratos vencidos pode ser ajustada na barra lateral.

## Melhorias incluídas

- Reconhecimento de cabeçalhos com diferenças de acentuação, caixa, espaços e nomes comuns de exportação.
- Validação da estrutura do arquivo com mensagens orientando como corrigir a planilha.
- Conversão numérica compatível com valores brasileiros, como `1.234,50`, sem distorcer milhares.
- Aviso explícito para registros sem data de início válida, que não podem ser classificados como vencidos.
- Lógica de negócio separada da interface para facilitar manutenção e testes automatizados.

## Testes

```bash
pytest -q
```
