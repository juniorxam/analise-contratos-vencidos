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
