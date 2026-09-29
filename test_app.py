import io

import pandas as pd

import app


def test_parse_raw_normalizes_headers_with_accents_and_spaces():
    raw = pd.DataFrame([
        ["NUMFUNC", "NUMVINC", "Nome Servidor", "Data de Início do Vínculo"],
        ["00123", "7", "Ana", "01/01/2020"],
    ])

    parsed = app.parse_raw(raw)

    assert "NUMFUNC" in parsed.columns
    assert "SERVIDOR" in parsed.columns
    assert "DATA DE INICIO - VINCULO" in parsed.columns


def test_parse_br_number_handles_brazilian_thousands_and_decimal():
    assert app.parse_br_number("1.234,50") == 1234.5
    assert app.parse_br_number("1234.50") == 1234.5
    assert pd.isna(app.parse_br_number("não informado"))


def test_clean_identifies_cpf_and_numfunc_even_with_leading_zeroes():
    raw = pd.DataFrame([
        ["NUMFUNC", "NUMVINC", "SERVIDOR", "CPF", "VINCULO", "DATA DE INICIO - VINCULO"],
        ["01234567890", "1", "Ana", "012.345.678-90", "Efetivo", "01/01/2024"],
    ])

    cleaned = app.clean(app.parse_raw(raw))

    assert cleaned.loc[0, "ALERTA_NUMFUNC_CPF"] == "S"
    assert cleaned.loc[0, "CPF_FORMATADO"] == "012.345.678-90"


def test_expired_contract_uses_strictly_more_than_two_years():
    raw = pd.DataFrame([
        ["NUMFUNC", "NUMVINC", "SERVIDOR", "VINCULO", "DATA DE INICIO - VINCULO"],
        ["1", "1", "No limite", "Temporário", "01/01/2024"],
        ["2", "1", "Vencido", "Temporário", "31/12/2023"],
    ])
    cleaned = app.clean(app.parse_raw(raw))

    result = app.alert_contratos_vencidos(cleaned, pd.Timestamp("2026-01-01"))

    assert result["SERVIDOR"].tolist() == ["Vencido"]


def test_invalid_schema_returns_actionable_message():
    raw = pd.DataFrame([["produto", "preço"], ["caneta", "2,50"]])

    message = app.validate_source_schema(raw)

    assert message is not None
    assert "cabeçalhos" in message
