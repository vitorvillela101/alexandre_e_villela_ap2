import requests
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path



# Configurações
token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0b2tlbl90eXBlIjoiYWNjZXNzIiwiZXhwIjoxNzQ3OTEyODc0LCJpYXQiOjE3NDUzMjA4NzQsImp0aSI6ImQxM2MxOWNlMDA3YzRjMTRiZDRmYjkxZGE2MGQ4N2RmIiwidXNlcl9pZCI6NjV9.1C50rycumEgTVv3eTK6qUj99vUqvxvtZVyO-wLPOVoc"
headers = {"Authorization": f"JWT {token}"}
url = "https://laboratoriodefinancas.com/api/v1/balanco"

# Pasta de saída: Área de Trabalho com timestamp
output_dir = Path.home() / 'Desktop'
output_file = output_dir / f"comparativo_indicadores_{datetime.now():%Y%m%d_%H%M}.xlsx"

# Lista de empresas e período
tickers = ["JBSS3", "MRFG3", "BRFS3", "BEEF3"]
periodo = "20244T"

# Função para buscar balanço
def fetch_balanco(ticker, ano_tri):
    resp = requests.get(url, headers=headers, params={"ticker": ticker, "ano_tri": ano_tri})
    resp.raise_for_status()
    dados = resp.json().get("dados", [])
    if not dados:
        raise ValueError(f"Nenhum dado para {ticker} no período {ano_tri}")
    return pd.DataFrame(dados[0]["balanco"])

# Função flexível de extração
def get_valor(df, *keys):
    for key in keys:
        mask = df['descricao'].str.contains(key, case=False, na=False)
        if mask.any():
            return df.loc[mask, 'valor'].iloc[0]
    return np.nan

# Cálculo seguro
def safe_ratio(num, den, dias=False):
    if pd.isna(num) or pd.isna(den) or den == 0:
        return np.nan
    r = num / den
    return r * 360 if dias else r

# Calcula indicadores a partir do balanço
# EVA = NOPAT - (WACC * Capital Investido)
# onde usamos Capital Investido = Ativo Total (AT)

def calcula_indicadores(df):
    # Extrair contas
    AC = get_valor(df, "Ativo Circulante")
    PC = get_valor(df, "Passivo Circulante")
    Estoque = get_valor(df, "Estoque", "Estoques")
    DA = get_valor(df, "Despesas Antecipadas")
    Disponivel = get_valor(df, "Disponibilidades", "Caixa")
    Aplic = get_valor(df, "Aplicações")
    ARLP = get_valor(df, "Ativo Realizável a Longo Prazo")
    PNC = get_valor(df, "Passivo Não Circulante")
    Clientes = get_valor(df, "Clientes", "Contas a receber")
    Fornecedores = get_valor(df, "Fornecedores")
    Receita = get_valor(df, "Receita Líquida", "Receita")
    CMV = get_valor(df, "Custo das Mercadorias Vendidas", "CMV")
    Compras = get_valor(df, "Compras")
    Emprest = get_valor(df, "Empréstimos", "Financiamentos")
    AT = get_valor(df, "Ativo Total")
    PT = get_valor(df, "Passivo Total")
    PL = get_valor(df, "Patrimônio Líquido")
    AP = get_valor(df, "Ativo Permanente")
    DF = get_valor(df, "Despesa Financeira Líquida", "Despesas Financeiras")
    BT = get_valor(df, "Benefício Tributário da Dívida", "BT Dívida")
    IR = get_valor(df, "IR Corrente")
    LAIR = get_valor(df, "LAIR")  # Lucro Antes IR/CSLL

    # Indicadores de liquidez
    CCL = AC - PC
    LC = safe_ratio(AC, PC)
    LS = safe_ratio(AC - Estoque - DA, PC)
    LI = safe_ratio(Disponivel, PC)
    LG = safe_ratio(AC + ARLP, PC + PNC)

    # Ciclos
    PME = safe_ratio(Estoque, CMV, dias=True)
    GE = safe_ratio(360, PME)
    PMRV = safe_ratio(Clientes, Receita, dias=True)
    PMPF = safe_ratio(Fornecedores, Compras, dias=True)
    CO = PMRV + PME
    CF = CO - PMPF

    # Capital de Giro
    CE = PME
    ACO = AC - Disponivel - Aplic
    PCO = PC - Emprest
    NCG = ACO - PCO
    ACF = Disponivel + Aplic
    PCF = Emprest
    ST = ACF - PCF

    # Estrutura de capital
    Rel_Capitais = safe_ratio(PT, PL)
    Endiv_Geral = safe_ratio(PT, PT + PL)
    Solv = safe_ratio(AT, PT)
    Comp_Endiv = safe_ratio(PC, PT)
    Imob_PL = safe_ratio(AP, PL)

    # Custo da dívida líquida de impostos
    Ki = safe_ratio(DF, PC + PCO)
    DFL = DF - BT
    Alq_IR_CSLL = safe_ratio(IR, LAIR)

    # Pesos e WACC
    Wi = safe_ratio(PT, (PT + PL))
    We = 1 - Wi
    Ke = Alq_IR_CSLL  # suposição simplificada
    CMPC = Wi * Ki + We * Ke

    # Resultado Operacional e retorno
    EBITDA = safe_ratio(get_valor(df, "Margem EBITDA") * Receita, 1)
    EBIT = EBITDA - get_valor(df, "Depreciação", "Amortização")
    NOPAT = EBIT - IR  # Lucro após IR
    ROI = safe_ratio(NOPAT, (PT + PL))
    ROE = safe_ratio(get_valor(df, "Lucro Líquido"), PL)
    GAF = safe_ratio(ROE, ROI)

    # Cálculo do EVA
    EVA = NOPAT - (CMPC * AT)

    return {
        'CCL': CCL, 'LC': LC, 'LS': LS, 'LI': LI, 'LG': LG,
        'PME': PME, 'GE': GE, 'PMRV': PMRV, 'PMPF': PMPF,
        'CO': CO, 'CF': CF, 'CE': CE,
        'ACO': ACO, 'PCO': PCO, 'NCG': NCG,
        'ACF': ACF, 'PCF': PCF, 'ST': ST,
        'Rel_Capitais': Rel_Capitais, 'Endiv_Geral': Endiv_Geral,
        'Solv': Solv, 'Comp_Endiv': Comp_Endiv, 'Imob_PL': Imob_PL,
        'Ki': Ki, 'DFL': DFL, 'BT': BT, 'Alq_IR_CSLL': Alq_IR_CSLL,
        'Wi': Wi, 'We': We, 'CMPC': CMPC,
        'EBITDA': EBITDA, 'EBIT': EBIT, 'NOPAT': NOPAT,
        'ROI': ROI, 'ROE': ROE, 'GAF': GAF,
        'EVA': EVA
    }


# Processamento
df_list = []
for t in tickers:
    try:
        df_bal = fetch_balanco(t, periodo)
        ind = calcula_indicadores(df_bal)
        ind['Ticker'] = t
        df_list.append(ind)
    except Exception as e:
        print(f"Erro em {t}: {e}")

# Exportação com formatação
df_res = pd.DataFrame(df_list)
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df_res.to_excel(writer, sheet_name='Indicadores', index=False)
    wb = writer.book
    ws = writer.sheets['Indicadores']
    fmt_header = wb.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1})
    for col_num, value in enumerate(df_res.columns.values):
        ws.write(0, col_num, value, fmt_header)
        ws.set_column(col_num, col_num, 15)

print(f"Comparativo salvo em: {output_file}")
print(df_res.shape, "- Tickers processados:", [d['Ticker'] for d in df_list])
