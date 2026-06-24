import pandas as pd
from datetime import datetime, timedelta

def categorize(desc):
    d = str(desc).upper()
    if 'GOFO' in d: return ('Ingresos GOFO', 'ingreso')
    if 'INTUIT' in d and 'PAYROLL' in d: return ('Intuit / Reembolso', 'ingreso')
    if 'WALLE CARGO' in d or 'BLUEBOX' in d: return ('Otros ingresos', 'ingreso')
    if 'ZELLE FROM' in d: return ('Otros ingresos', 'ingreso')
    if 'BUSINESS TO BUSINESS ACH PAYROLL' in d: return ('Payroll Drivers', 'gasto')
    if 'ZELLE TO' in d and any(x in d for x in ['PAGO SAT','SAT 5','SAT5','FINAL PAY SAT','PAYMENT SAT','SHORT']): return ('Payroll Drivers', 'gasto')
    if 'ZELLE TO' in d and 'DIEGO BRITO' in d: return ('Empleado Diego Brito', 'gasto')
    if 'ZELLE TO' in d and 'ELDAR' in d: return ('Renta / Landlord', 'gasto')
    if 'ZELLE TO' in d and 'MOVING' in d: return ('Mudanza / Moving', 'gasto')
    if 'ZELLE TO' in d and ('CARRO' in d or 'MANAGER' in d): return ('Gastos Manager', 'gasto')
    if 'ZELLE TO' in d and 'ALISON' in d: return ('Propiedades', 'gasto')
    if 'ZELLE TO' in d and 'GERMAIN' in d: return ('Gastos Manager / SHV', 'gasto')
    if 'ZELLE TO' in d and 'SUSANA ROMERO' in d: return ('Payroll Drivers', 'gasto')
    if 'ZELLE TO' in d and any(x in d for x in ['ARCHILL','TOW']): return ('Operaciones / Grúa', 'gasto')
    if 'ZELLE TO' in d and 'MOISES' in d: return ('Vehículos / Operaciones', 'gasto')
    if 'ZELLE TO' in d and 'DAGO' in d: return ('Mantenimiento', 'gasto')
    if 'ZELLE TO' in d and 'ANGLADA' in d: return ('Personal / Familia', 'gasto')
    if 'ZELLE TO' in d and 'SPEEDCO' in d: return ('Operaciones / Speedco', 'gasto')
    if 'ZELLE TO' in d and 'MP TEAM' in d: return ('Operaciones / MP Team', 'gasto')
    if 'ZELLE TO' in d and any(x in d for x in ['YESENIA','MARVIN']): return ('Payroll Drivers', 'gasto')
    if 'ZELLE TO' in d and any(x in d for x in ['CHANTEL','AKACIA','LINIECE','MILLIANNY','YALIMAR',
        'KEVIN','JOSH WELLS','DOUGLAS','MIGUEL','PEDRO','JARI','JAMIE','ANTONIO','EDER',
        'JEISON','RUTA PAGO','RIGMARY','ARELI','ALBERTO','ANDREA TOWERS']): return ('Payroll Drivers', 'gasto')
    if 'ZELLE TO' in d and 'WALLIE' in d: return ('Operaciones / Subcontrato', 'gasto')
    if 'ZELLE TO' in d and any(x in d for x in ['HOTEL','OIL','PAGO GERMA','SHREVEPORT']): return ('Gastos Manager / SHV', 'gasto')
    if 'ZELLE TO' in d: return ('Otros Zelle', 'gasto')
    if 'ONLINE TRANSFER TO ECHEVERRY' in d: return ('Transferencia Personal', 'gasto')
    if 'TRANSACTIONS FEE' in d or 'SERVICE FEE' in d: return ('Fees Bancarios', 'gasto')
    if 'RECURRING PAYMENT' in d and 'EXTRA SP' in d: return ('Extra Space Storage', 'gasto')
    if 'RECURRING PAYMENT' in d: return ('Pagos Recurrentes', 'gasto')
    if 'PURCHASE' in d and any(x in d for x in ['RACETRAC','SHELL','GAS']): return ('Gasolina', 'gasto')
    if 'PURCHASE' in d and 'LULULEMO' in d: return ('Personal / Ropa', 'gasto')
    if 'PURCHASE' in d and 'RC BOOKI' in d: return ('Viajes / Hotel', 'gasto')
    if 'PURCHASE' in d and 'HDN' in d: return ('Abogados', 'gasto')
    if 'PURCHASE' in d and 'FAITH AU' in d: return ('Vehículos / Auto', 'gasto')
    if 'PURCHASE' in d and 'DEEP SOU' in d: return ('Comida / Restaurantes', 'gasto')
    if 'PURCHASE' in d and 'EXTRA SP' in d: return ('Extra Space Storage', 'gasto')
    if 'PURCHASE' in d: return ('Compras / Tarjeta', 'gasto')
    return ('Otros', 'gasto')

def analyze_csv(filepath):
    df = pd.read_csv(filepath)
    df.columns = ['DATE','DESCRIPTION','AMOUNT','CHECK','STATUS']
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'], errors='coerce')
    df['DATE']   = pd.to_datetime(df['DATE'])
    df = df[df['AMOUNT'].notna()].copy()

    df[['CATEGORY','TYPE']] = df['DESCRIPTION'].apply(lambda x: pd.Series(categorize(x)))
    df['WEEK_START'] = df['DATE'].dt.to_period('W').apply(lambda x: x.start_time)
    df['WEEK_END']   = df['DATE'].dt.to_period('W').apply(lambda x: x.end_time)

    weeks = sorted(df['WEEK_START'].unique(), reverse=True)

    weekly_summary = []
    for ws in weeks:
        wdf = df[df['WEEK_START'] == ws]
        ingresos = round(wdf[wdf['TYPE']=='ingreso']['AMOUNT'].sum(), 2)
        gastos   = round(wdf[wdf['TYPE']=='gasto']['AMOUNT'].sum(), 2)
        neto     = round(ingresos + gastos, 2)
        gofo_ing = round(wdf[wdf['CATEGORY']=='Ingresos GOFO']['AMOUNT'].sum(), 2)
        payroll  = round(wdf[wdf['CATEGORY']=='Payroll Drivers']['AMOUNT'].sum(), 2)

        # Category breakdown
        by_cat = wdf.groupby(['CATEGORY','TYPE'])['AMOUNT'].sum().round(2).to_dict()
        cats = {}
        for (cat, typ), amt in by_cat.items():
            cats[cat] = {'amount': amt, 'type': typ}

        weekly_summary.append({
            'week_start': ws.strftime('%m/%d/%Y'),
            'week_end':   wdf['WEEK_END'].iloc[0].strftime('%m/%d/%Y'),
            'week_label': f"{ws.strftime('%m/%d')} – {wdf['WEEK_END'].iloc[0].strftime('%m/%d')}",
            'ingresos':   ingresos,
            'gastos':     abs(gastos),
            'neto':       neto,
            'gofo':       gofo_ing,
            'payroll':    abs(payroll),
            'categories': cats,
            'transactions': wdf[['DATE','DESCRIPTION','AMOUNT','CATEGORY','TYPE']].to_dict('records'),
        })

    # Overall stats
    total_ingresos = round(df[df['TYPE']=='ingreso']['AMOUNT'].sum(), 2)
    total_gastos   = round(abs(df[df['TYPE']=='gasto']['AMOUNT'].sum()), 2)
    avg_gofo_weekly = round(df[df['CATEGORY']=='Ingresos GOFO']['AMOUNT'].sum() / max(len(weeks),1), 2)

    return {
        'weeks': weekly_summary,
        'total_ingresos': total_ingresos,
        'total_gastos': total_gastos,
        'neto_total': round(total_ingresos - total_gastos, 2),
        'avg_gofo_weekly': avg_gofo_weekly,
        'date_range': f"{df['DATE'].min().strftime('%m/%d/%Y')} – {df['DATE'].max().strftime('%m/%d/%Y')}",
        'num_weeks': len(weeks),
    }
