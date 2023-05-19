from tkinter import N
import pandas as pd
import win32com.client as win32
from datetime import date
from pandas.tseries.offsets import BDay
import numpy as np
import time

import warnings
warnings.filterwarnings("ignore")

dia = date.today().day
mes = date.today().month 
ano = date.today().year 

if mes < 10:
    mes = "0"+ str(mes)

if dia < 10:
    dia = "0"+str(dia)

ano = str(ano)
dia = str(dia)
mes = str(mes)

trades =  pd.read_excel(r'G:\SPX_FTP\files\itau\files\Operações Realizadas SPX_{d}{m}{a}.xlsx'.format(d = dia, m = mes, a = ano), skiprows = 4)
trades.loc[trades['Mercado'] == 'OddLot', 'Ativo'] = trades['Ativo'].str[:-1]

trades_amount = trades.groupby(['Ativo','Natureza da Operação','Corretora']).aggregate({'Quantidade Negociada': 'sum'})
trades_amount.name = 'Amount'

# Calcula preço médio

def wavg(df, values, weights):
    try:
        value = df[values]
        weight = df[weights]
        return (value * weight).sum() / weight.sum()
    except ZeroDivisionError:
        return 0

trades_price = trades.groupby(['Ativo','Natureza da Operação','Corretora']).apply(wavg, 'Valor Negociado', 'Quantidade Negociada')
trades_price.name = 'Price'

trades_final = trades_amount.merge(trades_price, on = ['Ativo','Corretora','Natureza da Operação'], how = "left")
trades_final.reset_index([0,1,2], inplace = True)

trades_final.columns = ['Product','Dealer','Side', 'Amount', 'Price']

depara_b3 = pd.read_excel('G:\Batimento BackOffice\Renda Variável\Batimento B3\De_Para_B3(Pythons).xlsx', sheet_name= 'Geral')
corretoras = pd.DataFrame(depara_b3)
corretoras = corretoras[['corretoras_itau', 'corretoras_lote45']].dropna()
corretoras = corretoras.set_index('corretoras_itau').T.to_dict('list')

trades_final = trades_final.replace({"Dealer": corretoras})

products = depara_b3['produtos'].dropna()
products = products.to_list()
funds = depara_b3['funds'].dropna()
funds = funds.to_list()

columns_new = ["Trading Desk","Product","Amount","Price","Dealer","SettleDealer", "Trader", "PositionType", "ExecutionType", "Currency"]
lote45 = ['LOTE45']

start_time = time.time()      

aloc_general = pd.read_json("http://10.21.2.23:3000/api/trades/setfunds//date/{a}-{m}-{d}".format(a = ano, m = mes, d = dia))
aloc_general = aloc_general[aloc_general['IsCashFlowTrade'] == False]
aloc_general = aloc_general[aloc_general['IsReplicatedTrade'] == False]

aloc_general["Trade Date"] = pd.to_datetime(aloc_general["Trade Date"])
aloc_general["Effective"] = pd.to_datetime(aloc_general["Effective"])
                                                                                   
aloc_equity_first = aloc_general.loc[(aloc_general["ProductClass"].isin(products)) & (aloc_general["Trading Desk"].isin(funds)) & (~aloc_general["Dealer"].isin(lote45))]

aloc_equity = aloc_equity_first.copy()
aloc_equity = aloc_equity_first.reset_index(drop = True)
columns_old = aloc_equity.columns.tolist()
aloc_equity.loc[aloc_equity['Product'] == 'BPAN12_22', 'Product'] = 'BPAN12'
aloc_equity.loc[aloc_equity['Product'] == 'RRRP1_23', 'Product'] = 'RRRP1'
aloc_equity.loc[aloc_equity['Product'] == 'SEQL1_23', 'Product'] = 'SEQL1'

traders_lucas = depara_b3['trader_FIAS'].dropna()
traders_lucas = traders_lucas.to_list()
traders_thiago = depara_b3['trader_multiasset'].dropna()
traders_thiago = traders_thiago.to_list()
traders_matheus = depara_b3['trader_macros'].dropna()
traders_matheus = traders_matheus.to_list()

aloc_equity.loc[aloc_equity['Trader'].isin(traders_lucas), 'Trader'] = 'Lucas Rossi'
aloc_equity.loc[aloc_equity['Trader'].isin(traders_thiago), 'Trader'] = 'Thiago Santos'
aloc_equity.loc[aloc_equity['Trader'].isin(traders_matheus), 'Trader'] = 'Equities Brasil'

options = depara_b3['opções'].dropna()
options = options.to_list()

aloc_equity.loc[aloc_equity['ProductClass'].isin(options), 'Product'] = aloc_equity['Product'].str[:-3]
                
for element in columns_old:
    if element not in columns_new:
        aloc_equity.drop(columns = element, inplace = True)

aloc_equity.loc[aloc_equity['Amount'] < 0, 'PositionType'] = 'V'
aloc_equity.loc[aloc_equity['Amount'] > 0, 'PositionType'] = 'C'

aloc_equity.fillna(value = "", inplace = True)

print("--- %s seconds ---" % (time.time() - start_time))

try:

    lote_amount = aloc_equity.groupby(['Product','Dealer','PositionType','Trader']).aggregate({'Amount': 'sum' })

    lote_amount.name = "Amount"

    lote_price = aloc_equity.groupby(['Product','Dealer','PositionType', 'Trader']).apply(wavg, 'Price','Amount')

    lote_price.name = "Price"

    lote = lote_amount.merge(lote_price, on = ['Product','Dealer','PositionType', 'Trader'], how = "left")

    lote.reset_index([0,1,2,3], inplace = True)

    lote['Amount'] = lote['Amount'].abs()
    #lote['Amount'] = lote['Amount'].astype('Int64')

    lote.fillna('', inplace = True)

    lote.columns = ['Product','Dealer','Side','Trader','Amount','Price']
except NameError:
    pass

#VALIDAÇÃO

validation = lote.merge(trades_final, on = ['Product','Dealer','Side'], how = 'outer',suffixes=('', '_Recaps'))

validation.loc[validation['Amount'] == validation['Amount_Recaps'], 'Quantity OK?'] = 'OK'

validation.loc[abs(validation['Price'] - validation['Price_Recaps']) <= 0.001, 'Price OK?'] = 'OK'

validation.loc[(validation['Quantity OK?'] == 'OK') & (validation['Price OK?'] == 'OK'), 'Trade OK?'] = 'OK'

validation.fillna(0, inplace = True)

validation.loc[validation['Quantity OK?'] == 0, 'Quantity OK?'] = ''
validation.loc[validation['Price OK?'] == 0, 'Price OK?'] = ''
validation.loc[validation['Trader'] == 0, 'Trader'] = ''
validation.loc[validation['Trade OK?'] == 0, 'Trade OK?'] = ''

validation['Amount_Recaps'] = validation['Amount_Recaps'].astype('int64') 

labels_valid =  ['Product', 'Dealer', 'Side', 'Trader', 'Amount', 'Amount_Recaps', 'Price', 'Price_Recaps', 'Quantity OK?', 'Price OK?', 'Trade OK?']

validation = validation[labels_valid]

#DIVERGENCIAS

anterior = date.today()- BDay(1) #Mudar para BDay(1) quando for rodar pra valer
anterior.to_pydatetime()

dia_anterior = anterior.day
mes_anterior = anterior.month
ano_anterior = anterior.year

meses_int = [1,2,3,4,5,6,7,8,9,10,11,12]
meses_str = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

for i in range(len(meses_int)):
    if meses_int[i] == mes_anterior:
        mes_anterior = meses_str[i]

if dia_anterior < 10:
    dia_anterior = "0"+str(dia_anterior)
    
pos_anterior = pd.read_csv(r'G:\Carteiras Oficiais\Lote45\AllTradingDesksOverView{d}{m}{a}.txt'.format(d = dia_anterior, m = mes_anterior, a = ano_anterior), sep = '	')

prod_class = depara_b3['produtos'].dropna()
prod_class = prod_class.to_list()
pos_anterior =pos_anterior.loc[pos_anterior['ProductClass'].isin(prod_class)]
pos_books_b3 =pos_anterior.groupby(['Fund','Product','Book']).aggregate({'Amount': 'sum'})
pos_books_b3.reset_index([0,1,2], inplace = True)
pos_books_b3.columns = ['Trading Desk','Product','Book','Amount']

#trades = aloc_equity_first
trades = aloc_equity

thiago_lote = trades[trades['Trader'] == 'Thiago Santos']
matheus_lote = trades[trades['Trader'] == 'Equities Brasil']
lucas_lote = trades[trades['Trader'] == 'Lucas Rossi']

lote__ = aloc_equity_first
lote___amount = lote__.groupby(['Trading Desk', 'Product', 'Book']).aggregate({'Amount': 'sum' })
lote___amount.reset_index([0,1,2], inplace = True)
aloc = pd.concat([pos_books_b3, lote___amount]).groupby(["Trading Desk", "Product", "Book"], as_index=False)["Amount"].sum()

position = aloc.loc[aloc['Product'].isin(lote___amount['Product'])]
position['Product'] = position['Product'].replace(['RRRP1_23'],'RRRP1')
position['Product'] = position['Product'].replace(['SEQL1_23'],'SEQL1')
position.reset_index(drop = True, inplace=True)

html_rateio = """             
Favor verificar o rateio abaixo.
</span>
<br>
<br>
<div>
"""
space = """
</span>              
<br>
<br>
"""

def color_negative_red(value):
    try:      
        if int(value) < 0:
            color = 'red'
        elif int(value) > 0:
            color = 'green'
        else:
            color = 'black'
        
        return 'color: %s' % color
    except:
        pass
        
th_props = [
('font-size', '14px'),
('text-align', 'center'),
('font-weight', 'bold'),
('color', '##f7f7f9'),
('background-color', '#b0c5d6')
]

td_props = [
('font-size', '14px')
]

styles = [
dict(selector="th", props=th_props),
dict(selector="td", props=td_props)
]

today_rateio = date.today()

day_rateio = today_rateio.day
month_rateio = today_rateio.month 
year_rateio = today_rateio.year  

if month_rateio < 10:
    month_rateio = "0"+str(month_rateio)

if day_rateio < 10:
    day_rateio = "0"+str(day_rateio)

rateio_hoje = pd.read_excel(r'M:\Operation\Administracao Fundos\BTG Pactual\Carteira\PL Abertura_Rateio\Anexo\{a}{m}{d}_NAV_Funds%_Redemptions-N_R_E_C_L.xlsx'.format(a = year_rateio, m = month_rateio, d = day_rateio), skiprows = 19)

try:
    rateio_hoje.drop(['Unnamed: 2','Unnamed: 3'], axis = 1, inplace = True)
except:
    rateio_hoje.drop(['Unnamed: 2'], axis = 1, inplace = True)
    
rateio_hoje.drop(rateio_hoje.index[5:139], inplace = True)

macro = pd.DataFrame(depara_b3)
macro = macro[['fund_rateio', 'funds']].dropna()
macro = macro.set_index('fund_rateio').T.to_dict('list')

rateio_hoje = rateio_hoje.replace({"FUND": macro})

rateio_hoje.columns = ['Fund', 'Allocation']

lucas_books = depara_b3['books_FIAs'].dropna()
lucas_books = lucas_books.to_list()
thiago_books = depara_b3['books_multiasset'].dropna()
thiago_books = thiago_books.to_list()
matheus_books = depara_b3['books_macro'].dropna()
matheus_books = matheus_books.to_list()

####MARRETA BRABA PARA CONSTRUIR RATEIO DOS MACROS

position_1 = position.copy()
position_2 =  position_1.copy()

voando = validation.loc[validation['Trader'] == '']

voando = voando.copy()

voando['Amount'] = voando['Amount'].map('{:,.0f}'.format)
voando['Amount_Recaps'] = voando['Amount_Recaps'].map('{:,.0f}'.format)

voando['Price'] = 'R$ ' + voando['Price'].map('{:,.5f}'.format)
voando['Price_Recaps'] = 'R$ ' + voando['Price_Recaps'].map('{:,.5f}'.format)

voando_html  = voando.style.apply(lambda x: ['background: #bcffb8' if x == 'OK' else 'background: #ffb8b8' for x in voando['Trade OK?']], axis = 0).hide_index().render()   

###########################################
try:
    rateio_lucas = position.loc[position['Product'].isin(lucas_lote['Product']) & (position['Trading Desk'].isin(lucas_books))]
    new_lines = [
    {"Trading Desk": "SPX FALCON", "Product": "", "Book": "Long_Bancos", "Amount": 0},
    {"Trading Desk": "SPX FALCON INSTITUCIONAL", "Product": "", "Book": "Long_Bancos", "Amount": 0},
    {"Trading Desk": "SPX HORNET", "Product": "", "Book": "Long_Bancos", "Amount": 0},
    {"Trading Desk": "SPX APACHE", "Product": "", "Book": "Long_Bancos", "Amount": 0},
    {"Trading Desk": "SPX LONG BIAS", "Product": "", "Book": "Long_Bancos", "Amount": 0},
    {"Trading Desk": "SPX PATRIOT", "Product": "", "Book": "Long_Bancos", "Amount": 0},
    ]

    # Append the new lines to the DataFrame
    rateio_lucas = rateio_lucas.append(new_lines, ignore_index=True)
    rateio_lucas_ = rateio_lucas.pivot_table(values = 'Amount', index = ['Product','Book'], columns = 'Trading Desk', aggfunc = 'sum', margins = True)
    rateio_lucas_.fillna(0, inplace = True)
except:
    pass

try:
    rateio_thiago = position_1.loc[position_1['Product'].isin(thiago_lote['Product']) & (position_1['Book'].isin(thiago_books))]
    
    rateio_thiago_ = rateio_thiago.pivot_table(values = 'Amount', index = ['Product','Book'], columns = 'Trading Desk', aggfunc = 'sum', margins = True)
    rateio_thiago_.fillna(0, inplace = True)
except:
    pass

try:
    rateio_matheus = position_2.loc[position_2['Product'].isin(matheus_lote['Product']) & (position_2['Book'].isin(matheus_books))]
    new_lines = [
    {"Trading Desk": "SPX LANCER", "Product": "", "Book": "Bolsa_BR", "Amount": 0},
    {"Trading Desk": "SPX LANCER PLUS", "Product": "", "Book": "Bolsa_BR", "Amount": 0},
    {"Trading Desk": "SPX CANADIAN EAGLE FUND", "Product": "", "Book": "Bolsa_BR", "Amount": 0},
    {"Trading Desk": "SPX NIMITZ", "Product": "", "Book": "Bolsa_BR", "Amount": 0},
    {"Trading Desk": "SPX RAPTOR", "Product": "", "Book": "Bolsa_BR", "Amount": 0},
    ]

    # Append the new lines to the DataFrame
    rateio_matheus = rateio_matheus.append(new_lines, ignore_index=True)
    
    rateio_matheus_ = rateio_matheus.pivot_table(values = 'Amount', index = ['Product','Book'], columns = 'Trading Desk', aggfunc = 'sum', margins = True)
    rateio_matheus_.fillna(0, inplace = True)
except:
    pass

try:
    fundos_fias = rateio_lucas['Trading Desk'].unique().tolist()
    ordem_fias = [1,2,3,4,0]
    fundos_fias = [fundos_fias[i] for i in ordem_fias]
except:
    pass

html_body = """
<div>
<span style = "font-family: Calibri; font-size: 14; color: RGB(0,0,0)">
                
Olá,
</span>
                
<br>
<br>
                
<span style = "font-family: Calibri; font-size: 14; color: RGB(0,0,0)">
                    
Segue em anexo as divergências entre os trades do LOTE45 e os Recaps.""" 

"""</span>
                
<br>
<br>

<div>
"""

html_div = """
<div>
<br>

<span style = "font-family: Calibri; font-size: 14; color: RGB(0,0,0)">
Divergências em relação ao que foi boletado no LOTE:

<br>
<br>
</span>
</div>

"""
html_voando = """
<div>

<span style = "font-family: Calibri; font-size: 14; color: RGB(0,0,0)">
Favor verificar os trades que estão sem trader identificado:

<br>
<br>
</span>
</div>

"""
html_enquadramento = """
<div>

<span style = "font-family: Calibri; font-size: 14; color: RGB(0,0,0)">
Exposição em Renda Variável:

<br>
<br>
</span>
</div>

"""
html_lote = """
<div>

<span style = "font-family: Calibri; font-size: 14; color: RGB(0,0,0)">
Estes são teus trades no LOTE:

<br>
<br>
</span>
</div>

"""

html_att = """
<div>

<span style = "font-family: Calibri; font-size: 14; color: RGB(0,0,0)">
Att,
<br>
</span>
</div>

"""
try:

    lucas_perc = pd.crosstab([rateio_lucas['Product'], rateio_lucas['Book']],rateio_lucas['Trading Desk'], values=rateio_lucas['Amount'], aggfunc=sum, normalize='index').applymap(lambda x: "{0:.2f}%".format(100*x))

    rateio_lucas_ = rateio_lucas_[['SPX FALCON', 'SPX FALCON INSTITUCIONAL', 'SPX LONG BIAS','SPX HORNET','SPX PATRIOT', 'SPX APACHE','All']]

    rateio_lucas_.reset_index([0], inplace = True)
    rateio_lucas_.reset_index([0], inplace = True)
    lucas_perc.reset_index([0], inplace = True)
    lucas_perc.reset_index([0], inplace = True)

    rateio_lucas_previa = rateio_lucas_.merge(lucas_perc, on = ['Book', 'Product'], how = 'outer',suffixes=(' Amount ', ' Perc '))
    rateio_lucas_previa_html = rateio_lucas_previa.style.applymap(color_negative_red, subset = ['SPX FALCON Amount ', 'SPX FALCON INSTITUCIONAL Amount ', 'SPX LONG BIAS Amount ','SPX HORNET Amount ', 'SPX PATRIOT Amount ', 'SPX APACHE Amount ','All']).format({'SPX FALCON Amount ': '{:,.0f}', 'SPX FALCON INSTITUCIONAL Amount ': '{:,.0f}', 'SPX LONG BIAS Amount ': '{:,.0f}', 'SPX HORNET Amount ': '{:,.0f}', 'SPX PATRIOT Amount ': '{:,.0f}', 'SPX APACHE Amount ': '{:,.0f}' ,'All': '{:,.0f}'}).set_table_styles(styles).hide_index().render()

    lucas_lote = validation.loc[validation['Trader'] == 'Lucas Rossi']

    lucas_lote  = lucas_lote.copy()

    lucas_lote['Amount'] = lucas_lote['Amount'].map('{:,.0f}'.format)
    lucas_lote['Amount_Recaps'] = lucas_lote['Amount_Recaps'].map('{:,.0f}'.format)

    lucas_lote['Price'] = 'R$ ' + lucas_lote['Price'].map('{:,.5f}'.format)
    lucas_lote['Price_Recaps'] = 'R$ ' + lucas_lote['Price_Recaps'].map('{:,.5f}'.format)

    lucas_div = validation.loc[(validation['Trader'] == 'Lucas Rossi') & ((validation['Quantity OK?'] != 'OK') | (validation['Price OK?'] != 'OK'))]

    lucas_div = lucas_div.copy()

    lucas_div['Amount'] = lucas_div['Amount'].map('{:,.0f}'.format)
    lucas_div['Amount_Recaps'] = lucas_div['Amount_Recaps'].map('{:,.0f}'.format)

    lucas_div['Price'] = 'R$ ' + lucas_div['Price'].map('{:,.5f}'.format)
    lucas_div['Price_Recaps'] = 'R$ ' + lucas_div['Price_Recaps'].map('{:,.5f}'.format)

    lucas_lote_html  = lucas_lote.style.apply(lambda x: ['background: #bcffb8' if x == 'OK' else 'background: #ffb8b8' for x in lucas_lote['Trade OK?']], axis = 0).hide_index().render()
    lucas_div_html  = lucas_div.style.apply(lambda x: ['background: #bcffb8' if x == 'OK' else 'background: #ffb8b8' for x in lucas_div['Trade OK?']], axis = 0).hide_index().render()

    base_expo_LongBias = pd.read_json("http://backtech.spxcapital.com:3006/api/overview/newoverview/funds/SPX%20LONG%20BIAS/mainReport/Produtos_Consolidados/bookLevel/3/date/{a}-{m}-{d}".format(a = ano, m = mes, d = dia))
    expo_RV_LongBias = base_expo_LongBias[base_expo_LongBias['Book'].isin(['Bolsa_Produtos_Onshore', 'Bolsa_Prev'])]
    expo_RV_LongBias = expo_RV_LongBias.groupby(["Trading Desk"])["EQUITY"].sum().reset_index()
    expo_RV_LongBias = int(expo_RV_LongBias['EQUITY'])
    NAV_Lote_LongBias = int(pd.read_json("http://backtech.spxcapital.com:3001/api/nav/fund/SPX%20LONG%20BIAS/date//{a}-{m}-{d}".format(a = ano, m = mes, d = dia))['nav'])
    perc_expo_LongBias = "Exposição em Renda Variável - Long Bias: {0}%".format(round(expo_RV_LongBias/NAV_Lote_LongBias*100,2))

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'lucas.rossi@spxcapital.com; joao.torres@spxcapital.com'
    mail.cc = 'controle@spxcapital.com.br'
    mail.Subject = 'Divergências Bovespa - Lucas Rossi (Lote45)'

    mail.HTMLBody = html_body + html_div + lucas_div_html + html_voando + voando_html + html_enquadramento + perc_expo_LongBias + "<br>" + html_lote + lucas_lote_html +  html_rateio + space + rateio_lucas_previa_html + space + html_att
    mail.Display()
except:
    pass

#RATEIO MACROS
try:
    thiago_perc = pd.crosstab([rateio_thiago['Product'], rateio_thiago['Book']],rateio_thiago['Trading Desk'], values=rateio_thiago['Amount'], aggfunc=sum, normalize='index').applymap(lambda x: "{0:.2f}%".format(100*x))

    rateio_thiago_.reset_index([0], inplace = True)
    rateio_thiago_.reset_index([0], inplace = True)
    thiago_perc.reset_index([0], inplace = True)
    thiago_perc.reset_index([0], inplace = True)

    rateio_hoje_np = rateio_hoje.to_numpy()
    rateio_hoje_np = np.delete(rateio_hoje_np, (0), axis = 1)

    m = len(rateio_thiago_)
    data = np.zeros((m,1), dtype = rateio_hoje_np.dtype) + np.transpose(rateio_hoje_np)
    data = data[:, [2,3,4,1,0]]

    filtro_rateio = rateio_thiago_.to_numpy()
    filtro_rateio = np.delete(filtro_rateio, [0,1,-1], 1)
    base_ajuste = rateio_thiago_.to_numpy()
    base_ajuste = np.delete(base_ajuste, [0,1,2,3,4,5,6], 1)
    base_ajuste = np.zeros((1,5), dtype = base_ajuste.dtype) + base_ajuste #6 fundos!!!

    ajuste_rateio = data * base_ajuste - filtro_rateio
    ajuste_rateio = pd.DataFrame(ajuste_rateio, columns = ['SPX CANADIAN EAGLE FUND Correction', 'SPX LANCER Correction', 'SPX LANCER PLUS Correction','SPX NIMITZ Correction', 'SPX RAPTOR Correction'])
    ajuste_rateio['Book'] = rateio_thiago_['Book']
    ajuste_rateio['Product'] = rateio_thiago_['Product']
    ajuste_rateio = ajuste_rateio[['Book','Product','SPX CANADIAN EAGLE FUND Correction', 'SPX LANCER Correction', 'SPX LANCER PLUS Correction','SPX NIMITZ Correction' ,'SPX RAPTOR Correction']]

    ####RATEIO FINAL
    rateio_thiago_previa = rateio_thiago_.merge(thiago_perc, on = ['Book', 'Product'], how = 'outer',suffixes=(' Amount ', ' Perc '))
    rateio_thiago_final = rateio_thiago_previa.merge(ajuste_rateio, on = ['Book','Product'], how = 'outer')
    rateio_thiago_final = rateio_thiago_final.sort_values(['Book', 'Product'])
    rateio_thiago_final_html =  rateio_thiago_final.style.applymap(color_negative_red, subset = ['SPX CANADIAN EAGLE FUND Amount ', 'SPX LANCER Amount ', 'SPX LANCER PLUS Amount ', 'SPX NIMITZ Amount ', 'SPX RAPTOR Amount ','All', 'SPX LANCER PLUS Correction', 'SPX CANADIAN EAGLE FUND Correction', 'SPX LANCER Correction', 'SPX NIMITZ Correction',  'SPX RAPTOR Correction']).format({'SPX LANCER PLUS Amount ': '{:,.0f}', 'SPX LANCER Amount ': '{:,.0f}', 'SPX CANADIAN EAGLE FUND Amount ': '{:,.0f}', 'SPX NIMITZ Amount ': '{:,.0f}', 'SPX RAPTOR Amount ': '{:,.0f}', 'All': '{:,.0f}', 'SPX LANCER PLUS Correction': '{:,.1f}', 'SPX LANCER Correction': '{:,.1f}', 'SPX CANADIAN EAGLE FUND Correction': '{:,.1f}', 'SPX NIMITZ Correction': '{:,.1f}', 'SPX RAPTOR Correction': '{:,.1f}'}).set_table_styles(styles).hide_index().render()

    thiago_lote = validation.loc[validation['Trader'] == 'Thiago Santos']
    thiago_lote =thiago_lote.copy()
    thiago_lote['Amount'] = thiago_lote['Amount'].map('{:,.0f}'.format)
    thiago_lote['Amount_Recaps'] = thiago_lote['Amount_Recaps'].map('{:,.0f}'.format)
    thiago_lote['Price'] = 'R$' + thiago_lote['Price'].map('{:,.5f}'.format)
    thiago_lote['Price_Recaps'] = 'R$' + thiago_lote['Price_Recaps'].map('{:,.5f}'.format)
    thiago_div = validation.loc[(validation['Trader'] == 'Thiago Santos') & ((validation['Quantity OK?'] == '') | (validation['Price OK?'] == ''))]
    thiago_div = thiago_div.copy()
    thiago_div['Amount'] = thiago_div['Amount'].map('{:,.0f}'.format)
    thiago_div['Amount_Recaps'] = thiago_div['Amount_Recaps'].map('{:,.0f}'.format)
    thiago_div['Price'] = 'R$' + thiago_div['Price'].map('{:,.5f}'.format)
    thiago_div['Price_Recaps'] = 'R$' + thiago_div['Price_Recaps'].map('{:,.5f}'.format)
    thiago_lote_html  = thiago_lote.style.apply(lambda x: ['background: #bcffb8' if x == 'OK' else 'background: #ffb8b8' for x in thiago_lote['Trade OK?']], axis = 0).hide_index().render()
    thiago_div_html  = thiago_div.style.apply(lambda x: ['background: #bcffb8' if x == 'OK' else 'background: #ffb8b8' for x in thiago_div['Trade OK?']], axis = 0).hide_index().render()

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'reports.execution@spxcapital.com'
    mail.cc = 'controle@spxcapital.com.br'
    mail.Subject = 'Divergências Bovespa - Execution Desk (Lote45)'

    mail.HTMLBody = html_body + html_div + thiago_div_html + html_voando + voando_html + html_lote + thiago_lote_html + space + html_rateio + rateio_thiago_final_html + space + html_att
    mail.Display()
except:
    pass

try:
    matheus_perc = pd.crosstab([rateio_matheus['Product'], rateio_matheus['Book']],rateio_matheus['Trading Desk'], values=rateio_matheus['Amount'], aggfunc=sum, normalize='index').applymap(lambda x: "{0:.2f}%".format(100*x))
    
    rateio_matheus_.reset_index([0], inplace = True)
    rateio_matheus_.reset_index([0], inplace = True)
    matheus_perc.reset_index([0], inplace = True)
    matheus_perc.reset_index([0], inplace = True)
    
    rateio_hoje_np = rateio_hoje.to_numpy()
    rateio_hoje_np = np.delete(rateio_hoje_np, (0), axis = 1)
    
    m = len(rateio_matheus_)
    data = np.zeros((m,1), dtype = rateio_hoje_np.dtype) + np.transpose(rateio_hoje_np)
    data = data[:, [2,3,4,1,0]]
    filtro_rateio = rateio_matheus_.to_numpy()
    filtro_rateio = np.delete(filtro_rateio, [0,1,-1], 1)
    base_ajuste = rateio_matheus_.to_numpy()
    base_ajuste = np.delete(base_ajuste, [0,1,2,3,4,5,6], 1)
    base_ajuste = np.zeros((1,5), dtype = base_ajuste.dtype) + base_ajuste #6 fundos!!!
    
    ajuste_rateio = data * base_ajuste - filtro_rateio
    ajuste_rateio = pd.DataFrame(ajuste_rateio, columns = ['SPX CANADIAN EAGLE FUND Correction', 'SPX LANCER Correction', 'SPX LANCER PLUS Correction','SPX NIMITZ Correction', 'SPX RAPTOR Correction'])
    ajuste_rateio['Book'] = rateio_matheus_['Book']
    ajuste_rateio['Product'] = rateio_matheus_['Product']
    ajuste_rateio = ajuste_rateio[['Book','Product','SPX CANADIAN EAGLE FUND Correction', 'SPX LANCER Correction', 'SPX LANCER PLUS Correction','SPX NIMITZ Correction', 'SPX RAPTOR Correction']]
    
    ####RATEIO FINAL
    rateio_matheus_previa = rateio_matheus_.merge(matheus_perc, on = ['Book', 'Product'], how = 'outer',suffixes=(' Amount ', ' Perc '))
    rateio_matheus_final = rateio_matheus_previa.merge(ajuste_rateio, on = ['Book','Product'], how = 'outer')
    rateio_matheus_final_html =  rateio_matheus_final.style.applymap(color_negative_red, subset = ['SPX CANADIAN EAGLE FUND Amount ', 'SPX LANCER Amount ', 'SPX LANCER PLUS Amount ', 'SPX NIMITZ Amount ', 'SPX RAPTOR Amount ','All', 'SPX LANCER PLUS Correction', 'SPX CANADIAN EAGLE FUND Correction', 'SPX LANCER Correction', 'SPX NIMITZ Correction', 'SPX RAPTOR Correction']).format({'SPX LANCER PLUS Amount ': '{:,.0f}', 'SPX LANCER Amount ': '{:,.0f}', 'SPX CANADIAN EAGLE FUND Amount ': '{:,.0f}', 'SPX NIMITZ Amount ': '{:,.0f}', 'SPX RAPTOR Amount ': '{:,.0f}', 'All': '{:,.0f}', 'SPX LANCER PLUS Correction': '{:,.1f}', 'SPX LANCER Correction': '{:,.1f}', 'SPX CANADIAN EAGLE FUND Correction': '{:,.1f}', 'SPX NIMITZ Correction': '{:,.1f}', 'SPX RAPTOR Correction': '{:,.1f}'}).set_table_styles(styles).hide_index().render()
    
    matheus_lote = validation.loc[validation['Trader'] == 'Equities Brasil']
    matheus_lote =matheus_lote.copy()
    matheus_lote['Amount'] = matheus_lote['Amount'].map('{:,.0f}'.format)
    matheus_lote['Amount_Recaps'] = matheus_lote['Amount_Recaps'].map('{:,.0f}'.format)
    matheus_lote['Price'] = 'R$' + matheus_lote['Price'].map('{:,.5f}'.format)
    matheus_lote['Price_Recaps'] = 'R$' + matheus_lote['Price_Recaps'].map('{:,.5f}'.format)
    matheus_div = validation.loc[(validation['Trader'] == 'Equities Brasil') & ((validation['Quantity OK?'] == '') | (validation['Price OK?'] == ''))]
    matheus_div = matheus_div.copy()
    matheus_div['Amount'] = matheus_div['Amount'].map('{:,.0f}'.format)
    matheus_div['Amount_Recaps'] = matheus_div['Amount_Recaps'].map('{:,.0f}'.format)
    matheus_div['Price'] = 'R$' + matheus_div['Price'].map('{:,.5f}'.format)
    matheus_div['Price_Recaps'] = 'R$' + matheus_div['Price_Recaps'].map('{:,.5f}'.format)
    matheus_lote_html  = matheus_lote.style.apply(lambda x: ['background: #bcffb8' if x == 'OK' else 'background: #ffb8b8' for x in matheus_lote['Trade OK?']], axis = 0).hide_index().render()
    matheus_div_html  = matheus_div.style.apply(lambda x: ['background: #bcffb8' if x == 'OK' else 'background: #ffb8b8' for x in matheus_div['Trade OK?']], axis = 0).hide_index().render()
    
    base_expo_Lancer = pd.read_json("http://backtech.spxcapital.com:3006/api/overview/newoverview/funds/SPX%20LANCER/mainReport/Produtos_Consolidados/bookLevel/3/date/{a}-{m}-{d}".format(a = ano, m = mes, d = dia))
    expo_RV_Lancer = base_expo_Lancer[base_expo_Lancer['Book'].isin(['Bolsa_Produtos_Onshore', 'Bolsa_Prev'])]
    expo_RV_Lancer = expo_RV_Lancer.groupby(["Trading Desk"])["EQUITY"].sum().reset_index()
    expo_RV_Lancer = int(expo_RV_Lancer['EQUITY'])
    NAV_Lote_Lancer = int(pd.read_json("http://backtech.spxcapital.com:3001/api/nav/fund/SPX%20LANCER/date//{a}-{m}-{d}".format(a = ano, m = mes, d = dia))['nav'])
    perc_expo_Lancer = "Exposição em Renda Variável - Lancer Prev: {0}%".format(round(expo_RV_Lancer/NAV_Lote_Lancer*100,2))
    
    base_expo_Lancer_Plus = pd.read_json("http://backtech.spxcapital.com:3006/api/overview/newoverview/funds/SPX%20LANCER%20PLUS/mainReport/Produtos_Consolidados/bookLevel/3/date/{a}-{m}-{d}".format(a = ano, m = mes, d = dia))
    expo_RV_LancerPlus = base_expo_Lancer_Plus[base_expo_Lancer_Plus['Book'].isin(['Bolsa_Produtos_Onshore', 'Bolsa_Prev'])]
    expo_RV_LancerPlus = expo_RV_LancerPlus.groupby(["Trading Desk"])["EQUITY"].sum().reset_index()
    expo_RV_LancerPlus = int(expo_RV_LancerPlus['EQUITY'])
    NAV_Lote_LancerPlus = int(pd.read_json("http://backtech.spxcapital.com:3001/api/nav/fund/SPX%20LANCER%20PLUS/date//{a}-{m}-{d}".format(a = ano, m = mes, d = dia))['nav'])
    perc_expo_LancerPlus = "Exposição em Renda Variável - Lancer Plus: {0}%".format(round(expo_RV_LancerPlus/NAV_Lote_LancerPlus*100,2))
    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'reports.execution@spxcapital.com; levi.Jesus@spxcapital.com'
    mail.cc = 'controle@spxcapital.com.br'
    mail.Subject = 'Divergências Bovespa - Equities Brasil (Lote45)'
    
    mail.HTMLBody = html_body + html_div + matheus_div_html + html_voando + voando_html + html_enquadramento + perc_expo_Lancer + "<br>" + perc_expo_LancerPlus + "<br>" + html_lote + matheus_lote_html + space + html_rateio + rateio_matheus_final_html + space + html_att
    mail.Display()
except:
    pass

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'controle@spxcapital.com.br'
mail.cc = 'controle@spxcapital.com.br'
mail.Subject = 'Divergências Bovespa (Lote45)'

lote_amount = lote.groupby(['Product','Dealer','Side']).aggregate({'Amount': 'sum' })

lote_amount.name = "Amount"

lote_price = lote.groupby(['Product','Dealer','Side']).apply(wavg, 'Price','Amount')

lote_price.name = "Price"

lote = lote_amount.merge(lote_price, on = ['Product','Dealer','Side'], how = "left")

lote.reset_index([0,1,2], inplace = True)

lote['Amount'] = lote['Amount'].abs()

lote.fillna('', inplace = True)

lote.columns = ['Product','Dealer','Side','Amount','Price']

validation = lote.merge(trades_final, on = ['Product','Dealer','Side'], how = 'outer',suffixes=('', '_Recaps'))

validation.loc[validation['Amount'] == validation['Amount_Recaps'], 'Quantity OK?'] = 'OK'

validation.loc[abs(validation['Price'] - validation['Price_Recaps']) < 0.001, 'Price OK?'] = 'OK'

validation.loc[(validation['Quantity OK?'] == 'OK') & (validation['Price OK?'] == 'OK'), 'Trade OK?'] = 'OK'

validation.fillna(0, inplace = True)

validation.loc[validation['Quantity OK?'] == 0, 'Quantity OK?'] = ''
validation.loc[validation['Price OK?'] == 0, 'Price OK?'] = ''
validation.loc[validation['Trade OK?'] == 0, 'Trade OK?'] = ''

validation['Amount_Recaps'] = validation['Amount_Recaps'].astype('int64') 

labels_valid =  ['Product', 'Dealer', 'Side', 'Amount', 'Amount_Recaps', 'Price', 'Price_Recaps', 'Quantity OK?', 'Price OK?', 'Trade OK?']

validation = validation[labels_valid]

validation_html  = validation.style.apply(lambda x: ['background: #bcffb8' if x == 'OK' else 'background: #ffb8b8' for x in validation['Trade OK?']], axis = 0).hide_index().render()

## checagem de enquadramento
#dicionário opção vs ação underlyer
base_opt = pd.read_json("http://10.21.2.20/api/option/")
base_opt = base_opt[['cod_sys','product_class', 'currency', 'sec_type', 'underlying_ticker']]
options = depara_b3['opções'].dropna()
options = options.to_list()
base_opt = base_opt[base_opt['product_class'].isin(options)]
base_opt['underlying_ticker'] = base_opt['underlying_ticker'].str.replace(' (BZ EQUITY|INDEX)$', '')
dict_opt = base_opt.set_index('cod_sys').to_dict()['underlying_ticker']

#operações do dia alocadas por fundo
base_alocacao = pd.read_json("http://10.21.2.23:3000/api/trades/setfunds//date/{a}-{m}-{d}".format(a = ano, m = mes, d = dia))
base_alocacao = base_alocacao[base_alocacao['IsCashFlowTrade'] == False]
base_alocacao = base_alocacao[base_alocacao['IsReplicatedTrade'] == False]
base_alocacao = base_alocacao[['Trading Desk', 'ProductClass', 'Product', 'Amount']]
products = depara_b3['produtos'].dropna()
products = products.to_list()
funds = depara_b3['funds'].dropna()
funds = funds.to_list()
base_alocacao = base_alocacao[base_alocacao['ProductClass'].isin(products)]
base_alocacao = base_alocacao[base_alocacao['Trading Desk'].isin(funds)]

#pegar apenas underlyers de opções operadas
lista_opts_operadas = base_alocacao[base_alocacao['ProductClass'].isin(options)]
lista_opts_operadas = lista_opts_operadas.replace({"Product": dict_opt})
lista_opts_operadas = lista_opts_operadas[['Trading Desk', 'Product']]
lista_opts_operadas = lista_opts_operadas.rename(columns={'Trading Desk': 'Fund'})

opts_operadas = base_alocacao.replace({"Product": dict_opt})
opts_operadas = opts_operadas.rename(columns={'Trading Desk': 'Fund'})
opts_operadas = opts_operadas[['Fund', 'Product', 'Amount']]

#posição base para checar se algum fundo está operando opt descoberto
base_posicao = pd.read_csv(r'G:\Carteiras Oficiais\Lote45\AllTradingDesksOverView{d}{m}{a}.txt'.format(d = dia_anterior, m = mes_anterior, a = ano_anterior), sep = '	')
base_posicao = base_posicao[["Fund", 'Product', 'ProductClass', 'Amount']]
base_posicao = base_posicao[base_posicao['ProductClass'].isin(products)]
base_posicao = base_posicao.groupby(["Fund", 'Product', 'ProductClass'])["Amount"].sum().reset_index()
base_posicao = base_posicao[base_posicao['Amount'] != 0]
base_posicao = base_posicao.replace({"Product": dict_opt}) #consolidar a posição com as opts já operadas
base_posicao = base_posicao.groupby(["Fund", 'Product'])["Amount"].sum().reset_index()
base_posicao = base_posicao[base_posicao['Fund'].isin(funds)]

#consolidando posição em aberto + operações do dia (incluindo opções)
base_dia = pd.merge(base_posicao,opts_operadas, how = "outer")
base_dia = base_dia.groupby(["Fund", 'Product'])["Amount"].sum().reset_index()

#filtrando posição do dia apenas com ativos operados com opções no dia
base_dia_short = base_dia[base_dia['Amount'] < 0]
opt_descoberto = pd.merge(lista_opts_operadas, base_dia_short, on=['Product', 'Fund'])
depara_b3_enquadramento = pd.read_excel('G:\Batimento BackOffice\Renda Variável\Batimento B3\De_Para_B3(Pythons).xlsx', sheet_name= 'Enquadramento')
fundos_descobertos = depara_b3_enquadramento['Opts Descobertas'].dropna()
fundos_descobertos = fundos_descobertos.to_list()
opt_descoberto = opt_descoberto[opt_descoberto['Fund'].isin(fundos_descobertos)]
if opt_descoberto.shape[0] > 0:
    message_descoberto = "\nEXISTEM OPÇÔES DESCOBERTAS QUE NÃO PODEM SER ALOCADAS NOS FUNDOS\n" + opt_descoberto.to_html(index=False)
else:
    message_descoberto = '\nNÃO HÁ DESENQUADRAMENTO POR OPÇÕES DESCOBERTAS'
    
#filtrando posições short em fundos que não podem ficar vendidos
fundos_short = depara_b3_enquadramento['Posições Short'].dropna()
fundos_short = fundos_short.to_list()
fundos_sort_pos = base_dia[base_dia['Fund'].isin(fundos_short)]
fundos_sort_pos = fundos_sort_pos[fundos_sort_pos['Amount'] < -0.01]
if fundos_sort_pos.shape[0] > 0:
    message_short = "\nEXISTEM POSIÇÕES SHORT NOS FUNDOS ABAIXO\n" +  fundos_sort_pos.to_html(index=False)
else:
    message_short = '\nNÃO HÁ DESENQUADRAMENTO POR POSIÇÕES SHORT'
    
#Daytrade
fundos_daytrade = depara_b3_enquadramento['Daytrade'].dropna()
fundos_daytrade = fundos_daytrade.to_list()
base_alocacao_short = base_alocacao[base_alocacao['Amount'] < 0]
base_alocacao_short = base_alocacao_short[base_alocacao_short['Trading Desk'].isin(fundos_daytrade)]
base_alocacao_long = base_alocacao[base_alocacao['Amount'] > 0]
base_alocacao_long = base_alocacao_long[base_alocacao_long['Trading Desk'].isin(fundos_daytrade)]
ativos_daytrade = pd.merge(base_alocacao_short, base_alocacao_long, on=['Product','Trading Desk','ProductClass'], how = 'inner')
ativos_daytrade = ativos_daytrade[['Product','Trading Desk']].drop_duplicates()

if base_alocacao_short['Product'].isin(base_alocacao_long['Product']).any():
    message_daytrade = "\nEXISTEM DAYTRADES NOS FUNDOS ABAIXO\n" +  ativos_daytrade.to_html(index=False)
else:
    message_daytrade = '\nNÃO HÁ DESENQUADRAMENTO POR DAYTRADE'

base_b3 = pd.read_excel('G:\\Batimento BackOffice\\Renda Variável\\Batimento NEGS - IPN - API.xlsm', sheet_name= 'FERIADOS')

#checagem de ativos proibidos
lista_ativos_proibidos = base_b3[['Ativos Proibidos', 'Fundo.4']].dropna()
lista_ativos_proibidos = lista_ativos_proibidos.rename(columns={'Fundo.4': 'Fund', 'Ativos Proibidos': 'Product'})
lista_ativos_proibidos['Product'] = lista_ativos_proibidos['Product'].str.replace(' (Equity)$', '')
ativos_proibidos = pd.merge(lista_ativos_proibidos, base_dia, on=['Product', 'Fund'], indicator=True, how='inner').dropna()
ativos_proibidos = ativos_proibidos[['Product', 'Fund', 'Amount']]
if ativos_proibidos.shape[0] > 0:
    message_proibidos = "\nEXISTEM ATIVOS PROIBIDOS\n" + ativos_proibidos.to_html(index=False)
else:
    message_proibidos = '\nNÃO HÁ DESENQUADRAMENTO POR ATIVOS PROIBIDOS'
   
#checagem de ativos relacionados- prev
ativos_prev_fundos = base_b3[['Ativo Vetado', 'Seguradora']].dropna()
ativos_prev_seguradoras = base_b3[['Fundo.5', 'Seguradora.1']].dropna()
ativos_prev_seguradoras = ativos_prev_seguradoras.rename(columns={'Fundo.5': 'Fund', 'Seguradora.1': 'Seguradora'})
ativos_prev = pd.merge(ativos_prev_seguradoras, ativos_prev_fundos, how='left')
ativos_prev = ativos_prev.rename(columns={'Fundo.5': 'Fund', 'Ativo Vetado': 'Product'})
ativos_prev_controle = pd.merge(ativos_prev, base_dia, on=['Product', 'Fund'], indicator=True, how='right').dropna()
ativos_prev_controle = ativos_prev_controle[['Fund', 'Seguradora', 'Product', 'Amount']]
if ativos_prev_controle.shape[0] > 0:
    message_prev = "\nATIVOS RELACIONADOS ALOCADOS NOS FUNDOS DE PREVIDÊNCIA\n" + ativos_prev_controle.to_html(index=False)
else:
    message_prev = '\nNÃO HÁ ATIVOS RELACIONADOS NOS FUNDOS DE PREVIDÊNCIA'

#checagem de ativos positivos
ativos_positivos = base_b3[['Ativo Positivo', 'Fundo.6']].dropna()
ativos_positivos = ativos_positivos.rename(columns={'Fundo.6': 'Fund', 'Ativo Positivo': 'Product'})
ativos_sem_cadastro = pd.merge(ativos_positivos, base_dia, on=['Product', 'Fund'], indicator=True, how='outer')
ativos_sem_cadastro = ativos_sem_cadastro[ativos_sem_cadastro['_merge'] == 'right_only']
ativos_sem_cadastro = ativos_sem_cadastro[['Product', 'Fund', 'Amount']]
if ativos_sem_cadastro.shape[0] > 0:
    message_positivos = "\nEXISTEM ATIVOS SEM CADASTRO POSITIVO\n" + ativos_sem_cadastro.to_html(index=False)
else:
    message_positivos = '\nTODOS OS ATIVOS ESTÃO CADASTRADOS NA LISTA DE ATIVOS POSITIVOS\n'
    
try:
    mail.HTMLBody = message_descoberto + "<br>" + message_short  + "<br>" + message_daytrade + "<br>" + message_proibidos + "<br>" + message_prev + "<br>" + message_positivos + "<br>" + "<br>" + perc_expo_Lancer + "<br>" + perc_expo_LancerPlus + "<br>" + perc_expo_LongBias + "<br>" + "<br>" + validation_html
except:
    try:
        mail.HTMLBody = message_descoberto + "<br>" + message_short  + "<br>" + message_daytrade + "<br>" + message_proibidos + "<br>" + message_prev + "<br>" + message_positivos + "<br>" + "<br>" + perc_expo_Lancer + "<br>" + perc_expo_LancerPlus + "<br>" + "<br>" + validation_html
    except: 
        try:
            mail.HTMLBody = message_descoberto + "<br>" + message_short  + "<br>" + message_daytrade + "<br>" + message_proibidos + "<br>" + message_prev + "<br>" + message_positivos + "<br>" + "<br>" + perc_expo_LongBias + "<br>" + "<br>" + validation_html
        except:
            mail.HTMLBody = message_descoberto + "<br>" + message_short  + "<br>" + message_daytrade + "<br>" + message_proibidos + "<br>" + message_prev + "<br>" + message_positivos + "<br>" + "<br>" + validation_html
mail.Display()
