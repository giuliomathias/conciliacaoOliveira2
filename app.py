import streamlit as st
import pandas as pd
import pandas as pd
from datetime import datetime,timedelta
import holidays
import pyodbc
import base64
import io

br_holidays = holidays.BR()
br_holidays = holidays.country_holidays('BR')

dataInical = datetime.today()-timedelta(days=90)
dataFinal = datetime.today()

#sabado = 5
#domingo = 6

if dataInical.weekday() == 5:
    dataInical = dataInical-timedelta(days=1)
    while dataInical in br_holidays:
        dataInical = datetime.today()-timedelta(days=1)
    
if dataInical.weekday() == 6:
    dataInical = dataInical-timedelta(days=2)
    while dataInical in br_holidays:
        dataInical = datetime.today()-timedelta(days=1)
        
while dataInical in br_holidays:
    dataInical = dataInical-timedelta(days=1)
    if dataInical.weekday() == 6:
        dataInical = dataInical-timedelta(days=2)
        while dataInical in br_holidays:
            dataInical = datetime.today()-timedelta(days=1)
    if dataInical.weekday() == 5:
        dataInical = dataInical-timedelta(days=1)
        while dataInical in br_holidays:
            dataInical = datetime.today()-timedelta(days=1)

dataInical = '2023-10-31'
dataFinal = '2023-11-30'         

print(dataInical)
print(dataFinal)

conn = pyodbc.connect(
            'Driver={SQL Server};'
            'Server=SP3LS05;'
            'Database=SQL_TEMA_FUNDOS;'
                    )

def get_cota_tema_sql(dataInical,dataFinal,cnpj):
    query = f'''
    SELECT 
    C.CODCARTEIRA as CODCARTEIRA
    ,CLI.NOMCLIENTE as NOME 
    ,CPJU.NUMCGC as CGC
    ,C.DTACOTA as DATA
    ,C.VALCOTA as CotaTema
    ,RT.NomRotina as ROTINA
    ,RT.CodRotina as CODROTINA
    FROM SQL_TEMA_FUNDOS.DBO.FDO_COTA C WITH(NOLOCK)
    INNER JOIN SQL_TEMA_FUNDOS.DBO.CLI_GERAL CLI WITH(NOLOCK) ON C.CODCARTEIRA = CLI.CODCLIENTE
    INNER JOIN SQL_TEMA_FUNDOS.DBO.CLI_PJURIDICA CPJU WITH(NOLOCK) ON C.CODCARTEIRA = CPJU.CODCLIENTE
    INNER JOIN FDO_Rotina RT WITH(NOLOCK) ON C.CodRotina = RT.CodRotina 
    WHERE DTACOTA BETWEEN '{str(dataInical)[:10]}' and '{str(dataFinal)[:10]}'
    AND RT.CodRotina = '57000'
    AND CPJU.NUMCGC = '{cnpj}'
    '''    
    df_sql = pd.read_sql_query(query,conn)
    return df_sql 

def to_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer,index=False,sheet_name='Sheet1')
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def convert_df(df):
    # IMPORTANT: Cache the conversion to prevent computation on every rerun
    return df.to_csv(index=False).encode('utf-8')


st.title('Conciliação Oliveira')

uploaded_file = st.file_uploader("Selecione o arquivo Excel", type="xlsx")
if uploaded_file is not None:
    df1 = pd.read_excel(uploaded_file)
    df1['CNPJ'] = [x.replace('.','').replace('/','').replace('-','').zfill(14) for x in df1['CNPJ']]
    df1['Dt Posição'] = pd.to_datetime(df1['Dt Posição'])
    df1['Valor Cota'] = df1['Valor Cota'].apply(lambda x: float(x.replace(',','.')))

    cnpjs = df1['CNPJ'].drop_duplicates().tolist()
    dataframe=pd.DataFrame()
    for cnpj in cnpjs:
        cota_tema = get_cota_tema_sql(dataInical,dataFinal,cnpj)
        cota_tema['CODCARTEIRA'] = cota_tema['CODCARTEIRA'].astype(str)
        cota_tema['DATA'] = pd.to_datetime(cota_tema['DATA'])
       #st.write(cota_tema)

        dfInput = df1[df1['CNPJ']==cnpj][['Carteira','CNPJ','Dt Posição','Valor Cota']]
        dfInput = dfInput.set_axis(['Carteira','CNPJ','DATA','CotaArquivo'],axis=1)

        dfFinal = dfInput.merge(cota_tema[['CODCARTEIRA','NOME','DATA','CotaTema']],how='left',left_on='DATA',right_on='DATA')
        #print(dfFinal)
        dfFinal = dfFinal[['Carteira','NOME','CODCARTEIRA','CNPJ','DATA','CotaArquivo','CotaTema']]
        dfFinal = dfFinal.fillna('N/A')
        lista=[]
        for posicao in range(len(dfFinal)):
            cond1 = dfFinal['CotaArquivo'][posicao]
            cond2 = dfFinal['CotaTema'][posicao]
            try:
                lista.append(cond1 - cond2)
            except:
                if cond1 == 'N/A' and cond2 == 'N/A':
                    lista.append('não tem cota pra data')
                else:
                    if cond1 == 'N/A':
                        lista.append('não tem cota no arquivo')
                    else:
                        lista.append('não tem cota na TEMA')
        dfFinal['DIF'] = lista
        nome = dfFinal['Carteira'].drop_duplicates().to_list()[0]
        st.write(nome, ' conciliado')
        dataframe = pd.concat([dataframe,dfFinal])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, sheet_name='Sheet1', index=False)
    output.seek(0)
    st.download_button(label='Download', data=output, file_name='fundosConciliados.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    conn.close()


    
    
