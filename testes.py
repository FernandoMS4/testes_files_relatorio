import pandas as pd

def criar_df_skip_rows(path:str)->str:
    sk_rows = 0
    valid = False
    while valid == False:
        dataframe = pd.DataFrame(pd.read_excel(f'{path}.xlsx',skiprows=sk_rows))
        if "Unnamed" in str(dataframe.columns):
            sk_rows +=1
        else:
            valid = True
            df = dataframe.fillna(0)
    return df

def transformar_e_abrir_relatorio_responsys(df:str,agg:str)->str:
    try:
        df[agg] = df[agg].ffill()
    except:
        pass
    df.columns = df.columns.str.replace(" ","_").str.upper()
    agg = agg.upper()
    try:
        df = df.drop("MONTH_NUMBER",axis='columns')
    except:
        pass
    try:
        df2 = df[agg].drop_duplicates()
    except:
        pass
    for i in df:
        if i == agg:
            pass
        else:
            filtro = [agg,i]
            df_gen = df[filtro]
            try:
                if df[i].dtypes == 'float64':
                    df_gen = df_gen.groupby(agg).mean()
                    df2 = pd.merge(df2,df_gen, on=agg)
                elif df[i].dtypes == 'int64':
                    df_gen = df_gen.groupby(agg).sum()
                    df2 = pd.merge(df2,df_gen, on=agg)
                elif df[i].dtypes == 'object':
                    df2 = pd.merge(df2,df_gen,how='inner',on=agg)
                else:
                    pass
            except KeyError as e:
                print(e)
    return df2

if __name__ == "__main__":
    df = criar_df_skip_rows('unionv2')
    df = df[(df["Campaign"] != '0') or (df["Campaign"] != 'Rows 1 - 2958 (All Rows)')]
    df2= transformar_e_abrir_relatorio_responsys(df,"Campaign")
    df2.to_csv("teste_df_3_.csv",index=False,header=True)