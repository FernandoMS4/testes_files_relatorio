import pandas as pd

def transformar_e_abrir_relatorio_responsys(path:str,agg:str)->str:
    df = pd.DataFrame(pd.read_excel(f'{path}.xlsx',skiprows=3))
    try:
        df[agg] = df[agg].ffill()
    except:
        pass
    df = df.fillna(0)
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
                    df_gen = df[i].drop_duplicates()
                    df2 = pd.merge(df2,df_gen, on=agg)
                else:
                    pass
            except KeyError as e:
                print(e)
    return df2

if __name__ == "__main__":
    print(transformar_e_abrir_relatorio_responsys("unionv2","Campaign"))
