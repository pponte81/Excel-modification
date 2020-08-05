import pandas as pd

def formato_tabla(filename):
    xls = pd.ExcelFile(filename)
    hojas = xls.sheet_names
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    i = 0
    while i < len(xls.sheet_names):
        hoja = xls.sheet_names[i]
        df = pd.read_excel(filename, sheet_name=hoja)
        df = df.applymap(str)
        df.to_excel(writer, index=False, sheet_name=hoja)
        if df.shape[0] > 0:
            wb = writer.book
            ws = writer.sheets[hoja]
            col_names = [{'header': col_name} for col_name in df.columns]
            ws.add_table(0, 0, df.shape[0], df.shape[1]-1, {'columns': col_names})
            longitudes_campos = [max([len(row) for row in df[df.columns[y]]]) for y in range(len(df.columns))]
            longitudes_nombres = [len(col)+4 for col in df.columns]
            longitudes_max = list(zip(longitudes_campos, longitudes_nombres))
            longitudes_finales = [max(i) for i in longitudes_max]
            for idx, width in enumerate(longitudes_finales):
                ws.set_column(idx, idx, width)
        i +=1

    writer.save()
    writer.close()
    print("Fichero con formato tabla generado")
