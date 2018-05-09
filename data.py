import pandas as pd
import docx

vendor_label: str = '供应商'
quote_label: str = '报价'
index_label: str = '序号'
formatting: str = '%.2f'
unit_label: str = '单价'
no_quote: str = '无法报价'

df: pd.DataFrame = pd.read_excel('bay.xlsx', sheet_name=0, na_values=no_quote, index_col=0)


# prepare word document
def export(name: str, df: pd.DataFrame, quote_detail: list) -> None:
    doc = docx.Document()
    # add a table to the end and create a reference variable
    # extra row is so we can add the header row
    t = doc.add_table(df.shape[0] + 1, df.shape[1] + 1)
    t.style = 'TableGrid'

    t.cell(0, 0).text = index_label
    # add the header rows.
    for j in range(df.shape[-1]):
        t.cell(0, j + 1).text = df.columns[j]

    # add the header columns.
    for i in range(df.shape[0]):
        t.cell(i + 1, 0).text = str(df.index[i])

    # add the rest of the data frame
    for i in range(df.shape[0]):
        for j in range(df.shape[-1]):
            t.cell(i + 1, j + 1).text = str(df.values[i, j])

    for index, many_df in quote_detail:
        doc.add_paragraph(index_label + ','.join(map(str, index)))
        doc.add_paragraph(many_df.to_csv(sep='\t', float_format=formatting))

    # save the doc
    doc.save('./' + name + '.docx')
    return None


vendors = pd.read_excel('bay.xlsx', sheet_name=1, header=None).iloc[:, 0].tolist()
quotes = df.loc[:, vendors]

selection = quotes.idxmin(1)
selection.name = vendor_label
selection_quotes = quotes.min(1)
selection_quotes.name = quote_label
unit_quotes = selection_quotes / df.数量
unit_quotes.name = unit_label

df = df.join(selection)
df = df.join(selection_quotes)
df = df.join(unit_quotes)

groups = [group for group in df.groupby(vendor_label)]

no_quotes = df.loc[df[vendor_label].isna()].drop(vendors, 1).drop(vendor_label, 1)
print('无报价')
print(no_quotes.to_csv(sep='\t', float_format=formatting))

for vendor_name, vendor_df in groups:
    print('=============================')
    print(vendor_name)
    print('=============================')

    vendor_product = vendor_df.drop(vendors, 1).drop(vendor_label, 1)
    print('总价', formatting % vendor_df[quote_label].sum())
    print(vendor_product.to_csv(sep='\t', float_format=formatting))

    vendor_quotes = vendor_df.loc[:, vendors]

    quote_groups = vendor_quotes.notna().groupby(vendors).groups
    quote_detail = [(v.tolist(), vendor_quotes.loc[v].sum(skipna=False).sort_values().fillna(no_quote)) for k, v in
                    quote_groups.items()]

    print('>>> 明细')
    for index, many_df in quote_detail:
        print("序号:", ','.join(map(str, index)))
        print(many_df.to_csv(sep='\t', float_format=formatting))

    export(vendor_name, vendor_product, quote_detail)
