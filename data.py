import io

import pandas as pd

vendor_label: str = '供应商'
quote_label: str = '报价'
index_label: str = '序号'
formatting: str = '%.2f'
unit_label: str = '单价'
no_quote: str = '无法报价'
excel_filename: str = 'bay.xlsx'

df: pd.DataFrame = pd.read_excel(excel_filename, sheet_name=0, na_values=no_quote, index_col=0)

vendors_raw: pd.Series = pd.read_excel(excel_filename, sheet_name=1, header=None, index_col=0, squeeze=True)
vendors_name = vendors_raw.index
vendors = list(filter(lambda v: v in df, vendors_raw.index))
vendors_timeout = list(filter(lambda v: v not in df.columns, vendors_raw.index))

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

excel_writer = pd.ExcelWriter('final.xlsx')

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

    numbers = set(vendor_product['编号'].unique())
    print(numbers)

    vendor_quotes = vendor_df.loc[:, vendors]

    quote_groups = vendor_quotes.notna().groupby(vendors).groups
    quote_detail = [(v.tolist(), vendor_quotes.loc[v].sum(skipna=False).sort_values()) for k, v in
                    quote_groups.items()]

    detail = io.StringIO()

    for index, detail_df in quote_detail:
        print("询价单", ','.join(map(str, index)), '项', file=detail)
        for detail_vendor, row in detail_df.items():
            if pd.isna(row):
                print(vendors_raw[detail_vendor], '回复无法报价', file=detail)
            else:
                print(vendors_raw[detail_vendor], '报价人民币', '%.2f' % row, '元', file=detail)
        for v in vendors_timeout:
            print(vendors_raw[v], '逾期未回复', file=detail)
        print(file=detail)
    detail_str: str = detail.getvalue()
    detail.close()

    print(detail_str)

    vendor_product.to_excel(excel_writer, sheet_name=vendor_name)

excel_writer.save()
