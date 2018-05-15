import io

import docxtpl
import pandas as pd
from docxtpl import DocxTemplate

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

vendor_all = len(vendors_raw)
vendor_count = len(quotes.columns[~quotes.isna().all()])

no_quotes = df.loc[df[vendor_label].isna()].drop(vendors, 1).drop(vendor_label, 1)
print('无报价')
print(no_quotes.to_csv(sep='\t', float_format=formatting))

for vendor_name, vendor_df in groups:
    print('=============================')
    print(vendor_name)
    print('=============================')

    vendor_product = vendor_df.drop(vendors, 1).drop(vendor_label, 1)
    total: str = formatting % vendor_df[quote_label].sum()
    print('总价', total)
    print(vendor_product.to_csv(sep='\t', float_format=formatting))

    numbers = ','.join(map(str, vendor_product['编号'].unique()))
    print(numbers)

    vendor_quotes = vendor_df.loc[:, vendors]

    quote_groups = vendor_quotes.notna().groupby(vendors).groups
    quote_detail = [(v.tolist(), vendor_quotes.loc[v].sum(skipna=False).sort_values()) for k, v in
                    quote_groups.items()]

    detail: io.StringIO = io.StringIO()
    vendor_detail: io.StringIO = io.StringIO()

    for index, detail_df in quote_detail:
        print("询价单第", ','.join(map(str, index)), '项', file=detail)
        print("询价单第", ','.join(map(str, index)), '项共有', detail_df.count(), '家供应商', file=vendor_detail)
        for detail_vendor, row in detail_df.items():
            if pd.isna(row):
                print(vendors_raw[detail_vendor], '回复无法报价', file=detail)
            else:
                print(vendors_raw[detail_vendor], '报价人民币', '%.2f' % row, '元', file=detail)
        for v in vendors_timeout:
            print(vendors_raw[v], '逾期未回复', file=detail)
        print(file=detail)

    detail_str: str = detail.getvalue()
    vendor_detail_str: str = vendor_detail.getvalue()
    detail.close()
    vendor_detail.close()

    print(detail_str)

    doc = DocxTemplate("template.docx")

    context = {'numbers': numbers,
               'vendor': vendors_raw[vendor_name],
               'detail': docxtpl.R(detail_str),
               'total': total,
               'vendor_all': vendor_all,
               'vendor_count': vendor_count,
               'vendor_detail': docxtpl.R(vendor_detail_str)
               }
    print(context)
    doc.render(context)
    doc.save(vendor_name + ".docx")

    vendor_product.to_excel(excel_writer, sheet_name=vendor_name)

excel_writer.save()
