import io

import docxtpl
import pandas as pd
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from rmb_upper import num2chn

vendor_label: str = '供应商'
quote_label: str = '总价'
index_label: str = '序号'
formatting: str = '%.2f'
unit_label: str = '单价'
no_quote: str = '无法报价'
excel_filename: str = 'bay.xlsx'
contract: str = 'BN-2018-020'

df: pd.DataFrame = pd.read_excel(excel_filename, sheet_name='quote', na_values=no_quote, index_col=0)

vendors_raw: pd.Series = pd.read_excel(excel_filename, sheet_name='vendor', header=None, index_col=0, squeeze=True)
vendors_name = vendors_raw.index
vendors = list(filter(lambda v: v in df, vendors_raw.index))
vendors_timeout: list = list(filter(lambda v: v not in df.columns, vendors_raw.index))
vendor_invalid = list(filter(lambda v: v not in vendors_raw.index, df.columns[8:]))
if vendor_invalid:
    raise ValueError("Invalid vendor name found! " + str(vendor_invalid))

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

vendor_all = len(vendors_raw)
vendor_count = len(quotes.columns[~quotes.isna().all()])

book = load_workbook(excel_filename)
excel_writer = pd.ExcelWriter(excel_filename, engine='openpyxl')
excel_writer.book = book
excel_writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

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

    numbers = ','.join(sorted(map(str, vendor_product['采购依据'].unique())))
    print(numbers)

    items = ','.join(sorted(vendor_product['项目'].unique()))
    print(items)

    vendor_quotes = vendor_df.loc[:, vendors]

    quote_groups = vendor_quotes.notna().groupby(vendors).groups
    quote_detail = [(v.tolist(), vendor_quotes.loc[v].sum(skipna=False).sort_values()) for k, v in
                    quote_groups.items()]

    detail: io.StringIO = io.StringIO()
    vendor_detail: io.StringIO = io.StringIO()

    vendor_full_name = vendors_raw[vendor_name]

    for index, detail_df in quote_detail:
        print("询价单第", ','.join(map(str, index)), '项', sep='', file=detail)
        print("询价单第", ','.join(map(str, index)), '项共有', detail_df.count(), '家供应商', sep='', end='。', file=vendor_detail)
        for detail_vendor, row in detail_df.items():
            if pd.isna(row):
                print(vendors_raw[detail_vendor], '回复无法报价', sep='', file=detail)
            else:
                print(vendors_raw[detail_vendor], '报价人民币', '%.2f' % row, '元', sep='', file=detail)
        for v in vendors_timeout:
            print(vendors_raw[v], '逾期未回复', sep='', file=detail)

    detail_str: str = detail.getvalue()
    vendor_detail_str: str = vendor_detail.getvalue()
    detail.close()
    vendor_detail.close()

    print(detail_str)

    doc = DocxTemplate("template.docx")

    context = {'numbers': numbers,
               'items': items,
               'vendor': vendor_full_name,
               'detail': docxtpl.R(detail_str),
               'total': total,
               'vendor_all': vendor_all,
               'vendor_count': vendor_count,
               'vendor_detail': vendor_detail_str,
               'contract': contract
               }
    doc.render(context)
    doc.save(vendor_name + ".docx")

    sheet_name = '订单 ' + vendor_name
    if sheet_name in book.sheetnames:
        book.remove(book[sheet_name])

    ws = book.copy_worksheet(book['template'])
    ws.title = sheet_name
    ws['I4'] = vendor_full_name
    ws['K3'] = numbers
    ws['k2'] = contract
    ws['H7'] = vendor_df[quote_label].sum()
    ws['D8'] = "%s（￥%s）（含16%%增值税）" % (num2chn(vendor_df[quote_label].sum()), total)

    columns = ['名称', '规格', '规范', '数量', '单位', '单价', '总价', '交货周期', '采购依据', '申请部门', '项目']
    excel_vendor_detail = vendor_product.reindex(columns=columns, fill_value='')
    for r in dataframe_to_rows(excel_vendor_detail, index=True, header=False):
        ws.append(r)

excel_writer.save()
