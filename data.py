import pandas as pd


vendor_label: str = '供应商'
quote_label: str = '报价'
index_label: str = '序号'
formatting: str = '%.2f'
unit_label: str = '单价'

vendors: list = ['科罕', '中航供销', '美隆航', '韦斯科', '屹领', '赛方', '润洽']

df: pd.DataFrame = pd.read_clipboard(sep='\t', na_values='无法报价')
df = df.set_index(index_label)

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

for vendor_name, vendor_df in groups:
    print('=============================')
    print(vendor_name)
    print('=============================')

    vendor_product = vendor_df.drop(vendors, 1)
    print('总价', formatting % vendor_df[quote_label].sum())
    print(vendor_product.to_csv(sep='\t', float_format=formatting))

    vendor_quotes = vendor_df.loc[:, vendors]
    quote_count = vendor_quotes.count(1)
    (many, few) = (vendor_quotes.loc[quote_count >= 3], vendor_quotes.loc[quote_count < 3])

    few_detail = [(row[0], row[1].dropna().sort_values()) for row in few.iterrows()]

    if not few.empty:
        print(">>> 过少报价")
    for index, few_df in few_detail:
        print("序号:", index)
        print(few_df.to_csv(sep='\t', float_format=formatting))

    many_notna = many.notna()
    many_groups = many_notna.groupby(vendors).groups
    many_detail = [(v.tolist(), many.loc[v].sum(skipna=False).dropna().sort_values()) for k, v in many_groups.items()]

    if not many.empty:
        print('>>> 其它报价')
    for index, many_df in many_detail:
        print("序号:", ','.join(map(str, index)))
        print(many_df.to_csv(sep='\t', float_format=formatting))
