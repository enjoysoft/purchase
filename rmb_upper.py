import re


def num2chn(num):
    """
    将数字（数字字符串）转换为人民币大写
    :param num: int、float或者str
    :return: 返回处理结果，发生错误返回None
    """
    chr = ('零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖')
    bit = ('分', '角', '元', '拾', '佰', '仟', '万', '拾', '佰', '仟', '亿', '拾', '佰', '仟', '万')
    # 例 324562.003
    # 转换为 00265423
    try:
        num = float(num)
        # 将数字反序排列，从右至左依次
        num_str = ('%0.2f' % num).replace('.', '')[::-1]
    except ValueError:
        return None
    n = len(num_str)
    if n >= 15:
        return None
    result = []
    for i in range(0, n):
        # 非圆、万、亿，
        if num_str[i] == "0" and i != 2 and i != 6 and i != 10:
            s = chr[0]
        elif num_str[i] == "0" and (i == 2 or i == 6 or i == 10):
            s = bit[i]
        else:
            s = bit[i] + chr[int(num_str[i])]
        result.append(s)
    # 从左值右重排字符串
    rst = "".join(result)[::-1]
    # 去零，中间多个零合并为一个，尾部零全去， "零" == chr[0]
    # r = r"[零]+"
    r = "[" + chr[0] + "]+"
    rst = re.compile(r).sub(chr[0], rst).rstrip(chr[0])
    # 去零元、零万、零亿
    for i in [2, 6, 10]:
        rst = rst.replace(chr[0] + bit[i], bit[i])
    if rst[-1] == bit[2]:
        rst += "整"
    print('%0.2f' % num)
    print(rst)
    return rst
