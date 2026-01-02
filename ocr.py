from cnocr import CnOcr


import pandas as pd



# 读文件，保留第一列（索引 0），如果有表头默认会把第一行当列名
df = pd.read_excel("渠道明细.xlsx", usecols=[0])


# 如果文件第一行是表头并想跳过表头：
first_col_values = df.iloc[:, 0].dropna().tolist()  # 根据需要 dropna 或者 [1:]

# 额外的channel
first_col_values += ['二门诊']

# print(first_col_values)

img_fp = 'image.jpg'
ocr = CnOcr()  # 所有参数都使用默认值
out = ocr.ocr(img_fp)

# print(out)
# 过滤置信度低于 0.4 的结果
out = [ item for item in out if item['score'] >= 0.4 ]

# 匹配数字开头的
out = [ item for item in out if item['text'][0].isdigit() ]

# 匹配中文结尾的
out = [ item for item in out if item['text'][-1] >= '\u4e00' and item['text'][-1] <= '\u9fa5' ]

# 提取文本内容  
text = [ item['text'] for item in out ]
# print(text)

# 解析文本内容成结构化数据
contacters = []
for line in text:
    item = {}
    # 分离 数字小数点 和中文
    # 从左到右找到第一个非数字非小数点字符的位置，最多不超过8个字符
    split_index = 0
    for i, char in enumerate(line):
        if i >= 8:
            split_index = i
            break
        if not (char.isdigit() or char == '.'):
            split_index = i
            break
    num_part = line[:split_index]
    chinese_part = line[split_index:]
    # print(f"数字部分: {num_part}, 中文部分: {chinese_part}")
    item['日期'] = num_part
    # 把中文部分分成三部分 split_word 为分隔符，split_word 是firt_col_values中的某一个值
    for split_word in first_col_values:
        if split_word in chinese_part:
            parts = chinese_part.split(split_word)
            if len(parts) == 2:
                part1 = parts[0]
                part2 = split_word
                part3 = parts[1]
                # print(f"分割结果: 部分1: {part1}, 分隔符: {part2}, 部分3: {part3}")
                item['微信名'] = part1
                item['渠道'] = part2
                item['备注/意向'] = part3
                break

    contacters.append(item)

for contact in contacters:
    print(contact)