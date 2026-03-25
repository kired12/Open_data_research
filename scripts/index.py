import pandas as pd
import re

df = pd.read_csv('file.csv')

def clean_and_split(text):
    text = str(text)
    match = re.search(r'\b\d{6}\b', text)
    if match:
        idx = match.group(0)
        addr = text.replace(idx, '')
        addr = re.sub(r'^[,\s]+|[,\s]+$', '', addr)
        addr = re.sub(r',\s*,', ',', addr).strip()
        return addr, idx
    return text, ""

df[['Адрес', 'Индекс']] = df['Адрес'].apply(lambda x: pd.Series(clean_and_split(x)))
df = df[['Название', 'Адрес', 'Индекс']]

df.to_csv('formatted.csv', index=False, encoding='utf-8-sig')