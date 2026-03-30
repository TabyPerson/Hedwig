def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))

def filter_ignored(series, ignore_sentences):
    ignore_lower = [s.strip().lower() for s in ignore_sentences]
    return series[~series.str.strip().str.lower().isin(ignore_lower)]

def split_and_flatten(series):
    return (
        series.dropna()
        .astype(str)
        .str.replace('\r', '\n')
        .str.replace(';', '\n')
        .str.replace(',', '\n')
        .str.split('\n')
        .explode()
        .str.strip()
        .loc[lambda x: x != '']
    )

def extract_paths(text):
    exts = r"mp4|avi|mov|wmv|mkv|flv|webm|mpeg|mpg|pdf|docx|xlsx|txt|jpg|png|csv"
    pattern = rf'(\\\\[^\n\r,;]+?\.(?:{exts}))|([A-Za-z]:\\[^\n\r,;]+?\.(?:{exts}))|(/[^ \n\r,;]+?\.(?:{exts}))'
    matches = re.findall(pattern, text, re.IGNORECASE)
    paths = [m[0] or m[1] or m[2] for m in matches if any(m)]
    return [p.strip() for p in paths]

def extract_path_after_dot(text):
    pattern = r'([A-Za-z]:\\[^\s,;]+(?:\.[a-zA-Z0-9]+)?)|(/[^ \n\r\t,;]+(?:\.[a-zA-Z0-9]+)?)'
    matches = re.findall(pattern, text)
    paths = [m[0] if m[0] else m[1] for m in matches if m[0] or m[1]]
    return paths