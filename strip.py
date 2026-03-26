import re

def strip_comments(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        text = f.read()

    # HTML
    text = re.sub(r'<!--[\s\S]*?-->', '', text)
    # Block comments
    text = re.sub(r'/\*[\s\S]*?\*/', '', text)
    # Single line comments not preceded by colon (to save http://)
    text = re.sub(r'(?<!:)//.*', '', text)

    # remove excess blank lines
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        if line.strip() == '':
            if cleaned_lines and cleaned_lines[-1].strip() == '':
                continue
            else:
                cleaned_lines.append('')
        else:
            cleaned_lines.append(line)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write('\n'.join(cleaned_lines))

strip_comments('index.html')
strip_comments('api.js')
