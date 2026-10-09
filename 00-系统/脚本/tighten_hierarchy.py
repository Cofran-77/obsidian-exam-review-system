"""Keep the graph hierarchical while preserving board-page knowledge indexes."""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]


def read_meta(path):
    text = path.read_text(encoding='utf-8')
    match = re.match(r'^---\s*\n(.*?)\n---\s*\n(.*)$', text, re.S)
    data = {}
    body = text
    if match:
        body = match.group(2)
        for line in match.group(1).splitlines():
            if ':' in line and not line.startswith(' '):
                key, value = line.split(':', 1)
                data[key.strip()] = value.strip().strip("'\"")
    return data, body


def replace_board_knowledge_section(path, title):
    text = path.read_text(encoding='utf-8')
    start = text.find('## 🧠 本板块核心知识点')
    if start < 0:
        return False
    end = text.find('\n## ', start + 5)
    if end < 0:
        end = text.find('\n---', start + 5)
    if end < 0:
        end = len(text)
    section = f'''## 🧠 本板块核心知识点\n\n```dataview\nTABLE file.link as "知识点", importance as "重要度"\nFROM ""\nWHERE node_type = "knowledge_point" AND module = "{title}"\nSORT file.name ASC\n```\n'''
    new_text = text[:start] + section + text[end:]
    if new_text != text:
        path.write_text(new_text, encoding='utf-8')
        return True
    return False


def flatten_kp_board_links(path, title):
    text = path.read_text(encoding='utf-8')
    original = text
    text = re.sub(r'(?m)^> \*\*所属板块\*\*：\[\[板块·([^\]|]+)(?:\|[^\]]+)?\]\]$', r'> **所属板块**：\1', text)
    text = re.sub(r'(?m)^\[\[板块·[^\n]+\]\] · \[\[[^\n]+\]\]\s*$', f'> 所属板块：{title}', text)
    return text, text != original


def main():
    changed = 0
    for path in ROOT.rglob('01-板块/*.md'):
        if '.obsidian' in path.parts:
            continue
        data, _ = read_meta(path)
        changed += int(replace_board_knowledge_section(path, data.get('title', path.stem.replace('板块·', ''))))

    for path in ROOT.rglob('03-知识点/*.md'):
        if '.obsidian' in path.parts:
            continue
        data, _ = read_meta(path)
        text, did_change = flatten_kp_board_links(path, data.get('module', ''))
        if did_change:
            path.write_text(text, encoding='utf-8')
            changed += 1
    print(f'tightened hierarchy files: {changed}')


if __name__ == '__main__':
    main()
