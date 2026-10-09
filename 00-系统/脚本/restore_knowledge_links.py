"""Restore reading/navigation links while keeping the graph view focused."""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
SUBJECT_OUTLINES = {
    '马原': '00-马原知识大纲',
    '毛中特': '00-毛中特与习思想知识大纲',
    '新思想': '00-毛中特与习思想知识大纲',
    '史纲': '00-史纲知识大纲',
    '近现代史': '00-史纲知识大纲',
    '德法': '00-德法知识大纲',
    '思修': '00-德法知识大纲',
    '当代': '00-当代与时政知识大纲',
}


def frontmatter(text):
    match = re.match(r'^---\s*\n(.*?)\n---\s*\n(.*)$', text, re.S)
    if not match:
        return {}, text
    data = {}
    for line in match.group(1).splitlines():
        if ':' in line and not line.startswith(' '):
            key, value = line.split(':', 1)
            data[key.strip()] = value.strip().strip("'\"")
    return data, match.group(2)


def title_map():
    result = {}
    for path in ROOT.rglob('*.md'):
        if '.obsidian' in path.parts or '00-系统' in path.parts:
            continue
        data, _ = frontmatter(path.read_text(encoding='utf-8'))
        if data.get('id'):
            result[data['id']] = path.stem
        if data.get('title'):
            result[data['title']] = path.stem
        result[path.stem] = path.stem
    return result


def wiki(value):
    return value if value.startswith('[[') else f'[[{value}]]'


def restore_wrong_cards(path, data, text, names):
    kp = names.get(data.get('target_kp_id', ''))
    hub = names.get(data.get('thematic_hub_id', ''))
    if kp:
        text = re.sub(r"(?m)^target_kp:.*$", f"target_kp: '[[{kp}]]'", text)
    if hub:
        text = re.sub(r"(?m)^thematic_hub:.*$", f"thematic_hub: '[[{hub}]]'", text)
    return text


def restore_knowledge_page(path, data, text, names):
    hub_name = names.get(data.get('thematic_hub_id', ''))
    module_name = data.get('module', '')
    if module_name:
        text = re.sub(r'(?m)^> \*\*所属板块\*\*：\[\[板块·([^\]|]+)(?:\|[^\]]+)?\]\]$', r'> **所属板块**：\1', text)
    if hub_name:
        text = re.sub(
            r'(?m)^> \*\*所属核心专题\*\*：.*$',
            f'> **所属核心专题**：[[{hub_name}]]',
            text,
        )
        text = re.sub(
            r'(?m)^- 🏛️ \*\*所属专题总揽\*\*：.*$',
            f'- 🏛️ **所属专题总揽**：[[{hub_name}]]',
            text,
        )
    # Board membership is kept as metadata/plain text. The graph reaches a
    # knowledge point through its board's thematic hubs.
    return text


def restore_module_page(path, data, text, cards, names):
    # Rebuild the useful preview section that was removed from board pages.
    if '## 🚨 本板块对应错题' not in text and '## 🧭 本板块错题入口' in text:
        lines = ['## 🚨 本板块对应错题（点击进入错题深度复盘）']
        for card_path, card_data in cards:
            if card_data.get('subject') == data.get('subject') and card_data.get('module') == data.get('title'):
                lines.append(f'- [[{card_path.stem}]]')
        section = '\n'.join(lines) + '\n'
        text = text.replace(
            '## 🧭 本板块错题入口\n\n> 错题通过知识点节点反向汇入；本页只保留板块与专题关系。\n',
            section,
        )
    if '返回马原大纲' not in text:
        outline = SUBJECT_OUTLINES.get(data.get('subject', '马原'))
        if outline:
            text = text.rstrip() + f'\n---\n[[{outline}|⬅️ 返回学科大纲]]\n'
    return text


def restore_hub_page(path, data, text):
    if '返回所属板块' not in text:
        module = data.get('subject', '马原')
        # Infer the board from the hub title when possible.
        board = '马克思主义政治经济学' if '政治经济学' in data.get('title', '') else ''
        for candidate in ['导论与唯物论', '唯物辩证法', '认识论', '唯物史观（历史唯物主义）', '科学社会主义']:
            if candidate in data.get('title', ''):
                board = candidate
        if board:
            text = text.rstrip() + f'\n---\n[[板块·{board}|⬅️ 返回所属板块]] · [[00-马原知识大纲|🏠 返回马原大纲]]\n'
    return text


def main():
    names = title_map()
    cards = []
    for path in ROOT.rglob('*.md'):
        if '02-错题' in path.parts:
            data, _ = frontmatter(path.read_text(encoding='utf-8'))
            cards.append((path, data))

    changed = 0
    for path in ROOT.rglob('*.md'):
        if '.obsidian' in path.parts or '00-系统' in path.parts:
            continue
        text = path.read_text(encoding='utf-8')
        original = text
        data, body = frontmatter(text)
        if '02-错题' in path.parts:
            text = restore_wrong_cards(path, data, text, names)
        elif '03-知识点' in path.parts:
            text = restore_knowledge_page(path, data, text, names)
        elif '01-板块' in path.parts:
            text = restore_module_page(path, data, text, cards, names)
        elif '04-核心专题枢纽' in path.parts:
            text = restore_hub_page(path, data, text)
        if text != original:
            path.write_text(text, encoding='utf-8')
            changed += 1
    print(f'restored knowledge links in files: {changed}')


if __name__ == '__main__':
    main()
