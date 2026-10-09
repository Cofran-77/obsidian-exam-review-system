"""Restore the intended political-theory hierarchy in note links."""
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


def remove_outline_topic_links(text):
    # Subject outlines own their boards. Topic links belong on board pages.
    text = re.sub(
        r'(?ms)^## 🌐 .*?\n.*?(?=^## 📌 )',
        '## 🌐 专题分布\n\n> 各核心专题由所属板块承载，进入对应板块页面查看。\n\n',
        text,
    )
    text = re.sub(
        r'(?m)^- \[\[专题·[^\n]+\]\].*\n',
        lambda m: '> 专题入口已归入所属板块页面。\n' if '00-马原知识大纲' in m.string else m.group(0),
        text,
    )
    return text


def add_module_knowledge_links(path, meta, all_kps):
    text = path.read_text(encoding='utf-8')
    if '## 🧠 本板块核心知识点' in text:
        return False
    kps = [p for p, data in all_kps if data.get('subject') == meta.get('subject') and data.get('module') == meta.get('title')]
    if not kps:
        return False
    lines = ['## 🧠 本板块核心知识点']
    for kp in kps:
        lines.append(f'- [[{kp.stem}]]')
    section = '\n'.join(lines) + '\n\n'
    marker = '## 🚨 本板块对应错题'
    if marker in text:
        text = text.replace(marker, section + marker, 1)
    else:
        footer = '\n---\n'
        text = text.replace(footer, '\n' + section + footer, 1) if footer in text else text.rstrip() + '\n\n' + section
    path.write_text(text, encoding='utf-8')
    return True


def main():
    changed = 0
    all_kps = []
    for path in ROOT.rglob('*.md'):
        if '.obsidian' in path.parts or '00-系统' in path.parts:
            continue
        if '03-知识点' in path.parts:
            data, _ = read_meta(path)
            all_kps.append((path, data))

    for path in ROOT.rglob('00-*知识大纲.md'):
        if '.obsidian' in path.parts:
            continue
        text = path.read_text(encoding='utf-8')
        new_text = remove_outline_topic_links(text)
        if new_text != text:
            path.write_text(new_text, encoding='utf-8')
            changed += 1

    for path in ROOT.rglob('01-板块/*.md'):
        if '.obsidian' in path.parts:
            continue
        data, _ = read_meta(path)
        changed += int(add_module_knowledge_links(path, data, all_kps))
    print(f'repaired hierarchy files: {changed}')


if __name__ == '__main__':
    main()
