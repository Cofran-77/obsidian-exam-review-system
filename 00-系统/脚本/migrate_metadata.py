"""One-time metadata migration that preserves every Markdown body verbatim."""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
SUBJECTS = {
    '马原': ('MY', '试题/1-马原-试题.pdf'),
    '毛中特': ('MT', '试题/2-毛中特-试题.pdf'),
    '新思想': ('XX', '试题/3-新思想-试题.pdf'),
    '史纲': ('SG', '试题/4-近现代史-试题.pdf'),
    '近现代史': ('SG', '试题/4-近现代史-试题.pdf'),
    '思修': ('SF', '试题/5-思修-试题.pdf'),
    '德法': ('SF', '试题/5-思修-试题.pdf'),
}


def read_frontmatter(path):
    text = path.read_text(encoding='utf-8')
    match = re.match(r'^---\s*\n(.*?)\n---\s*\n(.*)$', text, re.S)
    if not match:
        return {}, text, None
    values = {}
    for line in match.group(1).splitlines():
        if ':' in line and not line.startswith(' '):
            key, value = line.split(':', 1)
            values[key.strip()] = value.strip().strip("'\"")
    return values, match.group(2), match


def insert_fields(path, fields):
    text = path.read_text(encoding='utf-8')
    match = re.match(r'^(---\s*\n)(.*?)(\n---\s*\n.*)$', text, re.S)
    if not match:
        return False
    existing = set()
    for line in match.group(2).splitlines():
        if ':' in line and not line.startswith(' '):
            existing.add(line.split(':', 1)[0].strip())
    additions = [f'{key}: {value}' for key, value in fields.items() if key not in existing]
    if not additions:
        return False
    new_text = match.group(1) + match.group(2).rstrip() + '\n' + '\n'.join(additions) + match.group(3)
    path.write_text(new_text, encoding='utf-8')
    return True


def add_node_frontmatter(path, fields):
    text = path.read_text(encoding='utf-8')
    if text.startswith('---\n'):
        return insert_fields(path, fields)
    path.write_text('---\n' + '\n'.join(f'{key}: {value}' for key, value in fields.items()) + '\n---\n\n' + text, encoding='utf-8')
    return True


def title_to_id(directory):
    result = {}
    for path in directory.rglob('*.md'):
        meta, body, _ = read_frontmatter(path)
        if meta.get('title') and meta.get('id'):
            result[meta['title']] = meta['id']
        result[path.stem.replace('知识点·', '').replace('专题·', '')] = meta.get('id', '')
    return result


def main():
    kp_ids = title_to_id(ROOT)
    changed = 0
    for path in ROOT.rglob('*.md'):
        if path.parent.name != '02-错题' or '.obsidian' in path.parts:
            continue
        meta, body, _ = read_frontmatter(path)
        if not meta:
            continue
        subject = meta.get('subject', '马原')
        code, source_file = SUBJECTS.get(subject, ('XX', ''))
        error = meta.get('error_type', '')
        parts = error.split('/', 1)
        kp_match = re.search(r'\[\[([^\]|]+)', body)
        kp_title = kp_match.group(1).replace('知识点·', '') if kp_match else ''
        hub_match = re.search(r'thematic_hub:\s*[\'\"]?\[\[([^\]|]+)', path.read_text(encoding='utf-8'))
        hub_title = hub_match.group(1).replace('专题·', '') if hub_match else ''
        fields = {
            'source_exam': '题库未标注',
            'source_file': source_file,
            'source_page': 'null',
            'question_type': '单选',
            'target_kp_id': kp_ids.get(kp_title, ''),
            'thematic_hub_id': kp_ids.get(hub_title, ''),
            'error_category': parts[0] if parts else '',
            'error_subtype': parts[1] if len(parts) > 1 else parts[0],
            'review_result': '未复习',
        }
        if insert_fields(path, fields):
            changed += 1

    for path in ROOT.rglob('*.md'):
        if '.obsidian' in path.parts:
            continue
        meta, body, _ = read_frontmatter(path)
        if meta.get('id', '').startswith('KP-'):
            hub_match = re.search(r'所属核心专题[^\n]*\[\[([^\]|]+)', body)
            hub_title = hub_match.group(1).replace('专题·', '') if hub_match else ''
            changed += int(insert_fields(path, {
                'node_type': 'knowledge_point',
                'thematic_hub_id': kp_ids.get(hub_title, ''),
            }))
        elif meta.get('id', '').startswith('TH-'):
            changed += int(insert_fields(path, {'node_type': 'thematic_hub'}))

    for subject_dir in sorted(path for path in ROOT.iterdir() if path.is_dir() and re.match(r'0[1-5]-', path.name)):
        subject = subject_dir.name.split('-', 1)[1].split('·', 1)[0]
        code = SUBJECTS.get(subject, (subject[:2].upper(), ''))[0]
        for index, path in enumerate(sorted((subject_dir / '01-板块').glob('*.md')), start=1):
            changed += int(add_node_frontmatter(path, {
                'id': f'MOD-{code}-{index:02d}',
                'node_type': 'module',
                'subject': subject,
                'title': path.stem.replace('板块·', ''),
            }))
    print(f'migrated files: {changed}')


if __name__ == '__main__':
    main()
