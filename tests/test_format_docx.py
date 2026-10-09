import unittest
import tempfile
from pathlib import Path
from datetime import date
from unittest.mock import patch
import format_docx as m

class WorkflowTests(unittest.TestCase):
    def test_frontmatter(self):
        meta, body = m.parse_yaml_frontmatter('---\nid: MY-001\ntags:\n  - 示例\n---\n# 原创示例\n')
        self.assertEqual(meta['tags'], ['示例'])
        self.assertIn('# 原创示例', body)
    def test_defaults(self):
        self.assertEqual(m.normalize_metadata({'id':'MY-001'})['review_count'], 0)
    def test_invalid_metadata(self):
        cards = [m.normalize_metadata({'id':'MY-001'}), m.normalize_metadata({'id':'MY-001','next_review':'bad','target_kp_id':'bad'})]
        errors = m.validate_cards(cards)
        self.assertTrue(any('重复 id' in e for e in errors))
        self.assertTrue(any('日期格式' in e for e in errors))
        self.assertTrue(any('ID 格式' in e for e in errors))
    def test_due(self):
        self.assertTrue(m.is_due({'status':'待复习','next_review':'2026-10-09'}, date(2026,10,9)))
        self.assertFalse(m.is_due({'status':'已掌握','next_review':'2026-10-09'}, date(2026,10,9)))
    def test_discovery(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root/'02-错题').mkdir()
            (root/'02-错题'/'card.md').write_text('test', encoding='utf-8')
            (root/'note.md').write_text('not a card', encoding='utf-8')
            self.assertEqual(len(m.discover_question_files(root)), 1)
    def test_original_example(self):
        questions = m.load_all_questions()
        self.assertTrue(any(q['q_num']==1 and q['subject']=='马原' for q in questions))
    def test_compile_outputs(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(m, 'WORKSPACE_DIR', tmp):
            questions = m.load_all_questions()
            m.compile_master_markdown(questions)
            m.build_docx_handout(questions)
            content = (Path(tmp)/'错题本.md').read_text(encoding='utf-8')
            self.assertIn('原创演示', content)
            from docx import Document
            doc = Document(Path(tmp)/'错题本.docx')
            self.assertGreater(len(doc.paragraphs), 0)

if __name__ == '__main__':
    unittest.main()

