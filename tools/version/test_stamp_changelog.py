"""Unit tests for stamp_changelog.py.

Run: python3 -m unittest discover -s tools/version -p "test_*.py"
(the stamp-version workflow runs this before it stamps anything).
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import stamp_changelog as sc  # noqa: E402


def doc(*headings):
    return '\n'.join(h + '\n\n- body\n' for h in headings)


class StampTest(unittest.TestCase):

    def test_patch_minor_major_and_default(self):
        cases = [
            ('(vNEXT patch)', 'v2.42.1'),
            ('(vNEXT minor)', 'v2.43.0'),
            ('(vNEXT major)', 'v3.0.0'),
            ('(vNEXT)', 'v2.42.1'),
            ('(vnext Minor)', 'v2.43.0'),
        ]
        for tag, want in cases:
            text = doc('## 2026-10-29 - new ' + tag, '## 2026-10-28 - old (v2.42.0)')
            out, assigned, _ = sc.stamp(text)
            self.assertEqual([want], assigned, tag)
            self.assertIn('## 2026-10-29 - new (' + want + ')', out)

    def test_several_pending_lowest_first(self):
        # One PR carried two entries (like #214: v2.39.0 then v2.40.0).
        text = doc('## b - second (vNEXT minor)',
                   '## a - first (vNEXT patch)',
                   '## z - base (v2.38.13)')
        out, assigned, _ = sc.stamp(text)
        self.assertEqual(['v2.38.14', 'v2.39.0'], assigned)
        self.assertIn('## b - second (v2.39.0)', out)
        self.assertIn('## a - first (v2.38.14)', out)

    def test_base_is_highest_not_topmost(self):
        text = doc('## x (vNEXT)', '## y (v2.40.0)', '## z (v2.41.0)')
        _, assigned, _ = sc.stamp(text)
        self.assertEqual(['v2.41.1'], assigned)

    def test_nothing_pending_is_unchanged(self):
        text = doc('## a (v1.0.0)')
        out, assigned, _ = sc.stamp(text)
        self.assertEqual(text, out)
        self.assertEqual([], assigned)

    def test_only_the_trailing_tag_is_replaced(self):
        text = doc('## a - keeps (vNEXT) in the title (vNEXT minor)', '## b (v1.2.3)')
        out, _, _ = sc.stamp(text)
        self.assertIn('## a - keeps (vNEXT) in the title (v1.3.0)', out)

    def test_non_heading_and_fenced_lines_ignored(self):
        text = ('## a (vNEXT)\n\n'
                'body mentions (vNEXT) and (v9.9.9)\n'
                '```markdown\n## example (v9.0.0)\n## example (vNEXT)\n```\n\n'
                '### Added (vNEXT)\n\n'
                '## b (v1.0.0)\n')
        out, assigned, _ = sc.stamp(text)
        self.assertEqual(['v1.0.1'], assigned)
        self.assertIn('## example (vNEXT)', out)
        self.assertIn('### Added (vNEXT)', out)

    def test_duplicates_reported_except_known(self):
        text = doc('## a (v2.41.2)', '## b (v2.41.2)', '## c (v2.41.1)',
                   '## d (v2.41.1)', '## e (v2.40.0)', '## f (v2.40.0)')
        _, _, dups = sc.stamp(text)
        self.assertEqual(['v2.40.0'], dups)

    def test_hand_written_number_colliding_with_stamp_is_reported(self):
        # A branch cut before the switch still writes a real number by hand.
        text = doc('## new (vNEXT)', '## hand (v2.42.1)', '## base (v2.42.0)')
        _, assigned, dups = sc.stamp(text)
        self.assertEqual(['v2.42.2'], assigned)
        self.assertEqual([], dups)
        text = doc('## hand (v2.42.1)', '## other (v2.42.1)')
        _, _, dups = sc.stamp(text)
        self.assertEqual(['v2.42.1'], dups)

    def test_main_rewrites_file_and_keeps_crlf(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, 'CHANGELOG.md')
            outp = os.path.join(d, 'out')
            with open(path, 'w', encoding='utf-8', newline='') as f:
                f.write('## a \u65e5\u672c (vNEXT minor)\r\n\r\n## b (v1.0.0)\r\n')
            sc.main([path, '--github-output', outp])
            with open(path, encoding='utf-8', newline='') as f:
                self.assertEqual('## a \u65e5\u672c (v1.1.0)\r\n\r\n## b (v1.0.0)\r\n', f.read())
            with open(outp, encoding='utf-8') as f:
                self.assertEqual('versions=v1.1.0\n', f.read())

    def test_real_changelog_parses(self):
        root = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')
        with open(os.path.join(root, 'CHANGELOG.md'), encoding='utf-8') as f:
            text = f.read()
        out, assigned, dups = sc.stamp(text)
        self.assertEqual([], dups)
        if not assigned:
            self.assertEqual(text, out)


if __name__ == '__main__':
    unittest.main()
