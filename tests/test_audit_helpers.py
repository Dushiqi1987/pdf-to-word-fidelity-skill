"""Structural audit regression tests; run with unittest discovery."""
import sys
from pathlib import Path
import unittest

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from audit_docx import NS, equation_labels, unresolved_pagerefs


def document(body):
    return etree.fromstring(('<w:document xmlns:w="' + NS['w'] + '" xmlns:m="' + NS['m'] + '"><w:body>' + body + '</w:body></w:document>').encode())


class AuditHelpersTest(unittest.TestCase):
    def test_native_delimiter_label(self):
        root = document('<w:p><m:oMath><m:r><m:t>x</m:t></m:r><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:t>6.16</m:t></m:r></m:e></m:d></m:oMath></w:p>')
        self.assertEqual(equation_labels(root), {'6.16'})

    def test_plain_suffix_and_body_false_positive(self):
        root = document('<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath><w:r><w:t>(7.2)</w:t></w:r></w:p><w:p><w:r><w:t>See (7.99)</w:t></w:r></w:p>')
        self.assertEqual(equation_labels(root), {'7.2'})

    def test_word_normalized_default_parentheses(self):
        root = document('<w:p><m:oMath><m:d><m:dPr/><m:e><m:r><m:t>6.21</m:t></m:r></m:e></m:d></m:oMath></w:p>')
        self.assertEqual(equation_labels(root), {'6.21'})

    def test_simple_and_complex_unresolved_fields(self):
        root = document('<w:p><w:fldSimple w:instr=" PAGEREF test "><w:r><w:t>?</w:t></w:r></w:fldSimple></w:p><w:p><w:r><w:instrText> PAGEREF other </w:instrText></w:r><w:r><w:t>?</w:t></w:r></w:p>')
        self.assertEqual(unresolved_pagerefs(root), 2)

    def test_resolved_simple(self):
        root = document('<w:p><w:fldSimple w:instr=" PAGEREF test "><w:r><w:t>12</w:t></w:r></w:fldSimple></w:p>')
        self.assertEqual(unresolved_pagerefs(root), 0)


if __name__ == '__main__':
    unittest.main()
