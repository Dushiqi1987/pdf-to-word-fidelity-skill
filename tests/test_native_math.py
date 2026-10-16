"""Structural regression tests. Rendered Word QA remains required separately."""

from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from native_math import NS, display_equation, find_office_xsl, latex_to_omml, q


class NativeMathTests(unittest.TestCase):
    def test_fraction_editable(self):
        root = latex_to_omml(r'\frac{x+1}{y}')
        self.assertEqual(root.tag, q('m', 'oMath'))
        self.assertEqual(len(root.findall('.//' + q('m', 'f'))), 1)
        self.assertNotIn('drawing', str(root.tag))

    def test_det_keeps_all_letters(self):
        root = latex_to_omml(r'\log\left|\det\frac{\partial f(x)}{\partial x}\right|')
        text = ''.join(root.xpath('.//m:t/text()', namespaces=NS))
        self.assertIn('det', text)

    def test_force_below_inside_delimiter(self):
        root = latex_to_omml(r'\sigma\left(x+\sum_{j\in N(i)}a_j\right)', force_below=True)
        self.assertTrue(root.findall('.//' + q('m', 'limLow')))
        self.assertFalse(root.findall('.//' + q('m', 'nary')))
        self.assertIn('a', ''.join(root.xpath('.//m:t/text()', namespaces=NS)))

    def test_two_limits_preserved(self):
        root = latex_to_omml(r'\sum_{j=1}^{K}x_j', force_below=True)
        self.assertTrue(root.findall('.//' + q('m', 'limLow')))
        self.assertTrue(root.findall('.//' + q('m', 'limUpp')))
        text = ''.join(root.xpath('.//m:t/text()', namespaces=NS))
        self.assertIn('K', text)
        self.assertIn('1', text)

    def test_numbered_display(self):
        root = display_equation(r'x=1', number='6.16')
        self.assertEqual(root.tag, q('m', 'oMathPara'))
        self.assertEqual(len(root.findall('.//' + q('m', 'oMath'))), 1)
        self.assertIn('#', root.xpath('.//m:t/text()', namespaces=NS))
        self.assertIn('6.16', root.xpath('.//m:t/text()', namespaces=NS))

    def test_fallback_backend(self):
        try:
            import mathml2omml
        except ImportError:
            self.skipTest('optional mathml2omml not installed')
        root = latex_to_omml(r'\frac{x}{y}+\sum_{i=1}^{N}x_i+\det A', backend='python')
        self.assertEqual(root.tag, q('m', 'oMath'))
        self.assertIn('det', ''.join(root.xpath('.//m:t/text()', namespaces=NS)))
        self.assertTrue(root.findall('.//' + q('m', 'f')))

    def test_explicit_missing_stylesheet_fails(self):
        with self.assertRaises(FileNotFoundError):
            find_office_xsl(Path(__file__).parent / 'not-an-office-stylesheet.xsl')

    def test_invalid_input_fails(self):
        for value in ('', ' '):
            with self.assertRaises(ValueError):
                latex_to_omml(value)
        with self.assertRaises(ValueError):
            latex_to_omml('x', backend='screenshots')
        with self.assertRaises(ValueError):
            display_equation('x', number='(1)')


if __name__ == '__main__':
    unittest.main()
