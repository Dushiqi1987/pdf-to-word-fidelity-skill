"""Create editable Word OMML, never formula screenshots.

Preferred backend: latex2mathml + Microsoft's locally installed MML2OMML.XSL.
Set MML2OMML_XSL or pass xsl_path; the proprietary stylesheet is not bundled.
Fallback: optional mathml2omml package (different conversion engine, so rendered
Word/PDF visual QA is mandatory). backend='python' explicitly exercises it.
Requires lxml and latex2mathml; returns lxml elements appendable to python-docx.
"""

from copy import deepcopy
import os
from pathlib import Path
import re

from lxml import etree as E

NS = {
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'ml': 'http://www.w3.org/1998/Math/MathML',
}


def q(prefix, name):
    return '{' + NS[prefix] + '}' + name


def find_office_xsl(xsl_path=None):
    """Return a real local stylesheet or None; explicit invalid paths fail."""
    explicit = xsl_path or os.environ.get('MML2OMML_XSL')
    if explicit:
        path = Path(explicit).expanduser()
        if not path.is_file():
            raise FileNotFoundError('MML2OMML stylesheet not found: ' + str(path))
        return path
    for root in filter(None, (os.environ.get('ProgramFiles'),
                              os.environ.get('ProgramFiles(x86)'),
                              'C:/Program Files', 'C:/Program Files (x86)')):
        for relative in ('Microsoft Office/root/Office16/MML2OMML.XSL',
                         'Microsoft Office/Office16/MML2OMML.XSL',
                         'Microsoft Office/Office15/MML2OMML.XSL'):
            path = Path(root) / relative
            if path.is_file():
                return path
    return None


def _run(text, size=24, upright=False):
    run = E.Element(q('m', 'r'))
    if upright:
        E.SubElement(E.SubElement(run, q('m', 'rPr')), q('m', 'sty'),
                     {q('m', 'val'): 'p'})
    props = E.SubElement(run, q('w', 'rPr'))
    E.SubElement(props, q('w', 'rFonts'), {
        q('w', 'ascii'): 'Cambria Math', q('w', 'hAnsi'): 'Cambria Math'})
    E.SubElement(props, q('w', 'sz'), {q('w', 'val'): str(size)})
    E.SubElement(run, q('m', 't')).text = text
    return run


def _force_below(omml, size):
    # Word can force n-ary limits sideways inside scalable delimiters even when
    # limLoc=undOvr. Explicit limit objects prevent that rendering regression.
    for operator in list(omml.iter(q('m', 'nary'))):
        props = operator.find(q('m', 'naryPr'))
        char = props.find(q('m', 'chr')) if props is not None else None
        symbol = char.get(q('m', 'val')) if char is not None else '\u2211'
        if symbol not in ('\u2211', '\u220f', '\u22c3', '\u22c2'):
            continue
        base = _run(symbol, size=int(round(size * 1.5)))
        for tag, wrapper_name in (('sub', 'limLow'), ('sup', 'limUpp')):
            bound = operator.find(q('m', tag))
            hidden = props.find(q('m', tag + 'Hide')) if props is not None else None
            if bound is None or len(bound) == 0 or (
                    hidden is not None and hidden.get(q('m', 'val'), '1') == '1'):
                continue
            wrapper = E.Element(q('m', wrapper_name))
            E.SubElement(wrapper, q('m', 'e')).append(base)
            limit = E.SubElement(wrapper, q('m', 'lim'))
            limit.extend(deepcopy(list(bound)))
            base = wrapper
        parent = operator.getparent()
        at = parent.index(operator)
        expression = operator.find(q('m', 'e'))
        parent.remove(operator)
        parent.insert(at, base)
        if expression is not None:
            for offset, child in enumerate(list(expression), 1):
                parent.insert(at + offset, child)


def latex_to_omml(latex, xsl_path=None, *, force_below=False,
                  backend='auto', font_size=12):
    """Return one native m:oMath element.

    backend: auto prefers local Office XSL, then optional Python fallback;
    office requires XSL; python requires mathml2omml. force_below explicitly
    stacks sum/product limits, including inside parentheses. font_size is pt.
    This converts already verified LaTeX, not PDF pixels or OCR by itself.
    """
    if not isinstance(latex, str) or not latex.strip():
        raise ValueError('A nonempty LaTeX expression is required.')
    if backend not in ('auto', 'office', 'python'):
        raise ValueError('backend must be auto, office, or python')
    if font_size <= 0:
        raise ValueError('font_size must be positive')
    from latex2mathml.converter import convert

    # Some Office XSL versions drop e/t from the operator form of \\det.
    latex = re.sub(r'\\det\b', lambda match: r'\mathrm{det}', latex)
    mathml = E.fromstring(convert(latex, display='block').encode('utf-8'))
    for node in mathml.iter():
        if not len(node) or node[0].tag != q('ml', 'mo'):
            continue
        symbol = ''.join(node[0].itertext())
        if symbol in ('\u2211', '\u220f', 'min', 'max'):
            if node.tag == q('ml', 'msub'):
                node.tag = q('ml', 'munder')
            elif node.tag == q('ml', 'msubsup'):
                node.tag = q('ml', 'munderover')
    xsl = find_office_xsl(xsl_path) if backend != 'python' else None
    if xsl is not None:
        result = E.XSLT(E.parse(str(xsl)))(mathml).getroot()
    else:
        if backend == 'office':
            raise RuntimeError('Office XSL unavailable; set MML2OMML_XSL.')
        try:
            from mathml2omml import convert as fallback
        except ImportError as error:
            raise RuntimeError('Install mathml2omml or provide MML2OMML_XSL.') from error
        fragment = fallback(E.tostring(mathml, encoding='unicode'))
        # mathml2omml 0.0.2 emits prefixed XML without namespace declarations.
        wrapper = E.fromstring(('<root xmlns:m="' + NS['m'] + '" xmlns:w="' +
                                NS['w'] + '">' + fragment + '</root>').encode())
        result = wrapper[0]
    omml = result if result.tag == q('m', 'oMath') else result.find('.//' + q('m', 'oMath'))
    if omml is None:
        raise ValueError('Converter returned no editable m:oMath element.')
    omml = deepcopy(omml)
    half_points = int(round(font_size * 2))
    for run in omml.iter(q('m', 'r')):
        props = run.find(q('w', 'rPr'))
        if props is None:
            props = E.Element(q('w', 'rPr'))
            index = 1 if run.find(q('m', 'rPr')) is not None else 0
            run.insert(index, props)
        for tag in ('rFonts', 'sz'):
            for child in props.findall(q('w', tag)):
                props.remove(child)
        E.SubElement(props, q('w', 'rFonts'), {
            q('w', 'ascii'): 'Cambria Math', q('w', 'hAnsi'): 'Cambria Math'})
        E.SubElement(props, q('w', 'sz'), {q('w', 'val'): str(half_points)})
    if force_below:
        _force_below(omml, half_points)
    return omml


def display_equation(latex, number=None, xsl_path=None, **kwargs):
    """Return m:oMathPara; optional number uses Word's native #(number) layout.

    Append to paragraph._p in python-docx. Number is an unparenthesized string
    such as '6.16'. Render in Word to verify available width and number layout.
    """
    math = latex_to_omml(latex, xsl_path, **kwargs)
    display = E.Element(q('m', 'oMathPara'), nsmap={'m': NS['m'], 'w': NS['w']})
    E.SubElement(E.SubElement(display, q('m', 'oMathParaPr')), q('m', 'jc'),
                 {q('m', 'val'): 'center'})
    display.append(math)
    if number is not None:
        number = str(number).strip()
        if not number or any(character in number for character in '\r\n()'):
            raise ValueError('number must be a nonempty label without parentheses/newlines')
        children = list(math)
        math[:] = []
        array = E.SubElement(math, q('m', 'eqArr'))
        E.SubElement(E.SubElement(array, q('m', 'eqArrPr')), q('m', 'maxDist'),
                     {q('m', 'val'): '1'})
        expression = E.SubElement(array, q('m', 'e'))
        expression.extend(children)
        size = int(round(kwargs.get('font_size', 12) * 2))
        expression.append(_run('#', size))
        delimiter = E.SubElement(expression, q('m', 'd'))
        props = E.SubElement(delimiter, q('m', 'dPr'))
        E.SubElement(props, q('m', 'begChr'), {q('m', 'val'): '('})
        E.SubElement(props, q('m', 'endChr'), {q('m', 'val'): ')'})
        E.SubElement(delimiter, q('m', 'e')).append(_run(number, size, upright=True))
    return display
