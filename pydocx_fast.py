"""Drop-in speedups for python-docx. `import pydocx_fast` once, then use `docx` as usual."""
from lxml import etree

from docx.oxml.ns import nsmap, qn
from docx.oxml.xmlchemy import BaseOxmlElement

_compiled: dict[str, etree.XPath] = {}


def xpath(self, xpath_str):
    # Upstream recompiles the expression on every call; that was ~80% of read time.
    xp = _compiled.get(xpath_str)
    if xp is None:
        xp = _compiled[xpath_str] = etree.XPath(xpath_str, namespaces=nsmap)
    return xp(self)


_original_first_child_found_in = BaseOxmlElement.first_child_found_in


def first_child_found_in(self, *tagnames):
    # Successors sit at the end of the parent in schema order (e.g. w:sectPr closes w:body),
    # so walk back from the last child instead of scanning every child: O(n^2) -> O(1) appends.
    # ponytail: assumes schema-ordered children; falls back to the upstream forward scan otherwise.
    tags = {qn(t) for t in tagnames}
    found = None
    child = next(self.iterchildren(reversed=True), None)  # len(self) is O(n) in lxml
    while child is not None and child.tag in tags:
        found, child = child, child.getprevious()
    return found if found is not None else _original_first_child_found_in(self, *tagnames)


BaseOxmlElement.xpath = xpath
BaseOxmlElement.first_child_found_in = first_child_found_in
