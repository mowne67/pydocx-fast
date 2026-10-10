# pydocx-fast

Drop-in performance fixes for [python-docx](https://github.com/python-openxml/python-docx). No API changes, byte-identical output.

```python
import pydocx_fast  # once, before use
from docx import Document
```

| 50,000 paragraphs | upstream | patched |
|---|---|---|
| write + save | 17.6s | 5.5s |
| load + read all text | 1.35s | 0.41s |

## What it fixes

1. `BaseOxmlElement.xpath()` recompiled the XPath string on every call (~80% of read time). Now compiled once and cached.
2. `first_child_found_in()` scanned every body child to find the trailing `w:sectPr`, making `add_paragraph` O(n) and document building O(n²). Now walks back from the last child. Falls back to the upstream scan for non-schema-ordered elements.

## Test

```sh
uv run python test_pydocx_fast.py
```

Runs upstream and patched in separate processes and asserts identical output.
