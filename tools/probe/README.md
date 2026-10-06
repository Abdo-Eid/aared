# tools/probe

Ground-truth probes: reverse-engineering real Office behavior by
interacting with it (AGENTS.md "Ground-truth probe"). Scratch evidence,
not a deliverable — each finished probe is deleted once its finding is
codified as a comment plus a fixture and a test.

## Why probes are disposable

A probe is tied to one code change or bug question. Once answered and
codified, it is irrelevant: the comment, fixture, and test reproduce the
result, so the snapshot trees can be re-derived or dropped anytime.
Losing an ignored probe directory (fresh clone, folder delete) loses
nothing except unfinished investigations — and an unfinished probe is the
only kind worth keeping.

## Standard shape

One directory is one question, and holds exactly one role script:

| Role | Script | Writes | Snapshot? |
|---|---|---|---|
| `probe` (throwaway measurement) | `probe.py` (or `probe.ps1`) | stdout only | never |
| `author` (builds a probe-local file) | `author.py` / `author.ps1` | `<name>.docx/.pptx/.xlsx` in-dir | optional `snapshot/` |
| `extend` (mutates `fixtures/`) | `extend.py` / `extend.ps1` | temp copy of `fixtures/*` | never (the diff lives in `fixtures/`) |

Every directory carries a `README.md`: the question, the exact Office
steps (or the script command), the finding (which XML changed), and where
it is codified (parse-site comment, fixture, test) — or "throwaway, no
fixture" when nothing landed.

New scripts import the shared helpers instead of copying them:

```python
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "_common"))
from office import word_app, fixture_copy, snapshot
from strip_label import main as strip_label
```

## `_common/`

`office.py` (Word/PowerPoint/Excel COM guards, fixture temp-copy,
unzip snapshot) and `strip_label.py` (org sensitivity-label removal,
run on a fixture after every Office save). The old
`tools/probe/pptx-lines/strip_label.py` path is a shim over the same
function.
