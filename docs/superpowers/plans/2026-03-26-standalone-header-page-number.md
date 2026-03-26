# Standalone Header And Page Number Modes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add two new standalone operations in the GUI, one that only updates page numbers and one that only updates headers, both using existing-sections-only behavior and no automatic section insertion.

**Architecture:** Extend the existing `only_insert` local-mode pattern used by TOC and cover so the formatter can route into two new minimal branches: strict page-number-only and strict header-only. Keep the GUI as a thin mode selector, put mode validation and routing inside `apply_format()`, and let `run_format()` choose a narrower Word post-process policy so standalone header/page-number runs do not accidentally refresh TOC formatting.

**Tech Stack:** Python 3.10, `python-docx`, Tkinter/ttkbootstrap GUI, `unittest`, Word COM post-processing via `pywin32`

---

## File Map

- Modify: `defaults/scau_2024.yaml`
  Responsibility: shipped YAML defaults and Chinese comments for the two new standalone flags.
- Modify: `thesis_config.py`
  Responsibility: in-memory default config used by GUI, CLI, and tests.
- Modify: `thesis_gui.py`
  Responsibility: standalone-mode state, Chinese remarks, mutual exclusion, config collect/load, and start-log messages.
- Modify: `thesis_formatter/formatter.py`
  Responsibility: local-mode validation, branch routing, and minimal save paths for strict standalone runs.
- Modify: `thesis_formatter/page.py`
  Responsibility: add a strict existing-sections-only page-number updater that never inserts a new section break.
- Modify: `thesis_runner.py`
  Responsibility: detect standalone modes for user logs and choose post-process policy (`full` / `fields_only` / `none`).
- Modify: `word_postprocess.py`
  Responsibility: support a lighter `fields_only` mode that updates document fields without running TOC restyling.
- Modify: `README.md`
  Responsibility: document the two new standalone operations and their strict-mode limitation in plain Chinese.
- Modify: `tests/test_thesis_gui_navigation.py`
  Responsibility: GUI harness regressions for standalone-mode mutual exclusion and start logging.
- Create: `tests/test_standalone_header_page_number.py`
  Responsibility: formatter-level regressions for strict page-number-only and header-only execution.
- Create: `tests/test_runner_local_modes.py`
  Responsibility: runner/postprocess regressions for standalone mode policy.

### Task 1: Add Config Flags And GUI Standalone Controls

**Files:**
- Modify: `defaults/scau_2024.yaml`
- Modify: `thesis_config.py`
- Modify: `thesis_gui.py`
- Modify: `tests/test_thesis_gui_navigation.py`

- [ ] **Step 1: Write the failing GUI regressions**

Add tests to `tests/test_thesis_gui_navigation.py` that lock the intended GUI behavior before any code changes.

```python
def test_standalone_modes_are_mutually_exclusive_and_clear_skip(self):
    self.gui._v_skip.set(True)
    self.gui._v_toc_only.set(True)
    self.gui._v_cover_only.set(True)
    self.gui._v_pn_only.set(True)

    self.gui._on_pn_only_toggle()

    self.assertFalse(self.gui._v_skip.get())
    self.assertFalse(self.gui._v_toc_only.get())
    self.assertFalse(self.gui._v_cover_only.get())
    self.assertTrue(self.gui._v_pn_only.get())
    self.assertFalse(self.gui._v_hf_only.get())


def test_start_logs_header_only_mode(self):
    # patch run_format + threading.Thread the same way as existing harness tests
    self.gui._v_hf_only.set(True)
    self.gui._start()
    self.assertIn("单独处理: 仅更新页眉", collected_logs)
```

- [ ] **Step 2: Run the GUI test file and verify it fails for the expected reasons**

Run: `py -3.10 -m unittest discover -s tests -p "test_thesis_gui_navigation.py" -v`
Expected: FAIL because `_v_pn_only` / `_v_hf_only` and the new toggle handlers do not exist yet, and the new log text is absent.

- [ ] **Step 3: Add the new config defaults in both config sources**

Update both config defaults so YAML files and in-memory defaults stay aligned.

```python
# thesis_config.py
"page_numbers": {
    "front_format": "upperRoman",
    "body_format": "decimal",
    "front_start": 1,
    "body_start": 1,
    "only_insert": False,
},
"header_footer": {
    "enabled": False,
    "scope": "body",
    "only_insert": False,
    ...
}
```

```yaml
# defaults/scau_2024.yaml
page_numbers:
  front_format: "upperRoman"
  body_format: "decimal"
  front_start: 1
  body_start: 1
  only_insert: false                 # 仅按现有分节调整页码，不自动补分节

header_footer:
  enabled: false
  scope: "body"
  only_insert: false                 # 仅按现有分节调整页眉，不自动补分节
```

- [ ] **Step 4: Add the standalone GUI state, pure-Chinese remarks, and mutual exclusion wiring**

Extend `thesis_gui.py` with two new BooleanVars and keep the naming consistent with the existing local-mode controls.

```python
self._v_pn_only = tk.BooleanVar(value=c["page_numbers"].get("only_insert", False))
self._v_hf_only = tk.BooleanVar(value=c["header_footer"].get("only_insert", False))
```

Add two new checkboxes to `_build_standalone()` with pure-Chinese notes:

```python
r = self._row_check(p, r, "单独改页码", self._v_pn_only, command=self._on_pn_only_toggle)
self._ttk.Label(p, text="只按现有分节调整页码，不自动补分节，尽量少动正文。", ...)

r = self._row_check(p, r, "单独改页眉", self._v_hf_only, command=self._on_hf_only_toggle)
self._ttk.Label(p, text="只按现有分节调整页眉，不自动补分节，尽量少动正文。", ...)
```

Make all four standalone modes mutually exclusive and clear `_v_skip` when any standalone mode is turned on.

```python
def _activate_standalone_mode(self, key):
    self._v_skip.set(False)
    self._v_toc_only.set(key == "toc")
    self._v_cover_only.set(key == "cover")
    self._v_pn_only.set(key == "pn")
    self._v_hf_only.set(key == "hf")
```

Update `_collect_config()` / `_load_vars_from_config()` and `_start()` logging so the new modes round-trip through config and show user-facing text without the term “严格模式”:

- `单独处理: 仅更新页码`
- `单独处理: 仅更新页眉`

- [ ] **Step 5: Run the GUI test file again and verify it passes**

Run: `py -3.10 -m unittest discover -s tests -p "test_thesis_gui_navigation.py" -v`
Expected: PASS for the existing navigation test and the new standalone-mode regressions.

- [ ] **Step 6: Commit the GUI/config plumbing**

```bash
git add defaults/scau_2024.yaml thesis_config.py thesis_gui.py tests/test_thesis_gui_navigation.py
git commit -m "feat: add standalone header and page number gui modes"
```

### Task 2: Implement Strict Formatter Branches For Existing Sections Only

**Files:**
- Modify: `thesis_formatter/formatter.py`
- Modify: `thesis_formatter/page.py`
- Create: `tests/test_standalone_header_page_number.py`

- [ ] **Step 1: Write the failing formatter regressions**

Create `tests/test_standalone_header_page_number.py` and cover both strict standalone branches.

```python
def test_page_number_only_mode_keeps_existing_section_count(self):
    cfg = _make_config()
    cfg["page_numbers"]["only_insert"] = True
    cfg["header_footer"]["enabled"] = False

    apply_format(str(input_path), str(output_path), config=cfg)

    out = Document(output_path)
    self.assertEqual(len(out.sections), 2)
    self.assertEqual(_section_pg_num_attrs(out.sections[0])["fmt"], "upperRoman")
    self.assertEqual(_section_pg_num_attrs(out.sections[1])["fmt"], "decimal")


def test_header_only_mode_uses_existing_sections_without_inserting_new_ones(self):
    cfg = _make_config()
    cfg["header_footer"]["enabled"] = True
    cfg["header_footer"]["only_insert"] = True

    apply_format(str(input_path), str(output_path), config=cfg)

    out = Document(output_path)
    self.assertEqual(len(out.sections), 2)
    self.assertEqual(out.sections[0].header.paragraphs[0].text, "")
    self.assertEqual(out.sections[1].header.paragraphs[0].text, "ODD_HEADER")
```

Also add a regression that multiple local modes are rejected explicitly.

```python
def test_multiple_local_modes_raise_runtime_error(self):
    cfg["toc"]["only_insert"] = True
    cfg["page_numbers"]["only_insert"] = True
    with self.assertRaises(RuntimeError):
        apply_format(..., config=cfg)
```

- [ ] **Step 2: Run the new formatter test file and verify it fails**

Run: `py -3.10 -m unittest discover -s tests -p "test_standalone_header_page_number.py" -v`
Expected: FAIL because the strict standalone branches and local-mode validation do not exist yet.

- [ ] **Step 3: Add local-mode validation and the strict branches in `apply_format()`**

Refactor the early branch section in `thesis_formatter/formatter.py` so all local modes are collected in one place.

```python
local_modes = {
    "toc": toc_cfg.get("only_insert", False),
    "cover": cover_cfg.get("only_insert", False),
    "page_numbers": cfg.get("page_numbers", {}).get("only_insert", False),
    "header_footer": cfg.get("header_footer", {}).get("only_insert", False),
}
active_local_modes = [name for name, enabled in local_modes.items() if enabled]
if len(active_local_modes) > 1:
    raise RuntimeError("单独处理模式不能同时启用多个。")
```

Add two new minimal branches after `doc = Document(input_path)`:

```python
if page_numbers_only:
    setup_page_numbers_strict(doc, cfg)
    doc.save(output_path)
    cfg.setdefault("_runtime", {})["local_mode"] = "page_numbers"
    return []

if header_only:
    setup_headers(doc, cfg)
    doc.save(output_path)
    cfg.setdefault("_runtime", {})["local_mode"] = "header_footer"
    return []
```

- [ ] **Step 4: Add a strict page-number helper in `thesis_formatter/page.py`**

Keep `setup_page_numbers()` for full-format mode, and add a new helper for standalone strict mode.

```python
def setup_page_numbers_strict(doc, cfg):
    first_body_h1 = find_first_body_heading(doc, cfg)
    body_section_index = get_body_start_section_index(doc, cfg, first_body_h1) if len(doc.sections) > 1 else 0
    # apply front/body page-number formats and footer PAGE fields only on existing sections
    # never call insert_section_break_before()
```

Reuse the existing footer rendering helpers (`add_page_number_field`, `set_section_page_number_format`, `_set_even_odd_on_doc`) so strict mode only changes page-number data and alignment.

- [ ] **Step 5: Run the new formatter test file again and verify it passes**

Run: `py -3.10 -m unittest discover -s tests -p "test_standalone_header_page_number.py" -v`
Expected: PASS for the strict page-number-only, header-only, and multi-mode rejection regressions.

- [ ] **Step 6: Commit the formatter branch work**

```bash
git add thesis_formatter/formatter.py thesis_formatter/page.py tests/test_standalone_header_page_number.py
git commit -m "feat: add strict standalone header and page number formatter modes"
```

### Task 3: Add Runner Post-Process Policy For Local Modes

**Files:**
- Modify: `thesis_runner.py`
- Modify: `word_postprocess.py`
- Create: `tests/test_runner_local_modes.py`

- [ ] **Step 1: Write the failing runner/postprocess regressions**

Create `tests/test_runner_local_modes.py` and patch the runner dependencies so the tests only exercise policy decisions.

```python
@mock.patch("thesis_runner.apply_format", return_value=[])
@mock.patch("thesis_runner.postprocess")
def test_page_number_only_mode_skips_postprocess(mock_postprocess, _mock_apply):
    cfg = _make_config()
    cfg["page_numbers"]["only_insert"] = True
    ok = run_format(str(input_path), str(output_path), logs.append, config=cfg)
    self.assertTrue(ok)
    mock_postprocess.assert_not_called()


@mock.patch("thesis_runner.apply_format", return_value=[])
@mock.patch("thesis_runner.postprocess")
def test_header_only_mode_with_chapter_title_runs_fields_only(mock_postprocess, _mock_apply):
    cfg = _make_config()
    cfg["header_footer"]["only_insert"] = True
    cfg["header_footer"]["enabled"] = True
    cfg["header_footer"]["odd_page_text"] = "{chapter_title}"
    ok = run_format(str(input_path), str(output_path), logs.append, config=cfg)
    self.assertTrue(ok)
    mock_postprocess.assert_called_once_with(str(output_path), config=cfg, mode="fields_only")
```

- [ ] **Step 2: Run the runner policy test file and verify it fails**

Run: `py -3.10 -m unittest discover -s tests -p "test_runner_local_modes.py" -v`
Expected: FAIL because `run_format()` has no local-mode postprocess policy and `postprocess()` has no `mode` parameter.

- [ ] **Step 3: Add a local-mode postprocess resolver in `thesis_runner.py`**

Compute one postprocess mode from config/runtime instead of using the current binary behavior.

```python
def _resolve_postprocess_mode(config):
    runtime = config.get("_runtime", {})
    if runtime.get("cover_only") or config.get("cover", {}).get("only_insert", False):
        return "none"
    if config.get("toc", {}).get("only_insert", False):
        return "full"
    if config.get("page_numbers", {}).get("only_insert", False):
        return "none"
    if config.get("header_footer", {}).get("only_insert", False):
        texts = [config.get("header_footer", {}).get("odd_page_text", ""), config.get("header_footer", {}).get("even_page_text", "")]
        return "fields_only" if any("{chapter_title}" in text for text in texts) else "none"
    return "full"
```

Update the user log text so standalone page-number mode says no postprocess is run, and standalone header mode only announces field refresh when chapter-title fields are present.

- [ ] **Step 4: Extend `word_postprocess.py` to support `fields_only`**

Add a `mode` parameter and gate the heavy TOC work behind it.

```python
def postprocess(docx_path, timeout=90, config=None, mode="full"):
    ...
    if mode == "full":
        for toc in doc.TablesOfContents:
            toc.Update()
    if mode in {"full", "fields_only"}:
        doc.Fields.Update()
    if mode == "full":
        # keep the current TOC font/style fix loop
```

Leave `mode="full"` as the default so existing callers keep the old behavior.

- [ ] **Step 5: Run the runner policy test file again and verify it passes**

Run: `py -3.10 -m unittest discover -s tests -p "test_runner_local_modes.py" -v`
Expected: PASS for the no-postprocess and fields-only policy regressions.

- [ ] **Step 6: Commit the runner/postprocess changes**

```bash
git add thesis_runner.py word_postprocess.py tests/test_runner_local_modes.py
git commit -m "feat: add standalone mode postprocess policy"
```

### Task 4: Update User Docs And Verify The Whole Feature Set

**Files:**
- Modify: `README.md`
- Test: `tests/test_thesis_gui_navigation.py`
- Test: `tests/test_standalone_header_page_number.py`
- Test: `tests/test_runner_local_modes.py`
- Test: `tests/test_front_matter_recognition.py`
- Test: `tests/test_multi_section_body_headers.py`

- [ ] **Step 1: Update the README for the new standalone actions**

Add a short section that describes the two new standalone options and their existing-sections-only limitation in plain Chinese.

```markdown
- 单独改页码：只按现有分节调整页码，不自动补分节，尽量少动正文。
- 单独改页眉：只按现有分节调整页眉，不自动补分节，尽量少动正文。
```

- [ ] **Step 2: Run the full unittest suite**

Run: `py -3.10 -m unittest discover -s tests -v`
Expected: all test files pass, including the new standalone-mode regressions and the existing front-matter/body-section regressions.

- [ ] **Step 3: Re-run only the high-risk entrance-layer tests if the full suite fails anywhere**

Run, in order, until the failing area is isolated:

```bash
py -3.10 -m unittest discover -s tests -p "test_thesis_gui_navigation.py" -v
py -3.10 -m unittest discover -s tests -p "test_standalone_header_page_number.py" -v
py -3.10 -m unittest discover -s tests -p "test_runner_local_modes.py" -v
```

Expected: one focused failure area that can be fixed without guessing. Follow @superpowers:systematic-debugging before changing code again.

- [ ] **Step 4: Final verification before completion**

Run: `py -3.10 -m unittest discover -s tests -v`
Expected: clean exit code 0 and no failing tests. Follow @superpowers:verification-before-completion before claiming the feature is done.

- [ ] **Step 5: Commit the docs and final verified state**

```bash
git add README.md
# plus any last-minute test or implementation fixes from Task 4
git commit -m "docs: document standalone header and page number modes"
```

