# Phase01 MySQL Study Presentation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `phase01`의 1~5주차 핵심 개념과 EXPLAIN 판독법을 컴퓨터공학 4학년 입문자에게 설명하는 15장 PowerPoint를 Simple Light Mode 템플릿으로 제작한다.

**Architecture:** retained Simple Light Mode `reference.pptx`를 전수 검사하고 15개 출력 슬라이드를 원본 슬라이드에 매핑한 `template-starter.pptx`를 만든다. `@oai/artifact-tool`로 starter를 불러와 상속된 요소만 편집하고, 원문 이미지와 텍스트를 삽입한 뒤 렌더·레이아웃·템플릿 충실도 검사를 반복한다.

**Tech Stack:** Node.js ES modules, `@oai/artifact-tool`, Presentations template-following scripts, PowerPoint `.pptx`, Poppler/LibreOffice 기반 렌더 도구, PowerShell

## Global Constraints

- 출력은 15장, 16:9 Simple Light Mode 원본 화면 비율을 유지한다.
- 시각 기준은 `C:\Users\dbsgh\.codex\plugins\cache\openai-curated-remote\openai-templates\0.1.0\skills\artifact-template-simple-light-mode\assets\reference.pptx` 하나만 사용한다.
- 원본 템플릿은 수정하지 않고 별도 PPTX를 생성한다.
- 외부 사실이나 성능 수치를 만들지 않고 `phase01`의 Markdown 7개와 포함 이미지에서만 내용을 가져온다.
- `python-pptx`, 직접 OOXML 수정, Python 기반 이미지 생성은 사용하지 않는다.
- 모든 슬라이드는 한 장당 핵심 주장 하나를 갖고, 화면 문구는 청중용 문장으로 작성한다.
- `type = ALL`, `Using filesort`, `Using temporary`를 무조건 오류로 표현하지 않는다.
- `Using index`, `Using index condition`, `Using index for group-by`를 구분한다.
- 코드와 EXPLAIN 캡처는 발표 화면에서 읽을 수 있는 크기를 유지한다.
- 중간 파일은 외부 임시 작업공간에 두고, 저장소에는 최종 PPTX와 승인된 문서만 둔다.

---

## File Structure

- Read: `phase01/01-mysql-architecture-and-query-execution.md`
- Read: `phase01/01-mysql-architecture-and-query-execution-practice.md`
- Read: `phase01/02-mysql_index.md`
- Read: `phase01/02-mysql_index_practice.md`
- Read: `phase01/03_mysql_optimizer_query_and_order.md`
- Read: `phase01/04-mysql_index.md`
- Read: `phase01/05-mysql_covering_index.md`
- Read: `phase01/images/03_mysql_optimizer_query_and_order_01_full_table_scan_explain.png`
- Read: `phase01/images/03_mysql_optimizer_query_and_order_07_index_sort_explain.png`
- Read: `phase01/images/03_mysql_optimizer_query_and_order_13_loose_index_scan_min_explain.png`
- Read: `phase01/images/03_mysql_optimizer_query_and_order_14_temporary_table_group_by_explain.png`
- Read: `phase01/images/05_mysql_covering_index_index_range_scan_02.png`
- Read: `phase01/images/05_innodb_vs_myisam_secondary_index.png`
- Read: `phase01/images/05_covering_index_not_applied.png`
- Read: `phase01/images/05_covering_index_applied.png`
- Create in external scratch: `tmp/source-notes.txt`
- Create in external scratch: `tmp/template-audit.txt`
- Create in external scratch: `tmp/template-frame-map.json`
- Create in external scratch: `tmp/deviation-log.txt`
- Create in external scratch: `tmp/template-starter.pptx`
- Create in external scratch: `tmp/slide-copy.txt`
- Create in external scratch: `tmp/phase01-mysql-study.mjs`
- Create: `outputs/phase01-mysql-study-summary.pptx`

---

### Task 1: Initialize and Audit the Retained Template

**Files:**
- Read: `C:\Users\dbsgh\.codex\plugins\cache\openai-curated-remote\openai-templates\0.1.0\skills\artifact-template-simple-light-mode\assets\reference.pptx`
- Create in external scratch: `tmp/template-audit.txt`
- Create in external scratch: `tmp/template-frame-map.json`
- Create in external scratch: `tmp/deviation-log.txt`

**Interfaces:**
- Consumes: retained Simple Light Mode reference PPTX and Presentations template-following scripts
- Produces: validated source inventory, source-slide mapping for output slides 1–15, and starter-deck edit target IDs

- [ ] **Step 1: Set deterministic workspace paths**

Run in PowerShell from the repository root:

```powershell
$node = 'C:\Users\dbsgh\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
$skillDir = 'C:\Users\dbsgh\.codex\plugins\cache\openai-primary-runtime\presentations\26.715.12143\skills\presentations'
$template = 'C:\Users\dbsgh\.codex\plugins\cache\openai-curated-remote\openai-templates\0.1.0\skills\artifact-template-simple-light-mode\assets\reference.pptx'
$scratchRoot = & $node -p "require('node:os').tmpdir()"
$workspace = Join-Path $scratchRoot 'codex-presentations\manual-20260721\phase01-mysql-study'
$tmpDir = Join-Path $workspace 'tmp'
New-Item -ItemType Directory -Force -Path $tmpDir, (Join-Path $tmpDir 'preview'), (Join-Path $tmpDir 'layout'), (Join-Path $tmpDir 'assets'), (Join-Path $tmpDir 'qa') | Out-Null
```

Expected: all scratch paths exist outside the repository.

- [ ] **Step 2: Initialize artifact-tool resolution**

Run:

```powershell
& $node "$skillDir\container_tools\setup_artifact_tool_workspace.mjs" --workspace $tmpDir
```

Expected: the scratch workspace can resolve `@oai/artifact-tool`.

- [ ] **Step 3: Inspect every source slide**

Run:

```powershell
& $node "$skillDir\template_following_scripts\inspect_template_deck.mjs" --workspace $tmpDir --pptx $template
```

Expected: `template-inspect.ndjson`, per-slide renders, layout JSON, extracted media, font evidence, and `template-manifest.json` are created.

- [ ] **Step 4: Review the complete inventory**

Inspect every rendered slide at full size, the montage, all layout JSON files, and `template-inspect.ndjson`. Record in `template-audit.txt`:

```text
Source slide inventory
- source slide number and narrative/layout role
- inherited title/body/image/code slots
- typography family, size, weight, paragraph spacing and insets
- recurring footer, date, page marker and brand chrome
- slides that can support title, section, process, comparison, screenshot and closing roles

Insertion contract
- preserve all inherited chrome
- rewrite only mapped placeholder/text IDs
- replace only mapped image IDs
- delete structural placeholders that cannot be intentionally filled
- do not add parallel layouts over inherited elements
```

- [ ] **Step 5: Build the 15-slide frame map**

Create `template-frame-map.json` with one mapping for each narrative role:

```json
{
  "outputSlides": [
    { "outputSlide": 1, "narrativeRole": "opening question", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 2, "narrativeRole": "five-week learning flow", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 3, "narrativeRole": "week 1 execution pipeline", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 4, "narrativeRole": "EXPLAIN tabular reading order", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 5, "narrativeRole": "EXPLAIN ANALYZE tree reading", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 6, "narrativeRole": "week 2 B+Tree mechanism", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 7, "narrativeRole": "composite index rule", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 8, "narrativeRole": "week 3 cost model", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 9, "narrativeRole": "ORDER BY and join sorting", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 10, "narrativeRole": "week 4 WHERE GROUP BY HAVING", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 11, "narrativeRole": "GROUP BY DISTINCT plan signals", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 12, "narrativeRole": "week 5 InnoDB index structure", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 13, "narrativeRole": "covering index comparison", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 14, "narrativeRole": "integrated query walkthrough", "reuseMode": "duplicate-slide", "editTargets": [] },
    { "outputSlide": 15, "narrativeRole": "verification checklist", "reuseMode": "duplicate-slide", "editTargets": [] }
  ],
  "omittedSourceSlides": []
}
```

For each entry, fill `sourceSlide` from the inspected inventory and populate `editTargets` with exact `shapeId`, `sourceElementId`, action (`rewrite`, `replace`, or `delete`), and bounded insertion zones when a source layout explicitly supports them. List every omitted source slide and reason.

- [ ] **Step 6: Validate the template plan**

Run:

```powershell
& $node "$skillDir\template_following_scripts\validate_template_plan.mjs" --workspace $tmpDir --map "$tmpDir\template-frame-map.json"
```

Expected: PASS with all 15 output slides mapped and every edit target resolved.

- [ ] **Step 7: Create and visually inspect the starter deck**

Run:

```powershell
& $node "$skillDir\template_following_scripts\prepare_template_starter_deck.mjs" --workspace $tmpDir --pptx $template --map "$tmpDir\template-frame-map.json" --out "$tmpDir\template-starter.pptx" --preview-dir "$tmpDir\template-starter-preview" --layout-dir "$tmpDir\template-starter-layout" --contact-sheet "$tmpDir\template-starter-contact-sheet.png"
```

Expected: 15 template-derived slides with no invented visual system.

---

### Task 2: Freeze Audience-Facing Copy and Asset Provenance

**Files:**
- Read: `docs/superpowers/specs/2026-07-20-phase01-presentation-design.md`
- Create in external scratch: `tmp/source-notes.txt`
- Create in external scratch: `tmp/slide-copy.txt`

**Interfaces:**
- Consumes: approved 15-slide design and the seven phase01 Markdown sources
- Produces: final slide titles, visible copy, SQL snippets, EXPLAIN values, image provenance, and alt text consumed by the builder

- [ ] **Step 1: Create the source ledger**

Write `source-notes.txt` with one block per slide:

```text
Slide 4 — EXPLAIN is read in four passes
Sources:
- phase01/01-mysql-architecture-and-query-execution.md, EXPLAIN section
- phase01/01-mysql-architecture-and-query-execution-practice.md, EXPLAIN summary
Asset:
- phase01/images/03_mysql_optimizer_query_and_order_01_full_table_scan_explain.png
Claim limits:
- rows is an estimate, not returned row count
- ALL is an access path, not an automatic defect
```

Repeat for slides 1–15. Every image entry must record the original file path and its intended crop.

- [ ] **Step 2: Write final slide copy**

Write `slide-copy.txt` with exactly these titles and content budgets:

```text
1. MySQL은 어떻게 덜 읽는가? — title + one subtitle
2. 다섯 주차는 하나의 최적화 흐름이다 — five short stages
3. SQL은 선언이고, 실행 방법은 옵티마이저가 정한다 — five-step pipeline + two role labels
4. EXPLAIN은 네 번에 나눠 읽는다 — four numbered inspection steps + one caution
5. EXPLAIN ANALYZE는 추정과 실제의 차이를 보여준다 — one tree example + four labels
6. B+Tree는 탐색한 뒤 필요한 구간만 읽는다 — seek/scan/row lookup + trade-off
7. 복합 인덱스는 왼쪽부터 검색 범위를 만든다 — lexicographic example + usable/unusable predicates
8. 옵티마이저는 인덱스 유무가 아니라 총비용을 비교한다 — cardinality/selectivity + scan decision
9. 인덱스 순서가 맞으면 정렬 결과를 바로 흘려보낸다 — index order/filesort + LIMIT
10. 그룹핑 전에 행 수를 먼저 줄인다 — WHERE/GROUP BY/HAVING sequence
11. GROUP BY와 DISTINCT도 같은 값을 모으는 비용을 가진다 — tight/loose/temporary + three Extra messages
12. 실제 레코드는 PK 순서의 리프에 있다 — clustered/secondary/double lookup
13. 커버링 인덱스는 두 번째 B+Tree 탐색을 없앤다 — Using index versus Using index condition
14. 한 쿼리에 다섯 주차가 모두 들어 있다 — integrated SQL + five checks
15. 좋은 인덱스는 실행 계획으로 검증한다 — five-question checklist
```

Limit each slide to one claim, 3–5 short text items, and at most one readable code or EXPLAIN example.

- [ ] **Step 3: Copy and normalize selected raster assets**

Copy the eight selected source images into `tmp/assets` without editing their meaning. Use `ensure_raster_image.py` only if a source is not already a supported PNG. Keep the original aspect ratio and record crop bounds in `source-notes.txt`.

- [ ] **Step 4: Review content accuracy**

Check that:

```text
- EXPLAIN type is described as access method, not a score.
- possible_keys and key are not conflated.
- rows and filtered are estimates.
- actual time, rows, and loops are identified as EXPLAIN ANALYZE measurements.
- Using filesort does not claim mandatory disk sorting.
- Using index, Using index condition, and Using index for group-by are distinct.
- SELECT * is described as reducing covering-index opportunities, not prohibited syntax.
```

Expected: all checks pass and no unsupported performance multiplier appears.

---

### Task 3: Author the 15 Slides with Artifact Tool

**Files:**
- Create in external scratch: `tmp/phase01-mysql-study.mjs`
- Read: `tmp/template-starter.pptx`
- Read: `tmp/template-frame-map.json`
- Read: `tmp/slide-copy.txt`
- Create: `outputs/phase01-mysql-study-summary.pptx`

**Interfaces:**
- Consumes: inherited starter deck, exact element IDs, approved copy, and byte-backed PNG assets
- Produces: editable 15-slide PPTX and per-slide PNG/layout exports

- [ ] **Step 1: Create the builder module skeleton**

Implement these exact interfaces in `phase01-mysql-study.mjs`:

```js
import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, PresentationFile } from '@oai/artifact-tool';

async function readImageBlob(imagePath) {
  const bytes = await fs.readFile(imagePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function resolveRequired(presentation, id, label) {
  const value = presentation.resolve(id);
  if (!value) throw new Error(`Missing inherited element: ${label} (${id})`);
  return value;
}

function mappedTarget(frameMap, slideNumber, role) {
  const slideMap = frameMap.outputSlides.find(
    (entry) => entry.outputSlide === slideNumber
  );
  const target = slideMap?.editTargets.find((entry) => entry.role === role);
  const id = target?.sourceElementId ?? target?.shapeId;
  if (!id) throw new Error(`Missing mapped target: slide ${slideNumber}, role ${role}`);
  return { ...target, id };
}

function replaceText(presentation, target) {
  const shape = resolveRequired(presentation, target.id, target.name);
  shape.text = target.text;
}

async function replaceImage(slide, presentation, target) {
  const inherited = resolveRequired(presentation, target.id, target.name);
  inherited.delete();
  slide.images.add({
    blob: await readImageBlob(target.path),
    contentType: 'image/png',
    alt: target.alt,
    fit: target.fit,
    position: target.position,
    geometry: target.geometry ?? 'rect'
  });
}
```

Do not use `presentation.slides.add()`; import and edit only the 15 inherited starter slides.

- [ ] **Step 2: Import and verify the starter deck**

Add:

```js
const presentation = await PresentationFile.importPptx(
  await FileBlob.load(process.env.STARTER_PPTX)
);

if (presentation.slides.items.length !== 15) {
  throw new Error(`Expected 15 slides, got ${presentation.slides.items.length}`);
}
```

Expected: the builder stops immediately if the frame map produced the wrong slide count.

- [ ] **Step 3: Encode inherited-element edits slide by slide**

Create an `edits` array with 15 objects. Each object must contain:

```js
const frameMap = JSON.parse(
  await fs.readFile(process.env.FRAME_MAP, 'utf8')
);
const slide4Title = mappedTarget(frameMap, 4, 'title');
const slide4Image = mappedTarget(frameMap, 4, 'primary-image');

{
  slideNumber: 4,
  textTargets: [
    { id: slide4Title.id, name: 'title', text: 'EXPLAIN은 네 번에 나눠 읽는다' }
  ],
  imageTargets: [
    {
      id: slide4Image.id,
      name: 'explain-grid',
      path: path.resolve('phase01/images/03_mysql_optimizer_query_and_order_01_full_table_scan_explain.png'),
      alt: 'EXPLAIN result showing type ALL, key NULL, rows 954, and Using where',
      fit: 'contain',
      position: slide4Image.position
    }
  ]
}
```

Every mapped edit target must include a stable `role` value such as `title`, `body`, `primary-image`, `code`, or `footer`; the builder resolves IDs exclusively from `template-frame-map.json`. Slides 4, 11 and 13 must use the real phase01 EXPLAIN images; slides 5 and 14 use editable text for tree output and SQL so they remain readable.

- [ ] **Step 4: Apply all edits without changing inherited typography**

For each edit:

```js
for (const edit of edits) {
  const slide = presentation.slides.items[edit.slideNumber - 1];
  for (const target of edit.textTargets ?? []) replaceText(presentation, target);
  for (const target of edit.imageTargets ?? []) {
    await replaceImage(slide, presentation, target);
  }
}
```

If text does not fit, shorten `slide-copy.txt` or remap to a more suitable source layout. Do not shrink inherited font sizes or overlay a second layout.

- [ ] **Step 5: Export previews, layouts, montage, and final PPTX**

Add:

```js
async function writeBlob(filePath, blob) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

for (const [index, slide] of presentation.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, '0')}`;
  await writeBlob(
    path.join(process.env.PREVIEW_DIR, `${stem}.png`),
    await presentation.export({ slide, format: 'png', scale: 2 })
  );
  const layout = await slide.export({ format: 'layout' });
  await fs.writeFile(
    path.join(process.env.LAYOUT_DIR, `${stem}.layout.json`),
    await layout.text()
  );
}

await writeBlob(
  path.join(process.env.QA_DIR, 'deck-montage.webp'),
  await presentation.export({ format: 'webp', montage: true, scale: 1 })
);

const pptx = await PresentationFile.exportPptx(presentation);
await fs.mkdir(path.dirname(process.env.FINAL_PPTX), { recursive: true });
await pptx.save(process.env.FINAL_PPTX);
```

- [ ] **Step 6: Run the builder**

Run:

```powershell
$env:STARTER_PPTX = "$tmpDir\template-starter.pptx"
$env:FRAME_MAP = "$tmpDir\template-frame-map.json"
$env:PREVIEW_DIR = "$tmpDir\preview\final"
$env:LAYOUT_DIR = "$tmpDir\layout\final"
$env:QA_DIR = "$tmpDir\qa"
$env:FINAL_PPTX = (Resolve-Path '.').Path + '\outputs\phase01-mysql-study-summary.pptx'
& $node "$tmpDir\phase01-mysql-study.mjs"
```

Expected: exactly 15 PNGs, 15 layout JSON files, one montage, and one PPTX.

---

### Task 4: Perform Slide-by-Slide Visual QA and Repair

**Files:**
- Inspect: `tmp/preview/final/slide-01.png` through `slide-15.png`
- Inspect: `tmp/layout/final/*.layout.json`
- Modify: `tmp/phase01-mysql-study.mjs`
- Modify: `tmp/slide-copy.txt`
- Regenerate: `outputs/phase01-mysql-study-summary.pptx`

**Interfaces:**
- Consumes: first-pass rendered deck
- Produces: visually legible, overlap-free and template-faithful final deck

- [ ] **Step 1: Review the montage for narrative flow**

Confirm the deck progresses through:

```text
question → execution → EXPLAIN → index → optimizer/sort → grouping → covering → integrated check
```

Expected: no abrupt section jump and no consecutive slides with indistinguishable composition.

- [ ] **Step 2: Review every slide at full size**

For each slide, record PASS or a concrete repair in `tmp/qa/slide-review.txt` for:

```text
primary claim
title wrapping
body clipping
code readability
EXPLAIN column readability
image crop fidelity
spacing rhythm
unintended overlap
inherited footer/page marker
source accuracy
```

- [ ] **Step 3: Fix copy before changing layout**

When a text box overflows, shorten the copy in `slide-copy.txt`. When content still does not fit, remap that output slide to a more suitable source layout and regenerate the starter. Never reduce inherited template font sizes just to fit copy.

- [ ] **Step 4: Fix EXPLAIN-specific legibility**

Verify:

```text
Slide 4: type, key, rows and Extra are readable without zoom.
Slide 5: actual time, rows and loops are visually distinct.
Slide 11: Using index for group-by and Using temporary are not conflated.
Slide 13: Using index condition and Using index are compared at equal scale.
```

- [ ] **Step 5: Regenerate after repairs**

Run the builder again and replace all final previews and the PPTX. Repeat full-size inspection until all 15 slide-review entries are PASS.

---

### Task 5: Run Structural and Template Fidelity Gates

**Files:**
- Test: `outputs/phase01-mysql-study-summary.pptx`
- Test: `tmp/template-frame-map.json`
- Test: `tmp/layout/final`
- Create in external scratch: `tmp/qa/final-verification.txt`

**Interfaces:**
- Consumes: visually approved PPTX
- Produces: evidence that the final deck has no overflow, empty placeholders, or template drift

- [ ] **Step 1: Run slide canvas overflow detection**

Run:

```powershell
& 'C:\Users\dbsgh\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' "$skillDir\container_tools\slides_test.py" $env:FINAL_PPTX
```

Expected: no content outside the original slide canvas.

- [ ] **Step 2: Run template fidelity verification**

Run:

```powershell
& $node "$skillDir\template_following_scripts\check_template_fidelity.mjs" --workspace $tmpDir --starter-pptx "$tmpDir\template-starter.pptx" --final-pptx $env:FINAL_PPTX --map "$tmpDir\template-frame-map.json" --starter-layout-dir "$tmpDir\template-starter-layout" --final-layout-dir "$tmpDir\layout\final" --edit-dir $tmpDir
```

Expected: PASS with only the mapped edits recorded in `deviation-log.txt`.

- [ ] **Step 3: Inspect final PPTX placeholders**

Use the fidelity output to confirm no `<p:sp>` containing `<p:ph>` has an empty or whitespace-only text body unless that exact placeholder is classified for deletion. Also verify no visible `Slide Number`, `Date`, `Footer`, `Click to add title`, or template sample text remains.

- [ ] **Step 4: Re-render the exported PPTX independently**

Run:

```powershell
& 'C:\Users\dbsgh\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' "$skillDir\container_tools\render_slides.py" $env:FINAL_PPTX
```

Expected: 15 rendered PNGs from the final exported file, matching the artifact-tool previews.

- [ ] **Step 5: Record final verification evidence**

Write `tmp/qa/final-verification.txt`:

```text
Slide count: 15 — PASS
Template fidelity: PASS
Canvas overflow: PASS
Empty placeholders: PASS
Full-size visual review: 15/15 PASS
EXPLAIN readability: slides 4, 5, 11, 13 PASS
Output: outputs/phase01-mysql-study-summary.pptx
```

- [ ] **Step 6: Confirm repository scope**

Run:

```powershell
git status --short
```

Expected: existing user-owned `.obsidian/` remains untouched; only the plan document and final PPTX are new or modified by this task.
