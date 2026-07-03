---
name: ppt-skill
description: |
  Generate general-purpose PowerPoint presentations (.pptx) from structured slide data.
  This skill is not limited to product managers or product proposals. It can be used
  for product proposals, project reports, training decks, roadshows, business reviews,
  event plans, company introductions, competitive analyses, and other presentation
  scenarios.

  Two rendering modes are supported:
  (A) Template mode: if the user provides a .pptx template, fill content into the
      template while preserving master slides, backgrounds, themes, and layouts.
  (B) Theme mode: if no template is provided, build slides from scratch with a built-in
      theme: default, neutral, neutral_dark, blue, green, orange, purple, red,
      or teal.

  The calling Agent is responsible for understanding the user's intent, deciding or
  confirming the outline, expanding slide content, optionally collecting image files,
  and formatting the result as a slides JSON array. The CLI is responsible for rendering
  slides JSON into a .pptx file.

  Trigger when the user asks to create or generate a PPT, slide deck, presentation,
  proposal deck, report deck, training deck, or similar. Do not use this skill for
  reading or editing an existing PPTX in place.
---
# PPT-Skill

## Positioning

`ppt-skill` is a general-purpose PPT generation skill. It turns structured slide content
into a `.pptx` file.

It is designed to work in Codex and in other Agent environments. Do not rely on Codex-only
features in the core workflow. Treat local files, CLI commands, JSON inputs, and ordinary
conversation as the portable interface.

This skill can help the Agent create an outline when the user only gives a topic, but
outline creation is an Agent workflow responsibility. The CLI remains a renderer:

```text
User intent or topic -> Agent outline -> Agent slides JSON -> CLI -> .pptx
```

## Responsibility Boundary

The Agent is responsible for:

- Understand the user's purpose, audience, expected depth, language, and style.
- Detect whether a PPT template was provided.
- Use the user's outline when provided.
- Generate a dynamic outline when the user only provides a topic.
- Expand the outline into concise slide content.
- Convert slide content into the supported slides JSON format.
- Collect optional image files from the user or other tools when images are needed.
- Call the CLI with `--template` or `--theme`.

The CLI is responsible for:

- Render slides JSON into `.pptx`.
- Fill content into a user-provided template when `--template` is used.
- Build slides from built-in themes when `--theme` is used.
- Embed provided image files into known slots.
- Skip missing optional images without blocking deck generation where possible.

The skill does not require or assume:

- A product-manager-only use case.
- Automatic business research.
- Automatic image generation.
- A fixed deck length.
- Codex-specific runtime behavior.

## Input Priority

Follow this priority order.

1. User-provided PPT template

   If the user provides a `.pptx` template, use template mode. The template controls
   visual style, master slides, backgrounds, theme, and available layouts.

2. User-provided outline

   If the user provides a PPT outline, follow that outline. The outline controls the
   section structure, page count, page order, and content intent. Only fill gaps when
   needed.

3. User-provided topic only

   If the user only provides a topic, create a suggested outline first. The outline must
   be dynamic: page count and page structure depend on purpose, audience, complexity,
   and user constraints.

When both a template and an outline are provided:

```text
PPT template = how it looks
PPT outline = what it says, how many pages, and in what order
slides JSON = executable content structure
CLI = renderer
```

## Rendering Modes

### Mode A: Template Fill

Use this mode when the user provides a `.pptx` template.

Rules:

- Do not ask the user to choose a built-in theme.
- Preserve template master slides, backgrounds, theme, and layouts.
- Convert the user outline or Agent-generated outline into slides JSON.
- Fill title, subtitle, body, tables, and lists into placeholders where available.
- Use fallback text boxes when a matching placeholder is not available.

Example:

```bash
ppt-skill create \
  --template brand.pptx \
  --slides /path/to/slides.json \
  --images '{}' \
  --output out.pptx
```

### Mode B: Built-In Theme

Use this mode when the user does not provide a template.

Rules:

- Ask or infer presentation style from the user request.
- If the user explicitly chooses a theme, use it.
- If the user says "you decide", "business style", or gives a clear use case, the Agent
  may choose a suitable theme and briefly state the choice.
- If the Agent cannot infer style safely, ask the user to choose.

Available themes:

- `default`: default business blue
- `neutral`: neutral minimal light theme
- `neutral_dark`: neutral minimal dark theme
- `blue`: technology blue
- `green`: forest green
- `orange`: warm business orange
- `purple`: elegant purple
- `red`: China red
- `teal`: teal

### Optional Visual Style: Neutral Minimal

Use `neutral` or `neutral_dark` only when the user explicitly or implicitly asks for
one of these styles:

- non-blue-purple, no blue/purple, avoid blue/purple, or similar constraints
- 中性极简风格, 非蓝紫风格, 克制商务风, 去科技蓝风格
- low-saturation, grayscale, graphite, black-white-gray, premium minimal, or whitespace-heavy style
- light/dark PPT theme variants

Rules for Neutral Minimal:

- Do not use blue, purple, or adjacent hues as primary colors or large background areas.
- Avoid large gradient backgrounds.
- Avoid left vertical accent bars as the default emphasis pattern. Use small markers,
  short rules, restrained labels, or typography instead.
- Use neutral and low-saturation colors as the base palette.
- Use a small amount of restrained accent color for key numbers, icons, labels, and
  conclusions.
- Emphasize whitespace, clear hierarchy, and one main idea per slide.
- Use visually sharp small-radius cards, chart containers, and image masks; target an
  equivalent 2px radius when the renderer supports radius control.
- Treat light and dark as separate PPT theme variants, not runtime theme switching.

Theme selection rule:

- If the user asks for Neutral Minimal / 非蓝紫 / 中性极简 style, choose `neutral` by
  default, or `neutral_dark` when the user asks for a dark style.
- In this case, do not choose `default`, `blue`, `purple`, or `teal` unless the user
  explicitly overrides the style constraint.

Example:

```bash
ppt-skill create \
  --theme blue \
  --slides /path/to/slides.json \
  --images '{}' \
  --output out.pptx
```

## Outline Workflow

The Agent should support three content-entry levels.

### Level 1: User Provides slides JSON

Use it directly after checking that it is a JSON array and each slide has a supported
`type`.

### Level 2: User Provides an Outline

Follow the user's outline. Preserve page order and page count unless the user asks for
optimization or the outline is impossible to render cleanly.

If the outline is sparse, enrich each page with concise slide-ready points. Do not invent
facts, data, customer names, pricing, legal claims, medical claims, or financial claims.
Use explicit placeholders such as `待补充数据` when the user has not provided evidence.

### Level 3: User Provides Only a Topic

Generate a suggested outline before generating the PPT.

The outline must be dynamic, not fixed to 8 pages. Use the user's purpose, audience,
complexity, and constraints to decide depth.

Suggested page-count bands:

- 5-7 pages: lightweight briefing, quick intro, meeting sync.
- 8-12 pages: standard proposal, project report, solution deck, customer discussion.
- 13-20 pages: detailed plan, roadshow, training deck, consulting-style report.
- More than 20 pages: only when the user explicitly asks for a course, whitepaper-style
  deck, complete business plan, or deep report.

Default behavior when the user gives only a topic and no constraints:

- Create a 10-12 page standard outline.
- Tell the user it can be shortened or expanded.
- If the user has already said "you decide" or "直接生成", proceed without another
  confirmation.

Recommended outline output format:

```text
建议做成 10-12 页标准版：

1. 页面标题
   目的：这一页解决什么表达任务
   重点：要讲的 2-4 个核心点
   类型：cover/problem/solution/capabilities/how_it_works/roadmap/budget/value/next_steps/...
```

## Image Strategy

This skill has no built-in image generation capability.

Images are optional inputs. The deck must still generate when no images are provided.

Use one of these modes:

1. No-image mode

   Default and most stable mode. Pass `--images '{}'`. The CLI uses text, shapes, cards,
   tables, timelines, colors, and layout to create visual structure.

2. User-provided image mode

   If the user provides logos, screenshots, product images, diagrams, team photos, or
   other assets, map them to image slots and pass them through `--images`.

3. External image provider mode

   If the current Agent environment has a separate image-generation, icon, screenshot,
   or asset-search capability, that external tool may create or collect images first.
   `ppt-skill` still only embeds the resulting local image files.

Do not require image generation to complete a PPT.

Available image slots:

| Page type | Common use | Slot name |
| --- | --- | --- |
| cover | background image, logo | `cover_bg`, `cover_logo` |
| solution | architecture or process diagram | `solution_diagram` |
| capabilities | capability card icons | `capability_1`, `capability_2`, ... |
| how_it_works | process diagram, step icons | `flow_diagram`, `step_icon_1`, ... |
| roadmap | roadmap or milestone icon | `roadmap_icon` |
| guarantee | guarantee item icons | `guarantee_1`, `guarantee_2`, ... |
| company | company logo, team photo, office photo | `company_logo`, `company_photo` |

List slots:

```bash
ppt-skill list-slots
```

## Conversation Guidance

Keep the conversation short and task-oriented.

First detect:

- Did the user provide a `.pptx` template?
- Did the user provide an outline?
- Did the user only provide a topic?
- Did the user specify purpose, audience, style, page count, or language?

If no template is provided and the theme is unclear, ask one concise question or infer
from the use case when safe.

Example when the user only gives a topic:

```text
我会先按“标准方案型 PPT”生成一个 10-12 页大纲，再制作 PPT。当前没有模板，我会使用内置商务主题。大纲会根据主题内容动态调整，不固定页数。
```

Example when the user provides a template and outline:

```text
我会按你提供的大纲制作内容，并填充进这个 PPT 模板。模板负责视觉风格，大纲负责页数、顺序和内容。
```

## CLI Reference

```bash
# List themes and image slots
ppt-skill list-themes
ppt-skill list-slots

# Template mode
ppt-skill create \
  --template brand.pptx \
  --slides /path/to/slides.json \
  --images '{"cover_logo":"/path/logo.png"}' \
  --output out.pptx

# Built-in theme mode
ppt-skill create \
  --theme purple \
  --slides /path/to/slides.json \
  --images '{}' \
  --output out.pptx
```

`--template` and `--theme` are mutually exclusive at the workflow level. If `--template`
is provided, template mode wins.

`--slides` can be a JSON file path or an inline JSON array.

`--images` can be a JSON object string or a JSON file path. Use `{}` when no images are
available.

## Slides JSON Format

The calling Agent must organize content as a slides JSON array. The array order is the
PPT page order. Page count is determined by array length.

Supported slide types follow.

### cover

```json
{"type":"cover","title":"Title","subtitle":"Subtitle"}
```

Image slots: `cover_bg`, `cover_logo`

### problem

```json
{"type":"problem","title":"Current Challenges","subtitle":"Optional subtitle","items":["Point 1","Point 2","Point 3"]}
```

### solution

```json
{"type":"solution","title":"Solution Overview","items":[
  {"problem":"Problem A","solution":"Solution A"},
  {"problem":"Problem B","solution":"Solution B"}
],"tagline":"One-sentence value proposition"}
```

Image slot: `solution_diagram`

### capabilities

```json
{"type":"capabilities","title":"Core Capabilities","items":[
  {"name":"Capability A","desc":"Short description"},
  {"name":"Capability B","desc":"Short description"}
]}
```

Image slots: `capability_1`, `capability_2`, ...

Short descriptions render as cards. Longer descriptions render as list-style content.

### how_it_works

```json
{"type":"how_it_works","title":"How It Works","items":[
  {"title":"Step 1","desc":"Description"},
  {"title":"Step 2","desc":"Description"}
]}
```

Image slots: `flow_diagram`, `step_icon_1`, `step_icon_2`, ...

### roadmap

```json
{"type":"roadmap","title":"Roadmap","items":[
  {"name":"Phase 1","time":"Timeline","items":["Milestone 1","Milestone 2"]},
  {"name":"Phase 2","time":"Timeline","items":["Milestone 3","Milestone 4"]}
]}
```

Image slot: `roadmap_icon`

### budget

```json
{"type":"budget","title":"Budget","items":[
  ["Item","Amount","Notes"],
  ["Development","500k","Development cost"],
  ["Total","500k","Total investment"]
]}
```

The first row is used as the table header.

### pricing

`pricing` is rendered with the same table layout as `budget`.

```json
{"type":"pricing","title":"Pricing","items":[
  ["Plan","Price","Scope"],
  ["Basic","100k","Core features"]
]}
```

### value

```json
{"type":"value","title":"Expected Value","items":[
  {"number":"80%","label":"Efficiency improvement","desc":"Description"},
  {"number":"30%","label":"Cost reduction","desc":"Description"}
]}
```

### next_steps

```json
{"type":"next_steps","title":"Next Steps","items":["Action 1","Action 2"],"contact":"Contact info"}
```

### guarantee

```json
{"type":"guarantee","title":"Guarantee System","items":[
  {"title":"Quality guarantee","desc":"Description"},
  {"title":"Service guarantee","desc":"Description"}
]}
```

Image slots: `guarantee_1`, `guarantee_2`, ...

### delivery

`delivery` is rendered with the same structure as `guarantee`.

```json
{"type":"delivery","title":"Delivery Plan","items":[
  {"title":"Deliverable A","desc":"Description"}
]}
```

### company

```json
{"type":"company","title":"About Us","items":["Company intro 1","Company intro 2"]}
```

Image slots: `company_logo`, `company_photo`

## Quality Rules

- Preserve the user's outline when provided.
- Do not force all decks into a product-proposal structure.
- Do not use a fixed page count for topic-only requests.
- Keep slide text concise. Prefer 3-5 points per page.
- Split dense sections into multiple pages instead of overloading one slide.
- Use placeholders for missing facts rather than fabricating data.
- Use ASCII-safe CLI output markers such as `[OK]`, `[WARN]`, and `[ERROR]`.
- Avoid emoji in documentation, code, and CLI output.

## Dependencies

```bash
uv pip install python-pptx
```

If `ppt-skill` is not installed as a command in the current Agent environment, run the
CLI module directly from this skill directory, for example:

```bash
python scripts/ppt_skill/cli.py create --theme blue --slides slides.json --images '{}' --output out.pptx
```
