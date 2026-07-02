# Page Types And Outline Strategy

`ppt-skill` renders a slides JSON array. The array order is the PPT page order, and the
array length is the PPT page count.

Do not force every deck into one fixed proposal structure. Page selection should follow
the user's outline when provided, or the Agent's dynamic outline when the user only gives
a topic.

---

## Outline Strategy

When the user provides an outline:

- Preserve the user's page order and page count.
- Expand sparse pages into concise slide-ready content when needed.
- Do not invent unsupported facts, figures, prices, legal claims, medical claims, or
  financial claims.

When the user only provides a topic:

- Generate a suggested outline first unless the user explicitly asks to directly create
  the deck.
- Choose page count dynamically based on purpose, audience, complexity, and user
  constraints.
- Default to a 10-12 page standard outline when no other constraints exist.

Page-count bands:

| Band | Page count | Suitable for |
| --- | ---: | --- |
| Lightweight | 5-7 | quick intro, meeting sync, short briefing |
| Standard | 8-12 | proposal, project report, solution deck, customer discussion |
| Detailed | 13-20 | training deck, roadshow, consulting-style report, detailed plan |
| Deep | 20+ | course deck, whitepaper-style deck, complete business plan |

Recommended outline item format:

```text
1. Page title
   Purpose: what this page needs to communicate
   Key points: 2-4 concise points
   Type: one supported page type
```

---

## Supported Page Types

### cover

Use for opening pages.

Content:

- Main title
- Optional subtitle
- Optional date, company, author, or logo if available

JSON:

```json
{"type":"cover","title":"Title","subtitle":"Subtitle"}
```

Image slots: `cover_bg`, `cover_logo`

### problem

Use for background, challenge, pain points, current status, or why change is needed.

Content:

- Page title
- Optional subtitle
- 3-5 concise items

JSON:

```json
{"type":"problem","title":"Challenges","subtitle":"Optional subtitle","items":["Point 1","Point 2","Point 3"]}
```

### solution

Use for solution overview, before-after comparison, problem-solution mapping, or core
proposal.

Content:

- Page title
- Problem-solution rows
- Optional tagline

JSON:

```json
{"type":"solution","title":"Solution Overview","items":[
  {"problem":"Problem A","solution":"Solution A"},
  {"problem":"Problem B","solution":"Solution B"}
],"tagline":"One-sentence value proposition"}
```

Image slot: `solution_diagram`

### capabilities

Use for capabilities, features, advantages, modules, pillars, or service items.

Content:

- Page title
- Multiple items with `name` and `desc`

JSON:

```json
{"type":"capabilities","title":"Core Capabilities","items":[
  {"name":"Capability A","desc":"Short description"},
  {"name":"Capability B","desc":"Short description"}
]}
```

Image slots: `capability_1`, `capability_2`, ...

### how_it_works

Use for workflows, mechanisms, process explanation, method, or implementation logic.

Content:

- Page title
- 3-5 steps with title and description

JSON:

```json
{"type":"how_it_works","title":"How It Works","items":[
  {"title":"Step 1","desc":"Description"},
  {"title":"Step 2","desc":"Description"}
]}
```

Image slots: `flow_diagram`, `step_icon_1`, `step_icon_2`, ...

### roadmap

Use for timelines, implementation plans, rollout plans, milestones, or phases.

Content:

- Page title
- Multiple phases with name, time, and milestone items

JSON:

```json
{"type":"roadmap","title":"Roadmap","items":[
  {"name":"Phase 1","time":"Timeline","items":["Milestone 1","Milestone 2"]},
  {"name":"Phase 2","time":"Timeline","items":["Milestone 3","Milestone 4"]}
]}
```

Image slot: `roadmap_icon`

### budget

Use for budgets, investment plans, cost breakdowns, resource plans, or tables.

Content:

- Page title
- Table rows
- First row as header

JSON:

```json
{"type":"budget","title":"Budget","items":[
  ["Item","Amount","Notes"],
  ["Development","500k","Development cost"]
]}
```

### pricing

Use for packages, pricing tiers, quotes, or commercial plans. Rendered like `budget`.

JSON:

```json
{"type":"pricing","title":"Pricing","items":[
  ["Plan","Price","Scope"],
  ["Basic","100k","Core features"]
]}
```

### value

Use for expected benefits, quantified value, KPI targets, ROI, or success metrics.

Content:

- Page title
- Big-number items

JSON:

```json
{"type":"value","title":"Expected Value","items":[
  {"number":"80%","label":"Efficiency improvement","desc":"Description"},
  {"number":"30%","label":"Cost reduction","desc":"Description"}
]}
```

### next_steps

Use for action plans, next steps, cooperation process, or close pages.

JSON:

```json
{"type":"next_steps","title":"Next Steps","items":["Action 1","Action 2"],"contact":"Contact info"}
```

### guarantee

Use for risks and mitigations, service guarantees, quality assurance, governance, or
support systems.

JSON:

```json
{"type":"guarantee","title":"Guarantee System","items":[
  {"title":"Quality guarantee","desc":"Description"},
  {"title":"Service guarantee","desc":"Description"}
]}
```

Image slots: `guarantee_1`, `guarantee_2`, ...

### delivery

Use for deliverables, delivery model, work packages, or implementation outputs.
Rendered like `guarantee`.

JSON:

```json
{"type":"delivery","title":"Delivery Plan","items":[
  {"title":"Deliverable A","desc":"Description"}
]}
```

### company

Use for company introduction, team, qualifications, cases, or credibility building.

JSON:

```json
{"type":"company","title":"About Us","items":["Company intro 1","Company intro 2"]}
```

Image slots: `company_logo`, `company_photo`

---

## Mapping Guidance

Use the closest available page type rather than inventing a new type. Examples:

- Market background -> `problem`
- Current situation -> `problem`
- Strategic goal -> `value` or `capabilities`
- Product features -> `capabilities`
- Architecture -> `solution` or `how_it_works`
- Process -> `how_it_works`
- Implementation plan -> `roadmap`
- Cost, quote, package -> `budget` or `pricing`
- Risk and mitigation -> `guarantee`
- Deliverables -> `delivery`
- Case studies -> `company` or `capabilities`, depending on content
