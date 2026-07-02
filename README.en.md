# ppt-skill

`ppt-skill` is a general-purpose skill for creating PowerPoint presentations.

You can think of it as a “PPT creation assistant”. Give it a topic, an outline, or a company PPT template, and it helps produce a `.pptx` file.

It is not limited to product proposals. It can also be used for project reports, training decks, pitch decks, event plans, company introductions, competitive analysis, work summaries, and business reviews.

## What It Can Do

- Create a complete PPT from a user-provided outline.
- Create a suggested outline first when the user only provides a topic.
- Use a company or brand PPT template when one is provided.
- Generate a PPT with built-in themes when no template is provided.
- Generate a deck even when no images are available.
- Add logos, screenshots, product images, or diagrams when the user provides them.
- Support common slide types such as cover pages, pain points, solutions, capabilities, process pages, roadmaps, tables, key numbers, next steps, and company introductions.

## Suitable Use Cases

You can use it for:

- Product proposals
- Customer solutions
- Project reports
- Training decks
- Pitch decks
- Event plans
- Company introductions
- Competitive analysis
- Annual summaries
- Business reviews

## How It Decides What To Do

The skill first checks whether the user has provided a PPT template or an outline.

If the user provides a PPT template, the template is used first. The template controls the visual style, including colors, backgrounds, layouts, and brand appearance.

If the user provides an outline, the PPT follows that outline. The outline controls what to say, how many pages to create, and the order of the pages.

If the user only provides a short topic, such as “Create a PPT about reducing customer service costs with AI”, the skill first helps create a suggested outline. The page count is not fixed. Simple topics can be shorter, while complex topics can be longer.

A simple way to understand it:

```text
PPT template: decides how it looks
PPT outline: decides what it says
PPT generator: turns the content into a .pptx file
```

## About Images

This skill does not generate images by itself.

If there are no images, it can still create a PPT using text, colors, cards, tables, process layouts, and big-number pages.

If the user provides logos, product images, screenshots, diagrams, or team photos, those images can be inserted into the PPT.

If another tool has already created images, those local image files can also be used.

## Built-In Themes

When the user does not provide a template, built-in themes can be used.

| Theme | Best for |
|---|---|
| `default` | General business reports and internal briefings |
| `blue` | Technology, AI, data, and platform topics |
| `green` | Health, sustainability, agriculture, and public welfare |
| `orange` | Sales proposals, marketing campaigns, and growth plans |
| `purple` | Branding, design, consulting, and premium services |
| `red` | Government-facing materials, annual meetings, and formal events |
| `teal` | Healthcare, education, and service design |

## Common Ways To Use It

### 1. User Provides an Outline

The user can provide an outline like this:

```text
Please create a PPT based on this outline:

1. Cover: How hyaluronic acid masks differ from ordinary masks
2. Pain point: Why ordinary hydrating masks only work briefly
3. Definition: What is an ordinary mask
4. Definition: What is a hyaluronic acid mask
5. Key differences
...
```

The skill follows the outline and does not force the deck into a fixed number of pages.

### 2. User Provides Only a Topic

The user can also say:

```text
Create a PPT about reducing customer service costs with AI.
```

In this case, the skill first creates a suggested outline. It usually starts with a standard report length, but the page count can change based on the topic.

### 3. User Provides a Template

If the user provides a company PPT template, that template is used first:

```text
Please use this company template and create a project report PPT.
```

The template controls the visual style. The outline or topic controls the content.

## Command Line Usage

If you use this skill locally or in another Agent environment, you can generate PPT files from the command line.

Install the dependency:

```bash
uv pip install python-pptx
```

List available themes:

```bash
ppt-skill list-themes
```

Generate with a built-in theme:

```bash
ppt-skill create \
  --theme purple \
  --slides slides.json \
  --images '{}' \
  --output output.pptx
```

Generate with a PPT template:

```bash
ppt-skill create \
  --template brand-template.pptx \
  --slides slides.json \
  --images '{}' \
  --output output.pptx
```

If the `ppt-skill` command is not installed in the current environment, run the script directly:

```bash
python scripts/ppt_skill/cli.py create --theme blue --slides slides.json --images '{}' --output output.pptx
```

## Content Format

When creating a PPT, the content is organized into page data. Regular users usually do not need to write this by hand; an Agent can prepare it automatically.

A simple example:

```json
[
  {
    "type": "cover",
    "title": "How Hyaluronic Acid Masks Differ From Ordinary Masks",
    "subtitle": "Understanding the difference and the opportunity"
  },
  {
    "type": "problem",
    "title": "Pain Point",
    "items": [
      "Hydrating masks are used often, but skin still feels dry",
      "Skin feels hydrated for a short time, then returns to normal",
      "Ordinary masks mainly moisturize the surface"
    ]
  }
]
```

Supported pages include cover, background, pain points, solution overview, capabilities, process steps, roadmaps, budget tables, pricing tables, key numbers, next steps, guarantees, delivery plans, and company introductions.

For detailed formats, see [SKILL.md](SKILL.md) and [references/page_types.md](references/page_types.md).

## Tips

- Provide a company template if you have one.
- Provide an outline if the page structure is already clear.
- If you only have a topic, you can still start with that.
- Provide real data when accuracy matters; do not ask the tool to guess factual numbers.
- Split dense content into multiple pages instead of putting everything on one slide.
- Provide local image file paths if images need to be inserted.

## Project Structure

```text
ppt-skill/
├── README.md
├── README.en.md
├── LICENSE
├── SKILL.md
├── references/
│   ├── default_themes.md
│   └── page_types.md
└── scripts/
    └── ppt_skill/
        ├── cli.py
        ├── template_fill.py
        ├── template_parser.py
        └── __init__.py
```

## License

MIT. See [LICENSE](LICENSE).
