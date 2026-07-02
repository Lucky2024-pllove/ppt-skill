# Built-In Themes

These themes are used only when the user does not provide a `.pptx` template.

Agent rule:

- If the user explicitly chooses a theme, use it.
- If the user provides a clear use case or style, the Agent may choose a suitable theme
  and briefly state the choice.
- If the user gives no style signal and does not say "you decide", ask one concise
  question before choosing.
- The CLI has a `default` fallback for robustness, but Agents should not treat it as a
  mandatory product rule.

---

## Theme List

### default - Default Business Blue

```text
Primary:    #1E2761
Secondary:  #CADCFC
Accent:     #F96167
Background: #FFFFFF
Title font: Calibri
Body font:  Calibri
Chinese:    Microsoft YaHei
Best for:   general business reports, internal briefings, neutral proposals
```

### blue - Technology Blue

```text
Primary:    #0066CC
Secondary:  #E8F4FD
Accent:     #FF6B35
Background: #FFFFFF
Title font: Arial
Body font:  Calibri
Chinese:    Microsoft YaHei
Best for:   technology, SaaS, AI, data, platform, engineering topics
```

### green - Forest Green

```text
Primary:    #2C5F2D
Secondary:  #97BC62
Accent:     #D4A017
Background: #F5F5F5
Title font: Georgia
Body font:  Calibri
Chinese:    Microsoft YaHei
Best for:   sustainability, agriculture, health, public welfare, long-term growth
```

### orange - Warm Business Orange

```text
Primary:    #E8611A
Secondary:  #FDE8D8
Accent:     #2D3436
Background: #FFFFFF
Title font: Arial Black
Body font:  Calibri
Chinese:    Microsoft YaHei
Best for:   sales proposals, campaigns, operations, growth, commercial plans
```

### purple - Elegant Purple

```text
Primary:    #6C3483
Secondary:  #E8DAEF
Accent:     #F1C40F
Background: #FFFFFF
Title font: Cambria
Body font:  Calibri
Chinese:    Microsoft YaHei
Best for:   brand, design, consulting, premium services, strategic narratives
```

### red - China Red

```text
Primary:    #C0392B
Secondary:  #FADBD8
Accent:     #F39C12
Background: #FDFEFE
Title font: Impact
Body font:  Calibri
Chinese:    Microsoft YaHei
Best for:   government-facing materials, ceremonies, annual meetings, formal local contexts
```

### teal - Teal

```text
Primary:    #028090
Secondary:  #00A896
Accent:     #02C39A
Background: #FDFDFD
Title font: Trebuchet MS
Body font:  Calibri
Chinese:    Microsoft YaHei
Best for:   healthcare, education, service design, modern but restrained topics
```

---

## Theme Data Format

```json
{
  "name": "default",
  "label": "Default Business Blue",
  "colors": {
    "primary": "#1E2761",
    "secondary": "#CADCFC",
    "accent": "#F96167",
    "background": "#FFFFFF",
    "text": "#333333",
    "text_light": "#666666"
  },
  "fonts": {
    "title": "Calibri",
    "body": "Calibri",
    "title_ea": "微软雅黑",
    "body_ea": "微软雅黑"
  },
  "slide_size": {
    "width": 13.333,
    "height": 7.5
  }
}
```
