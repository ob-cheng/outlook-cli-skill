# Skill Structure & Token Efficiency

## Directory layout

```
outlook-cli-skill/
├── SKILL.md              # Entry point — triggers & core instructions
├── outlook.py            # CLI entry point
├── outlook_cli/          # CLI implementation
│   ├── cli.py
│   ├── core/
│   ├── services/
│   └── utils/
├── docs/
│   ├── install.md        # Agent install & setup instructions
│   └── features.md       # Human-facing feature guide
├── references/
│   ├── commands.md
│   ├── direct-send.md
│   ├── features.md
│   ├── json-schemas.md
│   ├── scripts.md
│   ├── structure.md      # This file
│   ├── troubleshooting.md
│   ├── workflows.md
│   └── wsl.md
├── scripts/
│   ├── validate-export.py
│   └── format-email.py
├── assets/
│   └── email-template.md
├── requirements.txt
└── README.md
```

## Token Efficiency

| Component | Context cost | Notes |
|-----------|-------------|-------|
| SKILL.md | ~140 lines | Loaded when skill triggers |
| Reference files | 0 until loaded | Loaded on-demand |
| Scripts (validate-export.py, format-email.py) | 0 until run | Code never enters context — only stdout |
| CLI code (outlook_cli/) | 0 | Never enters context — runs via shell |
