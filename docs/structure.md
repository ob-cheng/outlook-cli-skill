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
│   ├── features.md       # Human-facing feature guide
│   ├── wsl.md            # WSL setup guide
│   ├── scripts.md        # Utility scripts and SKILL_DIR resolution
│   └── structure.md      # This file
├── references/
│   ├── commands.md
│   ├── direct-send.md
│   ├── features.md
│   ├── json-schemas.md
│   ├── troubleshooting.md
│   └── workflows.md
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
