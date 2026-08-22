# Configuration Reference

All user preferences are stored in `~/.outlook-cli/config.json` and managed via the CLI `config` subcommand.

## Settings Overview

| Key | Default | Description |
|---|---|---|
| `send_mode` | `draft` | `draft` or `send` — controls whether the `--send` flag is allowed |
| `draft_instructions` | `""` | Free-text drafting rules for the agent (e.g. "Keep it brief, mention next steps") |
| `humanizer_enabled` | `false` | Whether the agent runs the humanizer skill on email bodies before saving/sending |

## CLI Commands

```bash
# View all current settings
python outlook.py config show

# Enable direct sending (requires --send flag too)
python outlook.py config set send_mode send

# Revert to draft-only
python outlook.py config set send_mode draft

# Set custom drafting instructions (agent follows these verbatim)
python outlook.py config set draft_instructions "Keep it under 3 sentences"

# Enable humanizer processing
python outlook.py config set humanizer_enabled true

# Disable humanizer processing
python outlook.py config set humanizer_enabled false

# Reset a setting to default
python outlook.py config clear draft_instructions
```

## Send Mode Behavior

| send_mode | `--send` flag? | Result |
|---|---|---|
| `draft` | — | ✅ Saved as draft |
| `draft` | ✅ | ❌ Error: "send not allowed" |
| `send` | — | ✅ Saved as draft |
| `send` | ✅ | ✅ Sent immediately |

Key rule: **`--send` is the only way to direct-send.** Without it, even with `send_mode: send`, the email is saved as a draft. No way to accidentally send.

## Draft Instructions

When `draft_instructions` is set to non-empty text, the agent reads it before composing any email and follows those rules while writing the body. The instructions are saved **verbatim** (the agent never rewrites or interprets them) and are read from config on every send/reply/forward via `config show`.

Examples:
- `"Keep it brief, bullet points only"`
- `"Always include a call to action at the end"`
- `"Mention the Q3 deadline and link to the shared drive"`

## Humanizer Integration

When `humanizer_enabled` is `true`, the agent loads the [humanizer skill](https://github.com/blader/humanizer) after composing the email body and runs its 29-pattern checklist to strip AI writing patterns (filler phrases, copula avoidance, em dash spam, signposting, etc.). The result is a draft that reads like a real person wrote it.

If the humanizer skill is not installed in the agent's environment, the agent asks the user whether to install it. If declined, the draft is saved as-is with a note that humanization was skipped.

### The `--humanized` gate

The humanizer is a reasoning transform — there is no script to run it, and the CLI can't verify whether a body reads as human. So the CLI enforces the *protocol*, not the prose: when `humanizer_enabled` is on, `send`/`reply`/`forward` **fail** unless `--humanized` is passed.

```bash
# humanizer_enabled = true
python outlook.py send --to x@co.com --subject Hi --body "..."              # ✗ humanizer_required
python outlook.py send --to x@co.com --subject Hi --body "..." --humanized  # ✓ proceeds
```

The error (`code: "humanizer_required"` in `--json`) converts a silent skip into a loud, logged failure. `--humanized` is an assertion the step ran, not verification that it did — an agent could pass it dishonestly, but it can no longer skip the step *silently*. Applies to draft and direct-send alike, and to each `send`/`reply`/`forward` inside `batch`.

## Status Tags

On every send/reply/forward, the CLI prints status tags to stdout showing what's active:

```
[send mode: draft]
[draft instructions enabled]
[humanizer enabled]
```

The agent sees these tags in the command output. If a setting is expected but no tag appears, the agent skipped a step.

With `--json`, the same tags are returned as a `status_tags` array in the success payload (the human-readable lines above are suppressed in JSON mode), so the audit signal survives the programmatic path.

## First-Time Setup

On first activation, the agent follows the setup flow in [docs/install.md](../docs/install.md) to ask the user about these settings.
