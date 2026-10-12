# Skills

Agent skills by [Daniel Netzer](https://www.danielnetzer.com).

Things I've built to make work better—from tactile interfaces to cloud cost models. Each skill includes its instructions, references and any examples or scripts it needs.

## Install

Choose a skill from the collection:

```sh
npx skills add DanielNetzer/skills
```

Or install one directly:

```sh
npx skills add DanielNetzer/skills --skill tactile-ui
npx skills add DanielNetzer/skills --skill cloud-cost-projection
```

Add `--agent cursor` to target Cursor. To inspect the available skills without installing:

```sh
npx skills add DanielNetzer/skills --list
```

## The collection

| Skill | What it helps with | Try it |
| --- | --- | --- |
| [Tactile UI](skills/tactile-ui) | Glass materials, soft shapes and responsive controls, with keyboard, reduced-motion and fallback checks. | [Interactive demo](https://www.danielnetzer.com/skills/tactile-ui) |
| [Cloud Cost Projection](skills/cloud-cost-projection) | An invoice-reconciled cloud cost workbook: scenarios, commitments, cash timing and a check against the budget cap. | [Method and download](https://www.danielnetzer.com/skills/cloud-cost-projection) |

These are instructions and supporting tools for coding agents. Read a skill's `SKILL.md` before using it. Cloud Cost Projection's workbook generator needs Python and `openpyxl`; its verification workflow also uses LibreOffice. The sample config is illustrative, not a real company bill.

## Use without the CLI

Copy a folder from `skills/` into your agent's skill directory. Keep its references, examples and scripts alongside `SKILL.md`. For Codex, that can be `~/.codex/skills/<skill-name>/`.

## Structure

```text
skills/
├── tactile-ui/
│   ├── SKILL.md
│   ├── references/
│   ├── examples/
│   └── agents/
└── cloud-cost-projection/
    ├── SKILL.md
    ├── references/
    ├── examples/
    └── scripts/
```

## Changes and contributions

Found a problem or tried a skill on something interesting? Open an issue with the task, what happened and what you expected. Keep credentials, private company data and agent transcripts out of reports.

Each new skill belongs in its own folder under `skills/`, with a clear trigger, useful instructions and a working example when appropriate. Keep claims tied to checks that actually ran.

This collection is the home for future updates. The original standalone repositories remain available so existing links and installs continue to work. See [migration notes](MIGRATION.md).

[Website](https://www.danielnetzer.com) · [X](https://x.com/NetzerDaniel) · [LinkedIn](https://www.linkedin.com/in/daniel-netzer/)

MIT © Daniel Netzer. Individual skills retain their license and source attribution.
