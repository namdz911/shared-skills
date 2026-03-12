# Shared Skills

Reusable [Claude Code](https://claude.ai/claude-code) skills for teaching automation. Each skill is a self-contained folder with a skill file (`.md`), supporting scripts, and references.

## Skills

| Skill | Description |
|-------|-------------|
| [grade-excel-pset](grade-excel-pset/) | Grade individual Excel problem-set submissions using automated scoring + 4-agent ensemble review |

## How to Use

1. **Copy** the skill folder into your project
2. **Read** the skill's `README.md` for setup and prerequisites
3. **Invoke** the skill in Claude Code (e.g., `/grade-pset <assignment-dir>`)

Each skill folder contains everything needed: the skill file that Claude Code reads, any scripts it depends on, and reference documents.

## Prerequisites

- [Claude Code](https://claude.ai/claude-code)
- Python 3.11+
- Skill-specific dependencies listed in each skill's README
