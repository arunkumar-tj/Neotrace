# CLAUDE.md

This file provides guidance for AI assistants (like Claude) working in this repository.

## Repository Overview

- **Repository**: arunkumar-tj/Claude
- **Status**: Newly initialized repository
- **Primary Branch**: `claude/add-claude-documentation-M4gm6`

> This CLAUDE.md should be updated as the project evolves with actual source code, dependencies, and tooling.

---

## Git Conventions

### Branch Naming
- Feature branches: `feature/<short-description>`
- Bug fixes: `fix/<short-description>`
- AI-assisted branches: `claude/<description>-<session-id>`
- Documentation: `docs/<short-description>`

### Commit Messages
- Use the imperative mood: "Add feature" not "Added feature"
- Keep subject lines under 72 characters
- Reference issues when applicable: `Fix #123: ...`
- Include a session URL at the end of AI-generated commits

### Push Rules
- Always push with tracking: `git push -u origin <branch-name>`
- Never push directly to `main` or `master` without a pull request
- Branch names must match the assigned development branch exactly

### Retry Logic for Network Failures
If a push fails due to network errors, retry with exponential backoff:
- Attempt 1: immediate
- Attempt 2: wait 2s
- Attempt 3: wait 4s
- Attempt 4: wait 8s
- Attempt 5: wait 16s

---

## Development Setup

> Fill in the following sections once the project stack is determined.

### Prerequisites
- [ ] Document required runtime versions (Node.js, Python, Go, etc.)
- [ ] Document required global tools
- [ ] Document environment variables (use `.env.example` as a template, never commit `.env`)

### Installation
```bash
# Example — replace with actual commands
git clone <repo-url>
cd Claude
# install dependencies here
```

### Running the Project
```bash
# Example — replace with actual commands
# npm start / python main.py / go run . / etc.
```

### Running Tests
```bash
# Example — replace with actual commands
# npm test / pytest / go test ./... / etc.
```

### Building / Compiling
```bash
# Example — replace with actual commands
# npm run build / make build / etc.
```

---

## Code Style & Conventions

### General Principles
- Prefer clarity over cleverness
- Avoid over-engineering: implement the minimum needed for the current task
- Do not add unused imports, dead code, or commented-out code
- Keep functions small and focused on a single responsibility
- Validate at system boundaries (user input, external APIs); trust internal code

### Security
- Never commit secrets, credentials, or API keys
- Sanitize all user input before use
- Avoid common vulnerabilities: SQL injection, XSS, command injection, path traversal
- Use parameterized queries for database access
- Follow OWASP Top 10 guidelines

### Error Handling
- Handle errors explicitly; avoid silent failures
- Only add error handling for scenarios that can realistically occur
- Log errors with enough context to debug without exposing sensitive data

---

## AI Assistant Instructions

### When Making Changes
1. Always read a file before editing it
2. Understand existing patterns before introducing new ones
3. Only change what was explicitly requested — avoid "while I'm here" refactoring
4. Run linting/tests after making changes (once tooling is configured)
5. Commit with a descriptive message and push to the designated branch

### What to Avoid
- Adding docstrings, comments, or type annotations to untouched code
- Creating unnecessary abstractions or helper utilities for one-off operations
- Adding backwards-compatibility shims for removed code
- Hardcoding configuration values that should come from environment variables
- Introducing dependencies without justification

### File Creation Policy
- Prefer editing existing files over creating new ones
- Do not create README files or documentation unless explicitly requested
- Create new files only when strictly necessary for the task

### Task Management
- Use the TodoWrite tool for multi-step tasks (3+ steps)
- Mark tasks `in_progress` before starting them
- Mark tasks `completed` immediately after finishing them
- Only one task should be `in_progress` at a time

---

## Project Structure

> Update this section once source code exists.

```
Claude/
├── CLAUDE.md          # This file — AI assistant guidance
├── .git/              # Git internals
└── ...                # Source files to be added
```

---

## CI/CD

> Document pipeline details here once CI/CD is configured (GitHub Actions, GitLab CI, etc.).

---

## Contact & Contribution

- Repository owner: `arunkumar-tj`
- For issues, open a GitHub issue on the repository
- All contributions should go through pull requests against the main branch
