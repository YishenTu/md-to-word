# Repository Guidelines

## Architecture

- `src/core/` orchestrates `MarkdownPreprocessor` → `PandocProcessor` → `WordPostprocessor`.
- `src/parsers/` extracts structured Markdown data without depending on DOCX APIs.
- `src/formatters/` owns Word layout and OpenXML formatting by document element type.
- `md_to_word.py` owns CLI policy and atomic output publication; `examples/` must remain self-contained.

Keep processors composable. Put Markdown interpretation in preprocessors or parsers and Word presentation in formatters. Internal control tokens must be explicit, narrowly scoped, and removed by the consuming pipeline stage; structured data such as signatures belongs in metadata instead of text markers.

## Style and maintenance

- Keep production code compatible with Python 3.11, even when developing on a newer interpreter.
- Treat `pyproject.toml` as the source of truth for formatting, linting, and type-checking configuration; use type hints for new or changed interfaces.
- Raise project-specific exceptions from `src/utils/exceptions.py` at pipeline boundaries.
- Use Conventional Commits when asked to commit: `feat:`, `fix:`, `docs:`, `refactor:`, or `chore:`.

## Documentation ownership

- Keep `README.md` human-facing: installation, usage, Markdown authoring contracts, configuration, visible behavior, failure behavior, and contributor commands.
- Keep detailed architecture, configuration, and security explanations in the matching file under `docs/`.
- Keep agent-only ownership boundaries, implementation constraints, verification rules, and safety invariants in `AGENTS.md`; do not put task state, agent workflow, or temporary decisions in human-facing documentation.
- Update every affected documentation owner when input contracts, configuration, layout, CI, architecture, security, or failure behavior changes.

## Validation

- Install development tools: `python -m pip install -r requirements.txt -r requirements-dev.txt`.
- Run lint and formatting checks: `ruff check .` and `ruff format --check .`.
- Run type checking: `mypy`.
- Audit runtime dependencies: `pip-audit -r requirements.txt`.
- Check the runtime environment: `python md_to_word.py --check-config`.
- Run all tests: `python -m unittest discover -v`.
- Convert the bundled example and visually inspect the result after layout changes.

Test DOCX behavior through OpenXML properties rather than binary diffs. Pandoc-dependent tests skip automatically when Pandoc is unavailable; keep all other tests independent of external applications. GitHub Actions is the authoritative cross-version check for Python 3.11 and 3.14.

## Security

- Invoke Pandoc with an argument list through `subprocess.run`; never build shell command strings.
- Normalize paths with the existing validators. Apply a resolved-path boundary only when an operation has an explicit base directory.
- Treat missing local assets and Pandoc warnings as conversion failures; never silently publish incomplete output.
- Respect the Obsidian environment variables documented in `docs/configuration.md`; never hardcode machine-specific paths.
