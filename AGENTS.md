# Repository Guidelines

## Architecture

- `src/core/` orchestrates `MarkdownPreprocessor` → `PandocProcessor` → `WordPostprocessor`.
- `src/parsers/` extracts structured Markdown data without depending on DOCX APIs.
- `src/formatters/` owns Word layout and OpenXML formatting by document element type.
- `src/config/` contains shared layout settings; `src/utils/` contains validation, exceptions, and reusable helpers.
- `md_to_word.py` owns CLI policy and atomic output publication; `examples/` must remain self-contained.

Keep processors composable. Put Markdown interpretation in preprocessors or parsers and Word presentation in formatters. Internal control tokens must be explicit, narrowly scoped, and removed by the consuming pipeline stage; structured data such as signatures belongs in metadata instead of text markers.

## Style and maintenance

- Target Python 3.11+, four-space indentation, and a 120-column limit.
- Follow Ruff formatting and lint rules from `pyproject.toml`; use type hints for new or changed interfaces.
- Use `snake_case` for modules, functions, and variables; `PascalCase` for classes; `UPPER_SNAKE_CASE` for constants.
- Raise project-specific exceptions from `src/utils/exceptions.py` at pipeline boundaries.
- Update `README.md` and relevant files in `docs/` whenever input contracts, configuration, layout, CI, or failure behavior changes.
- Use Conventional Commits when asked to commit: `feat:`, `fix:`, `docs:`, `refactor:`, or `chore:`.

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
