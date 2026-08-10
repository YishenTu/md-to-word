# Architecture

## Scope

The project converts body-oriented Markdown documents into DOCX files using the core page geometry, typography, paragraph layout, and page-number conventions of GB/T 9704-2012. It intentionally does not model every optional official-document element.

## Pipeline

```text
Markdown file
  -> MarkdownPreprocessor and SignatureBlockParser
  -> PandocProcessor
  -> WordPostprocessor and specialized formatters
  -> DOCX validation
  -> atomic publication
```

### Markdown input stage

`SignatureBlockParser` inspects only the strict terminal document region after note metadata is removed. It accepts a signatory and date either at the body end or immediately before a terminal attachment declaration. When it finds an unambiguous block, it moves the values into normalized metadata and leaves one position-only signature anchor in the Markdown body. The parser has no dependency on `python-docx`; this keeps Markdown grammar separate from presentation logic.

`MarkdownPreprocessor` handles transformations required by this output profile: optional frontmatter and ending metadata removal, attachment normalization into explicitly marked item paragraphs, caption positioning, list normalization, heading policy, and body separators. The paragraph formatter consumes the attachment markers after Pandoc conversion and applies the shared name position and per-item hanging indent. Fenced code blocks are opaque during these transforms, and Markdown soft breaks remain Pandoc's responsibility.

Body separators, signature positions, and normalized attachment items become explicit, narrowly scoped control tokens. `WordPostprocessor` consumes and removes them. The signature anchor carries no document values; the structured signatory and date remain in metadata.

### Pandoc stage

`PandocProcessor` owns the external process invocation and temporary Markdown input. It enables MathML, disables implicit figures, preserves tabs, disables wrapping, and fails on Pandoc warnings. Arguments are passed as a list without shell interpretation.

The Markdown source directory is passed as a Pandoc resource path so standard relative images resolve independently of the process working directory.

### Word stage

`WordPostprocessor` loads Pandoc's DOCX and coordinates formatters in a deterministic order:

1. `PageFormatter` applies page geometry, the document grid, default spacing, and mirrored page numbers.
2. `ParagraphFormatter` applies body and heading layout and guarantees attachment spacing.
3. `DocumentTitleFormatter` inserts the metadata-provided title, which comes from the sole Markdown level-one heading or falls back to the filename.
4. `ListFormatter` and `TableFormatter` normalize native Word structures.
5. The postprocessor consumes page-break sentinels and resolves Obsidian image syntax.
6. `ImageFormatter` sizes images, applies wrapping, and collapses anchor paragraphs before captions.
7. `SignatureFormatter` replaces the consumed signature anchor with normalized Word paragraphs at the authored position.

`BaseFormatter` owns shared run-font and document-grid helpers. All run formatting assigns a Chinese font and also assigns Times New Roman to the Latin font slots, regardless of content level.

### Publication stage

The CLI writes to a staging DOCX in the destination directory. It completes Pandoc conversion and Word formatting, reopens the package with `python-docx`, preserves an existing destination's file mode, and calls `os.replace` only after validation. An existing output therefore survives preprocessing, Pandoc, image, formatting, or package-validation failures.

## Ownership rules

- Core processors own pipeline stages, orchestration, and error boundaries.
- Parsers own structured Markdown recognition and normalization.
- Formatters own one DOCX presentation concern each.
- `DocumentConfig` owns layout constants and derived printable dimensions.
- The CLI owns user interaction, input limits, output-path policy, and atomic publication.
- Missing local assets are errors and are never silently omitted.

The parser and formatter for signatures deliberately remain separate. The parser owns terminal-region recognition and authored position; the formatter owns OpenXML layout. Merging them would couple Markdown grammar to Word presentation and make both harder to test or replace.

## Testing and CI

Unit tests cover preprocessing, parsing, path policy, formatter-level OpenXML, and CLI behavior. Pipeline tests run Pandoc in temporary directories and assert semantic DOCX/OpenXML properties instead of binary equality. The bundled example is also converted end to end and asserts attachments, lists, images, page breaks, title spacing, and signatures.

GitHub Actions separates quality checks from runtime tests. Ruff, mypy, and `pip-audit` run once on Python 3.14; the full test suite and configuration validator run with Pandoc on Python 3.11 and 3.14. Pandoc-dependent local tests skip only when the executable is unavailable.
