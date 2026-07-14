# Security and failure model

## Trust boundary

This project is a local CLI, not a sandbox for hostile Markdown. The user intentionally grants it access to the selected input, output, and configured image locations. Absolute paths and parent-directory components are therefore valid after normalization.

Callers that need a directory boundary can pass a base directory to `validate_safe_path`; the validator then resolves symlinks and requires the result to remain within that base. The normal CLI does not claim such a boundary around user-selected files.

## Command execution

Pandoc is invoked through `subprocess.run` with an argument list and without a shell. Markdown paths and resource paths are not interpreted as shell syntax. The converter does not run code blocks or embedded HTML.

Pandoc warnings are treated as failures. This is especially important for missing standard Markdown images, which might otherwise produce a DOCX that looks successful while omitting content.

## Paths and external resources

Input and output paths are expanded and normalized with `Path.resolve`. On Windows, invalid filename characters are rejected. Relative Obsidian images are resolved only from explicit local search roots; remote Obsidian embeds are not fetched.

Standard Markdown image handling is delegated to Pandoc with the Markdown source directory as its resource path. Users should treat Pandoc's parser and supported image decoders as part of the trusted local toolchain and keep them updated.

## Output integrity

The final path is restricted to the DOCX format. All work occurs in a uniquely named staging file in the destination directory. The completed package is reopened after formatting and atomically replaces the destination only after validation. Existing file permissions are retained when an output is replaced.

If preprocessing, Pandoc, image insertion, formatting, or validation fails, the staging file is removed and an existing destination remains unchanged. Missing images are fatal instead of being silently omitted.

## XML handling

DOCX elements are created through `python-docx` and `OxmlElement`. User-provided text is assigned as element text or attributes rather than interpolated into raw XML fragments. OpenXML manipulation remains implementation-sensitive, so tests assert the generated XML properties used by Word and WPS.

## Dependency assurance

Runtime and development dependencies are pinned in `requirements.txt` and `requirements-dev.txt`. CI runs `pip-audit` against runtime dependencies on every push and pull request. A clean audit reduces known-vulnerability risk but does not prove that Pandoc, image codecs, or document viewers are vulnerability-free.

Do not commit private source documents, generated DOCX files, credentials, local Vault paths, or unrelated screenshots. The repository intentionally tracks the self-contained Markdown example and its two image assets, plus the public standard reference used to verify layout decisions.
