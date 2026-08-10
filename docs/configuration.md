# Configuration

## Requirements

- Python 3.11 or newer
- Pandoc available on `PATH`
- `python-docx==1.2.0`

Run `python3 md_to_word.py --check-config` to display detected dependencies, image paths, fonts, and warnings. Font and Obsidian-path warnings do not fail validation; missing Pandoc, an unsupported Python version, or a mismatched runtime dependency does.

## Document layout

All layout defaults live in `src/config/config.py` under `DocumentConfig`.

| Setting | Default |
| --- | --- |
| Page | A4 portrait, 210 mm × 297 mm |
| Margins | 34.58 mm top, 32.58 mm bottom, 28 mm left, 26 mm right |
| Body grid | 28 characters × 22 lines |
| Body line pitch | 10.39 mm |
| Document-title line pitch | 12.51 mm |
| Body first-line indent | 32 pt, equivalent to two size-three characters |
| Paragraph spacing | 0 pt before and after |

Printable width and height are derived from the page size and margins. Image width, character-grid spacing, and signature positioning use those derived values rather than duplicating physical measurements.

The page setup uses the electronic-document margin values associated with GB/T 33476.2-2016 while the resulting typography, grid, and page-number conventions target GB/T 9704-2012. These values are configuration defaults, not a claim that every optional official-document element is implemented.

## Fonts and sizes

The default profile uses:

- `FZXiaoBiaoSong-B05S` at 22 pt for the document title.
- `FangSong` at 16 pt for body text and signatures.
- `SimHei` at 16 pt for first-level headings.
- `STKaiTi` at 16 pt for second-level headings.
- `SimSun` at 14 pt for page numbers.
- `FangSong` at 12 pt for tables and captions.
- `Times New Roman` for every Latin letter and number, including titles, headings, body text, lists, tables, captions, signatures, and page numbers.

Word or WPS substitutes fonts that are not installed. Install the configured fonts on every machine that must produce reproducible output, or change the names to fonts available in that environment.

## Markdown input conventions

### Frontmatter and title

YAML frontmatter is optional and filtered before conversion. When the document contains exactly one Markdown level-one heading, its visible text becomes the internal document title and that heading is removed from the body. Otherwise, the title falls back to the source filename. Multiple level-one headings cause the authored heading hierarchy to shift down by one level. The generated internal title uses the configured size-two `xiaobiaosong` font and centered title layout.

### Page breaks

An independent `---`, `***`, or `___` line in the body creates a page break. A leading pair of `---` lines is treated as YAML frontmatter boundaries instead. Do not use a body separator where a visible horizontal rule is expected.

### Ordered and unordered lists

Ordered list markers are preserved as visible body text rather than Word auto-numbering. This provides deterministic fonts and indentation across Word and WPS.

Markdown `-` and `*` items become native Word bullet lists with a controlled `•` marker. The first-level marker starts two body-character cells from the text-area edge, item text starts one additional cell to the right, and wrapped lines align with the first item character. Nested levels retain the marker and move right by two character cells per level. List paragraphs use body typography, zero paragraph spacing, and the body document grid.

### Attachment declarations

Write `附件：` on its own line, followed by a blank line and a contiguous ordered list. The preprocessor marks each item as a separate attachment paragraph. The paragraph formatter aligns all item names at one fixed character-grid position, so wrapped lines begin under the first character of the name while later item numbers remain aligned. It also guarantees one grid-aligned blank line before the declaration. Authored alignment spaces, tabs, or extra blank lines are unnecessary.

### Stamped-document signature

Provide a single authority and document date either at the end of the Markdown body or immediately before a terminal attachment declaration:

```markdown
The final body paragraph.

Example Municipal Government Office
2026-07-14
```

The signatory must be the nearest non-empty line above the date, and a blank line must separate the signature block from prior body content. Blank lines between the signatory and date are accepted and normalized during conversion. The date must be the final body line or be followed only by a terminal attachment declaration. Trailing `---`, single-word tags such as `#work`, and `Date:` note metadata are removed before signature recognition. The date accepts ISO or Chinese syntax and is emitted without zero-padded month or day values. A `Date` field in YAML frontmatter remains filtered note metadata and does not create a signature block.

The converter preserves the authored order between the signature and attachment declaration. It inserts two body-grid blank lines before the generated signatory and keeps the spacing lines, signatory, and date together across pagination. The date is placed four body-character cells from the right edge. The signatory is centered against the date by sharing its character-grid layout region. Only text positioning is generated; no seal image is inserted or sized.

## Images

Standard Markdown images are resolved by Pandoc against the Markdown source directory. Obsidian embeds are searched in this order:

1. Markdown source directory
2. DOCX staging directory
3. Configured Obsidian attachment directory and vault
4. Local `images`, `assets`, and current directories

Supported local extensions are `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.svg`, and `.webp`. Remote Obsidian embeds are not downloaded. Missing local embeds and Pandoc resource warnings fail the conversion so incomplete output cannot be mistaken for a successful document.

Images are expanded to the configured printable width while preserving aspect ratio. Top-and-bottom wrapping is enabled by default. Empty image anchor paragraphs use a collapsed exact line, preventing an extra blank line before a caption.

## Obsidian environment variables

- `OBSIDIAN_VAULT_PATH`: absolute path to a vault; takes precedence over discovery.
- `OBSIDIAN_VAULT_NAME`: vault name used for discovery in common macOS locations.
- `OBSIDIAN_ATTACHMENTS_FOLDER`: attachment folder within the vault.

The environment is read when `DocumentConfig` is imported. Set variables before starting the process. Repository examples keep their assets under `examples/assets/` and require no Obsidian configuration.

## Pandoc options

`DocumentConfig.PANDOC_CONFIG` selects MathML, fails on warnings, preserves tabs, disables source wrapping, and controls image wrapping. The converter also disables implicit figures so alt text remains accessibility metadata instead of becoming an unintended visible caption.

Changing these options can alter Pandoc's OpenXML structure and should be accompanied by pipeline tests.
