# Markdown to MathType Word

A Windows-oriented agent skill that converts scientific Markdown into a Word `.docx` whose formulas are **editable MathType `Equation.DSMT4` OLE objects** — not plain text, images, or Word OMML equations. Text and math styles default to Times New Roman 11 pt.

> This repository root is the skill itself: `SKILL.md` is the agent-facing instruction file, and `scripts/` contains the conversion pipeline.

## What it does

1. Builds an OMML intermediate DOCX from a research-paper subset of Markdown (headings, paragraphs, bullets, basic HTML tables, `\(...\)` inline math, `\[...\]` display math, fractions, roots, Greek letters, scripts, common operators, simple aligned equations).
2. Converts the OMML equations into native MathType objects through Word/MathType automation, in configurable batches.
3. Verifies the result: zero remaining OMML equations, matching `Equation.DSMT4` OLE count, and Times New Roman declarations inside every MathType object.

The source Markdown is never overwritten. The converter refuses to overwrite existing outputs or reuse work paths, and keeps intermediate/checkpoint artifacts for audit.

## Requirements

- Windows with **Microsoft Word desktop** and **64-bit MathType 7** installed. Both must be closed before conversion; the tool never kills user-owned Office processes.
- Python 3 with `python-docx` and `lxml`:

  ```bat
  py -m pip install -r requirements.txt
  ```

- For agent use, a client that supports the `SKILL.md` format (for example WorkBuddy or Claude Code).

## Install as a skill

Clone this repository directly into your skills directory:

```bat
:: WorkBuddy (user-level)
git clone https://github.com/<YOUR_NAME>/markdown-to-mathtype-word.git "%USERPROFILE%\.workbuddy\skills\markdown-to-mathtype-word"
```

```bash
# Claude Code
git clone https://github.com/<YOUR_NAME>/markdown-to-mathtype-word.git ~/.claude/skills/markdown-to-mathtype-word
```

Then ask the agent to use the `markdown-to-mathtype-word` skill. The bundled `agents/openai.yaml` provides a display name and a default prompt for clients that read it.

## Direct command-line use

```bat
scripts\Convert-MarkdownToMathTypeWord.bat "C:\path\paper.md" "C:\path\paper_MathType_TNR11.docx"
```

- Optional third argument: MathType conversion batch size (default `10`).
- Optional PowerShell flags: `-PythonExe`, `-IntermediateOnly` (OMML diagnostic pass only).

```powershell
.\scripts\Convert-MarkdownToMathTypeWord.ps1 -MarkdownPath "C:\path\paper.md" -IntermediateOnly
```

## Standalone verification

```bat
py scripts\verify_mathtype_docx.py "C:\path\paper_MathType_TNR11.docx"
```

It checks ZIP integrity, OLE / `Equation.DSMT4` counts, zero residual OMML, and the Times New Roman declaration in every MathType object. Always open the final DOCX in Word and spot-check inline, display, fraction, script, and unit formulas before delivery.

## Limitations

- This is a **research-paper profile**, not a general Pandoc/LaTeX replacement. Unsupported TeX commands abort the run instead of silently producing altered formulas.
- Requires real Word + MathType licenses; those applications are not included in this repository.
- The bundled `MathType_Times_Symbol_11.eqp` is a MathType preference file for Times New Roman 11 pt equation styles.

## License

MIT — see [LICENSE](LICENSE). Word and MathType are commercial products of their respective owners.
