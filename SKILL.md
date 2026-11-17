---
name: markdown-to-mathtype-word
description: "Convert scientific Markdown into a Word document whose formulas are editable MathType Equation.DSMT4 objects, with Times New Roman 11-point primary text and math styles. Use for Windows Word plus MathType 7 workflows, not for image-only or ordinary-text equations."
---

# Markdown To Mathtype Word

Convert a scientific Markdown file to a `.docx` without rasterizing or flattening formulas. The final validation must report zero native OMML equations and a matching number of `Equation.DSMT4` OLE objects.

## Preconditions

- Microsoft Word desktop and 64-bit MathType 7 must be installed. Require the user to close both before starting; never terminate a user-owned Word or MathType process.
- Input can have a `.md`/`.markdown` extension or be an extensionless Markdown file. Math uses `\(...\)` for inline formulas and `\[...\]` for display formulas. The bundled builder supports the research-paper subset used here: headings, paragraphs, bullets, basic HTML tables, fractions, roots, Greek letters, scripts, common operators, and simple aligned equations.
- Do not use this workflow for arbitrary untested LaTeX. If the builder reports an unsupported TeX command, stop and inspect the intermediate rather than delivering a possibly altered formula.

## Run

Run the portable entry point in `scripts`:

```bat
Convert-MarkdownToMathTypeWord.bat "C:\path\paper.md" "C:\path\paper_MathType_TNR11.docx"
```

The optional third argument is the MathType batch size; `10` is the default. The PowerShell controller preserves the Markdown source, creates an OMML intermediate and validated checkpoints beside the requested output, and refuses to overwrite any existing artifact.

For an OMML-only diagnostic pass, use:

```powershell
.\Convert-MarkdownToMathTypeWord.ps1 -MarkdownPath "C:\path\paper.md" -IntermediateOnly
```

## Required result

- Ordinary Word text and styles are Times New Roman, 11 pt.
- Formula objects are editable MathType `Equation.DSMT4` OLE objects, not plain text, pictures, or Word OMML equations.
- The bundled MathType preference file sets Text, Function, Variable, Vector, and Number styles to Times New Roman at a full equation size of 11 pt. Symbol and MT Extra remain only where MathType requires them for Greek letters and mathematical operators.
- The converter normalizes empty/NBSP unit spacing to MathType `mspace`, avoiding the known literal-question-mark defect.

## Verification and handoff

`verify_mathtype_docx.py` checks ZIP integrity, OLE count, `Equation.DSMT4` count, zero remaining OMML, and the Times New Roman declaration in every MathType object. Review the generated audit JSON and open the final file in Word to inspect representative inline, display, fraction, script, and unit formulas before delivery. Do not claim a formula is MathType merely because it looks like one.
