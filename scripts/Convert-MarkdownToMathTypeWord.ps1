<#
  Convert scientific Markdown to a DOCX with editable MathType Equation.DSMT4
  objects.  The source Markdown and any existing DOCX are never overwritten.

  Markdown support intentionally matches the research-paper profile used in
  this project: headings, paragraphs, bullets, basic HTML tables, inline
  \(...\) equations, and display \[...\] equations.  It is not a general
  replacement for a full Markdown/Pandoc renderer.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)] [string] $MarkdownPath,
    [Parameter(Position = 1)] [string] $OutputDocx,
    [ValidateRange(1, 100)] [int] $BatchSize = 10,
    [string] $PythonExe,
    [switch] $IntermediateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$toolRoot = Split-Path -Parent $PSCommandPath
$builder = Join-Path $toolRoot 'build_scientific_md_omml.py'
$converter = Join-Path $toolRoot 'convert_omml_to_mathtype_native.ps1'
$verifier = Join-Path $toolRoot 'verify_mathtype_docx.py'
$preferences = Join-Path $toolRoot 'MathType_Times_Symbol_11.eqp'
$powerShell51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

function Resolve-PythonWithDocx {
    param([string] $Requested)

    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($Requested)) { $candidates.Add($Requested) }
    $codexPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    if (Test-Path -LiteralPath $codexPython) { $candidates.Add($codexPython) }
    $systemPython = Get-Command python -ErrorAction SilentlyContinue
    if ($null -ne $systemPython) { $candidates.Add($systemPython.Source) }

    foreach ($candidate in @($candidates | Select-Object -Unique)) {
        try {
            & $candidate -c 'import docx, lxml' *> $null
            if ($LASTEXITCODE -eq 0) { return (Resolve-Path -LiteralPath $candidate).Path }
        }
        catch { }
    }
    throw 'Python with python-docx and lxml was not found. Install those packages or supply -PythonExe.'
}

function Test-OfficeClosed {
    $open = @(Get-Process WINWORD, MathType -ErrorAction SilentlyContinue)
    if ($open.Count -gt 0) {
        $names = ($open | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
        throw "Close $names before conversion. The tool will not terminate user-owned Office processes."
    }
}

function Get-DocxInfo {
    param([string] $Path)
    $raw = & $script:python $verifier $Path --stage inspect --json
    if ($LASTEXITCODE -ne 0) { throw "DOCX verification could not inspect: $Path" }
    return ($raw | ConvertFrom-Json)
}

foreach ($required in @($builder, $converter, $verifier, $preferences, $powerShell51)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "Required workflow component was not found: $required"
    }
}

$input = (Resolve-Path -LiteralPath $MarkdownPath -ErrorAction Stop).Path

if ([string]::IsNullOrWhiteSpace($OutputDocx)) {
    $defaultName = '{0}_MathType_TNR11.docx' -f [IO.Path]::GetFileNameWithoutExtension($input)
    $OutputDocx = Join-Path (Split-Path -Parent $input) $defaultName
}
$output = [IO.Path]::GetFullPath($OutputDocx)
if ([IO.Path]::GetExtension($output).ToLowerInvariant() -ne '.docx') {
    throw 'OutputDocx must have a .docx extension.'
}
$outputDirectory = Split-Path -Parent $output
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    throw "Output folder does not exist: $outputDirectory"
}
if (Test-Path -LiteralPath $output) {
    throw "Refusing to overwrite existing output: $output"
}

$outputStem = [IO.Path]::GetFileNameWithoutExtension($output)
$intermediate = Join-Path $outputDirectory ($outputStem + '_OMML_intermediate.docx')
$checkpointDirectory = Join-Path $outputDirectory ($outputStem + '_MathType_checkpoints')
foreach ($protectedPath in @($intermediate, $checkpointDirectory)) {
    if (Test-Path -LiteralPath $protectedPath) {
        throw "Refusing to reuse existing work path: $protectedPath"
    }
}

$script:python = Resolve-PythonWithDocx -Requested $PythonExe
Test-OfficeClosed

Write-Host "1/3 Creating OMML intermediate: $intermediate"
$builderOutput = @(& $script:python $builder $input $intermediate 2>&1)
$builderExitCode = $LASTEXITCODE
$builderOutput | ForEach-Object { Write-Host $_ }
if ($builderExitCode -ne 0) { throw 'Markdown-to-OMML creation failed.' }
if (($builderOutput -join "`n") -match 'OMML parser warnings:') {
    throw "The Markdown contains unsupported TeX commands. Inspect the retained intermediate before correcting the source; no MathType final output was created."
}

$initial = Get-DocxInfo -Path $intermediate
if ([int]$initial.omml -lt 1) {
    throw 'The Markdown source produced no OMML equations. Use \(...\) or \[...\] delimiters; no text-only formula output was created.'
}
if ([int]$initial.ole_objects -ne 0 -or [int]$initial.equation_dsm_t4 -ne 0) {
    throw 'The intermediate DOCX unexpectedly contains pre-existing MathType objects.'
}
if ($IntermediateOnly) {
    & $script:python $verifier $intermediate --stage omml
    Write-Host "Intermediate-only mode complete: $intermediate"
    exit 0
}

Test-OfficeClosed
New-Item -ItemType Directory -Path $checkpointDirectory -ErrorAction Stop | Out-Null
$finalCandidate = Join-Path $checkpointDirectory 'final_candidate.docx'

Write-Host "2/3 Converting $($initial.omml) equations to editable MathType objects in batches of $BatchSize."
$remaining = [int]$initial.omml
$converted = 0
$batch = 0
$current = $intermediate
while ($remaining -gt 0) {
    $batch += 1
    $take = [Math]::Min($BatchSize, $remaining)
    $firstIndex = $remaining - $take + 1
    $indexes = @($firstIndex..$remaining)
    $next = if ($take -eq $remaining) {
        $finalCandidate
    }
    else {
        Join-Path $checkpointDirectory ('checkpoint_{0:D3}.docx' -f $batch)
    }

    Write-Host "  Batch ${batch}: equations $firstIndex-$remaining"
    $arguments = @(
        '-NoLogo', '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass',
        '-File', $converter,
        '-SourceDocx', $current,
        '-OutputDocx', $next,
        '-PreferenceFile', $preferences,
        '-EquationIndexes'
    ) + @($indexes | ForEach-Object { [string]$_ }) + @('-SaveEvery', [string]$take)
    & $powerShell51 @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "MathType conversion failed in batch $batch. Previous validated checkpoints remain in: $checkpointDirectory"
    }

    $after = Get-DocxInfo -Path $next
    $expectedRemaining = $remaining - $take
    $expectedOle = $converted + $take
    if ([int]$after.omml -ne $expectedRemaining -or
        [int]$after.ole_objects -ne $expectedOle -or
        [int]$after.equation_dsm_t4 -ne $expectedOle) {
        throw "Checkpoint $batch failed count validation: OMML=$($after.omml), OLE=$($after.ole_objects), Equation.DSMT4=$($after.equation_dsm_t4)."
    }
    $remaining = $expectedRemaining
    $converted = $expectedOle
    $current = $next
}

Write-Host '3/3 Verifying final MathType object structure and primary font declaration.'
& $script:python $verifier $finalCandidate --stage mathtype --expected-equations $initial.omml
if ($LASTEXITCODE -ne 0) { throw 'Final MathType verification failed.' }
Copy-Item -LiteralPath $finalCandidate -Destination $output -ErrorAction Stop
$final = Get-DocxInfo -Path $output
if ([int]$final.omml -ne 0 -or [int]$final.equation_dsm_t4 -ne [int]$initial.omml) {
    throw 'The final copy failed post-copy verification.'
}

$audit = [ordered]@{
    input_markdown = $input
    input_sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $input).Hash
    intermediate_docx = $intermediate
    output_docx = $output
    output_sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $output).Hash
    equations = [int]$initial.omml
    equation_dsm_t4 = [int]$final.equation_dsm_t4
    omml_remaining = [int]$final.omml
    primary_math_font = 'Times New Roman'
    primary_math_size_pt = 11
    preference_file = $preferences
    batch_size = $BatchSize
    checkpoint_directory = $checkpointDirectory
    generated_utc = [DateTime]::UtcNow.ToString('o')
}
$auditPath = [IO.Path]::ChangeExtension($output, '.audit.json')
$audit | ConvertTo-Json | Set-Content -LiteralPath $auditPath -Encoding UTF8
Write-Host "Success: $output"
Write-Host "Audit: $auditPath"
