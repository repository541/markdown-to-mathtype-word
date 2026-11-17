<#
  Converts native Word OMML formulas in a disposable DOCX copy into editable
  MathType Equation.DSMT4 OLE objects using MathType's installed native API.

  This intentionally avoids Word's MathType ribbon macro dispatcher.  The
  native steps mirror MathType's own Word template: OMML -> MathML through
  MathType's OMML2MML.XSL, a DSEquations OLE object, MTSetEqnFromLangStr, and
  MTCloseOleObject.  The supplied .eqp is applied to every newly created
  formula; its text/variable/function/number/vector styles are Times New
  Roman at a full equation size of 11 pt.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $SourceDocx,
    [Parameter(Mandatory = $true)] [string] $OutputDocx,
    [string] $PreferenceFile,
    [int[]] $EquationIndexes,
    [ValidateRange(1, 10000)] [int] $SaveEvery = 10
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($PreferenceFile)) {
    $PreferenceFile = Join-Path (Split-Path -Parent $PSCommandPath) 'MathType_Times_Symbol_11.eqp'
}

$diagnosticPath = [Environment]::GetEnvironmentVariable('MATHTYPE_DIAG')
function Write-Diagnostic {
    param([string] $Message)
    if (-not [string]::IsNullOrWhiteSpace($diagnosticPath)) {
        [IO.File]::AppendAllText($diagnosticPath, "$(Get-Date -Format o) $Message`r`n")
    }
}
Write-Diagnostic 'converter-started'

$MathTypeXsl = 'C:\Program Files (x86)\MathType\Office Support\64\OMML2MML.XSL'
$MathTypeWll = 'C:\Program Files (x86)\MathType\MathPage\64\MathPage.wll'

function Get-DocxObjectCounts {
    param([Parameter(Mandatory = $true)] [string] $Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $documentXml = $archive.Entries | Where-Object FullName -eq 'word/document.xml'
        if ($null -eq $documentXml) { throw "word/document.xml is missing: $Path" }
        $reader = [IO.StreamReader]::new($documentXml.Open())
        try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        [pscustomobject]@{
            NativeOmml = [regex]::Matches($xml, '<m:oMath(?:\s|>)').Count
            MathTypeOle = @($archive.Entries | Where-Object FullName -like 'word/embeddings/oleObject*.bin').Count
            MathTypeProgId = [regex]::Matches($xml, 'ProgID="Equation\.DSMT4"').Count
        }
    }
    finally { $archive.Dispose() }
}

function Release-ComObjectSafely {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function Test-TaggedDisplayEquation {
    param(
        [Parameter(Mandatory = $true)] $Document,
        [Parameter(Mandatory = $true)] $Range
    )

    # The layout builder uses a leading and trailing tab only for the 27
    # centered, right-numbered display equations.  This avoids treating
    # ordinary inline formulas as display equations when applying preferences.
    $before = if ($Range.Start -gt $Document.Content.Start) {
        $Document.Range($Range.Start - 1, $Range.Start).Text
    } else { '' }
    $after = if ($Range.End -lt $Document.Content.End) {
        $Document.Range($Range.End, $Range.End + 1).Text
    } else { '' }
    return ($before -eq "`t" -and $after -eq "`t")
}

if (-not (Test-Path -LiteralPath $SourceDocx -PathType Leaf)) {
    throw "Source document not found: $SourceDocx"
}
if (-not (Test-Path -LiteralPath $PreferenceFile -PathType Leaf)) {
    throw "MathType preference file not found: $PreferenceFile"
}
if (-not (Test-Path -LiteralPath $MathTypeXsl -PathType Leaf) -or -not (Test-Path -LiteralPath $MathTypeWll -PathType Leaf)) {
    throw 'The installed 64-bit MathType conversion components were not found.'
}
if (Test-Path -LiteralPath $OutputDocx) {
    throw "Refusing to overwrite an existing output file: $OutputDocx"
}
if (@(Get-Process WINWORD,MathType -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'Close Word and MathType before this disposable conversion session is started.'
}
$preExistingMathTypePids = @(Get-Process MathType -ErrorAction SilentlyContinue | ForEach-Object Id)

$sourceCounts = Get-DocxObjectCounts -Path $SourceDocx
if ($sourceCounts.NativeOmml -lt 1) {
    throw 'The source document contains no native OMML formulas to convert.'
}

Copy-Item -LiteralPath $SourceDocx -Destination $OutputDocx
Write-Diagnostic 'source-copied'

$word = $null
$document = $null
$styleXml = $null
$sourceXml = $null
$apiStarted = $false
$converted = 0
$displayConverted = 0
$inlineConverted = 0
$initialCount = 0
try {
    Write-Diagnostic 'before-word-create'
    $word = New-Object -ComObject Word.Application
    Write-Diagnostic 'word-created'
    # OLE activation needs a real Word window.  It stays minimized and is
    # closed at the end of this isolated conversion session.
    $word.Visible = $true
    # MathType OLE activation is more reliable with a normal, visible host
    # window than with a hidden or minimized Word instance.
    $word.WindowState = 0 # wdWindowStateNormal
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($OutputDocx, $false, $false)
    Write-Diagnostic 'document-opened'
    $document.Activate()
    $word.ActiveWindow.View.Type = 3 # wdPrintView
    if (-not ('MathTypeNativeConverter' -as [type])) {
        $interopSourcePath = Join-Path (Split-Path -Parent $PSCommandPath) 'MathTypeNativeConverter.cs'
        Add-Type -TypeDefinition ([IO.File]::ReadAllText($interopSourcePath))
    }
    Write-Diagnostic 'native-interop-ready'

    $initialCount = [int]$document.OMaths.Count
    if ($initialCount -ne $sourceCounts.NativeOmml) {
        throw "Word found $initialCount OMML formulas but the package count is $($sourceCounts.NativeOmml)."
    }
    if ($PSBoundParameters.ContainsKey('EquationIndexes')) {
        $targetIndexes = @($EquationIndexes | Sort-Object -Unique)
        foreach ($index in $targetIndexes) {
            if ($index -lt 1 -or $index -gt $initialCount) {
                throw "Equation index $index is outside the document's 1-$initialCount range."
            }
        }
    } else {
        $targetIndexes = 1..$initialCount
    }
    $targetSet = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($index in $targetIndexes) { [void]$targetSet.Add([int]$index) }

    $styleXml = New-Object -ComObject Msxml2.DOMDocument.6.0
    $styleXml.async = $false
    if (-not $styleXml.Load($MathTypeXsl)) {
        throw "Could not load MathType's OMML2MML stylesheet: $($styleXml.parseError.reason)"
    }

    $prefs = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $PreferenceFile))
    $initStatus = [MathTypeNativeConverter]::MTInitAPI([int16]0, [int16]30)
    Write-Diagnostic "mt-init-status=$initStatus"
    if ($initStatus -lt 0) { throw "MTInitAPI failed with status $initStatus." }
    $apiStarted = $true

    $missing = [Type]::Missing
    for ($index = $initialCount; $index -ge 1; $index--) {
        if (-not $targetSet.Contains($index)) { continue }
        Write-Diagnostic "formula-$index-start"

        $math = $document.OMaths.Item($index)
        $range = $math.Range
        $start = [int]$range.Start
        $isDisplay = Test-TaggedDisplayEquation -Document $document -Range $range
        $wordOpenXml = [string]$range.WordOpenXML

        $sourceXml = New-Object -ComObject Msxml2.DOMDocument.6.0
        $sourceXml.async = $false
        if (-not $sourceXml.LoadXML($wordOpenXml)) {
            throw "Formula $index could not be read as OMML: $($sourceXml.parseError.reason)"
        }
        $mathML = [string]$sourceXml.transformNode($styleXml.documentElement)
        Release-ComObjectSafely $sourceXml
        $sourceXml = $null
        if ([string]::IsNullOrWhiteSpace($mathML) -or $mathML -notmatch '<(?:mml:)?math\b') {
            throw "Formula $index could not be transformed to MathML."
        }

        # During OMML round-tripping, Word may serialize a non-breaking unit
        # separator either as NBSP or as an empty mtext token.  MathType's
        # MathML importer renders both variants as a literal question mark.
        # Its installed MathML translator maps mspace to a true spacing
        # object, preserving the intended unit separation without a glyph.
        $mathML = [regex]::Replace(
            $mathML,
            '<mml:mtext(?:\s+[^>]*)?>(?:\s|\u00A0|&#xA0;|&#160;)*</mml:mtext>',
            '<mml:mspace width="0.167em"/>'
        )

        # Mode 1 is the MathType display preference; mode 5 adds the inline
        # flag.  Both use the same TNR-11 font/size scheme from the .eqp file.
        $preferenceMode = if ($isDisplay) { [int16]1 } else { [int16]5 }
        $prefsStatus = [MathTypeNativeConverter]::MTSetMTPrefs($preferenceMode, $prefs, [int16]-1)
        Write-Diagnostic "formula-$index-prefs-status=$prefsStatus"
        if ($prefsStatus -ne 0) { throw "MTSetMTPrefs failed for formula $index with status $prefsStatus." }

        # Delete exactly the native formula range and insert the OLE formula
        # at the same character position.  Processing from end to start keeps
        # all still-pending formula indexes and locations stable.
        [void]$range.Delete()
        $insertRange = $document.Range($start, $start)
        $shape = $document.InlineShapes.AddOLEObject('DSEquations', '', $false, $false, $missing, $missing, $missing, $insertRange)
        Write-Diagnostic "formula-$index-ole-created"
        $shape.OLEFormat.DoVerb(2)
        Write-Diagnostic "formula-$index-ole-activated"
        $oleObject = $shape.OLEFormat.Object
        $setStatus = [MathTypeNativeConverter]::SetMathML($oleObject, $mathML)
        Write-Diagnostic "formula-$index-set-status=$setStatus"
        if ($setStatus -ne 0) { throw "MTSetEqnFromLangStr failed for formula $index with status $setStatus." }
        $closeStatus = [MathTypeNativeConverter]::MTCloseOleObject(1, $oleObject)
        Write-Diagnostic "formula-$index-close-status=$closeStatus"
        if ($closeStatus -ne 0) { throw "MTCloseOleObject failed for formula $index with status $closeStatus." }

        Release-ComObjectSafely $oleObject
        $oleObject = $null
        Release-ComObjectSafely $shape
        $shape = $null
        Release-ComObjectSafely $insertRange
        Release-ComObjectSafely $range
        Release-ComObjectSafely $math

        $converted++
        if ($isDisplay) { $displayConverted++ } else { $inlineConverted++ }
        if (($converted % $SaveEvery) -eq 0) { $document.Save() }
        Write-Diagnostic "formula-$index-complete"
    }

    $remaining = [int]$document.OMaths.Count
    if ($remaining -ne ($initialCount - $converted)) {
        throw "Unexpected remaining OMML count: $remaining after converting $converted of $initialCount."
    }
    $document.Save()
    Write-Diagnostic 'document-saved'
}
finally {
    Release-ComObjectSafely $sourceXml
    Release-ComObjectSafely $styleXml
    if ($null -ne $document) { $document.Close(0); Release-ComObjectSafely $document }
    if ($null -ne $word) { $word.Quit(); Release-ComObjectSafely $word }
    # Do not call MTTermAPI here.  In this MathType build its native teardown
    # can terminate the hosting PowerShell process after Word has already
    # saved a valid checkpoint.  Each conversion batch runs in an isolated
    # worker, so releasing the Word COM session and allowing that worker to
    # exit safely unloads the WLL/API connection.
    # Any MathType process created after the preflight was launched solely by
    # this isolated worker (the preflight refuses user-owned MathType). Close
    # it explicitly so the next checkpoint starts from a clean session.
    foreach ($mathTypeProcess in @(Get-Process MathType -ErrorAction SilentlyContinue)) {
        if ($preExistingMathTypePids -notcontains $mathTypeProcess.Id) {
            try { Stop-Process -Id $mathTypeProcess.Id -Force -ErrorAction Stop } catch { }
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$outputCounts = Get-DocxObjectCounts -Path $OutputDocx
$expectedRemaining = $sourceCounts.NativeOmml - $converted
 $expectedOle = $sourceCounts.MathTypeOle + $converted
if ($outputCounts.NativeOmml -ne $expectedRemaining -or $outputCounts.MathTypeOle -ne $expectedOle -or $outputCounts.MathTypeProgId -ne $expectedOle) {
    throw "Package verification failed: OMML=$($outputCounts.NativeOmml) (expected $expectedRemaining); OLE=$($outputCounts.MathTypeOle) (expected $expectedOle); Equation.DSMT4=$($outputCounts.MathTypeProgId) (expected $expectedOle)."
}

[pscustomobject]@{
    Output = (Resolve-Path -LiteralPath $OutputDocx).Path
    SourceOmml = $sourceCounts.NativeOmml
    Converted = $converted
    DisplayEquationsConverted = $displayConverted
    InlineEquationsConverted = $inlineConverted
    RemainingOmml = $outputCounts.NativeOmml
    MathTypeOleObjects = $outputCounts.MathTypeOle
    EquationDSMT4Objects = $outputCounts.MathTypeProgId
} | Format-List
