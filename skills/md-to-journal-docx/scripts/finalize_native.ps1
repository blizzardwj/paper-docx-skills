[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$SourcePath,
    [Parameter(Mandatory)] [string]$OutputDirectory,
    [Parameter(Mandatory)] [string]$OutputStem,
    [string]$ComProgId = 'KWps.Application',     # WPS (default) or Word.Application
    [switch]$Force
)

# Native save through the host: open the builder's DOCX read-only, update internal fields twice
# (PAGE/NUMPAGES/REF/SEQ…), SaveAs DOCX + DOC, reopen both read-only and record metrics.
# Run in an interactive Windows logon session. No MathType/OLE objects are created.
# Existing final outputs are replaced only when -Force is supplied.
$ErrorActionPreference = 'Stop'
$hostLabel = if ($ComProgId -like 'KWps*') { 'WPS Office' } else { 'Microsoft Word' }
$automationBefore = @(Get-Process wps, winword -ErrorAction SilentlyContinue | ForEach-Object Id)
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Release-ComReference($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-NormalizedPath([string]$Path) {
    return [IO.Path]::GetFullPath($Path)
}

function Test-SamePath([string]$First, [string]$Second) {
    return [string]::Equals((Get-NormalizedPath $First), (Get-NormalizedPath $Second), [StringComparison]::OrdinalIgnoreCase)
}

function Disable-OptionalDocumentCompatibilityCheck($Document, [Collections.Generic.List[string]]$Warnings) {
    # Some Word COM interfaces do not expose this optional property.
    # Owned-instance DisplayAlerts=0 remains the primary save-dialog suppression.
    try {
        $Document.CheckCompatibility = $false
        return $true
    } catch {
        $Warnings.Add('Optional Document.CheckCompatibility could not be set; continuing with owned-instance DisplayAlerts=0. ' + $_.Exception.Message)
        return $false
    }
}

function Get-OmmlAudit([string]$Path) {
    $package = [IO.Compression.ZipFile]::OpenRead($Path)
    $parts = [Collections.Generic.List[object]]::new()
    $total = 0
    $mainCount = 0
    try {
        if ($null -eq $package.GetEntry('word/document.xml')) { throw 'Missing word/document.xml in DOCX.' }
        foreach ($entry in $package.Entries) {
            if ($entry.FullName -notmatch '^word/(document|header\d+|footer\d+|footnotes|endnotes|comments\d*)\.xml$') { continue }
            $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
            try { $content = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $xml = [Xml.XmlDocument]::new()
            $xml.PreserveWhitespace = $true
            $xml.XmlResolver = $null
            $xml.LoadXml($content)
            $namespaces = [Xml.XmlNamespaceManager]::new($xml.NameTable)
            $namespaces.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
            $count = [int]$xml.SelectNodes('//m:oMath', $namespaces).Count
            $paragraphCount = [int]$xml.SelectNodes('//m:oMathPara', $namespaces).Count
            $total += $count
            if ($entry.FullName -eq 'word/document.xml') { $mainCount = $count }
            $parts.Add([ordered]@{ part = $entry.FullName; omath_count = $count; omath_paragraph_count = $paragraphCount })
        }
    } finally { $package.Dispose() }
    return [ordered]@{
        total_omath_count = $total
        main_document_omath_count = $mainCount
        scope = 'Main document, headers, footers, footnotes, endnotes and comments XML parts'
        parts = $parts.ToArray()
    }
}

function Update-InternalFields($Document) {
    # Avoid DATE/TIME changes, external LINK/INCLUDETEXT/DDE retrieval and TOC rewrites.
    $safeCommands = @('PAGE', 'NUMPAGES', 'SECTIONPAGES', 'SECTION', 'REF', 'PAGEREF', 'NOTEREF', 'STYLEREF', 'SEQ')
    $records = [Collections.Generic.List[object]]::new()
    $stories = $null
    try {
        $stories = $Document.StoryRanges
        foreach ($storyType in @(1, 2, 3, 5, 6, 7, 8, 9, 10, 11)) {
            $range = $null
            try { $range = $stories.Item($storyType) } catch { continue }
            $chainIndex = 0
            # WPS: NextStoryRange can keep returning the same range instead of null -> stop on repeats
            $seenRanges = [Collections.Generic.HashSet[string]]::new()
            try {
                while ($null -ne $range) {
                    $chainIndex++
                    if ($chainIndex -gt 10000) { throw 'Unexpected story-range chain length.' }
                    $rangeKey = "$([int]$range.StoryType):$([int]$range.Start):$([int]$range.End):$(([string]$range.Text).GetHashCode())"
                    if (-not $seenRanges.Add($rangeKey)) { break }
                    $fields = $null
                    try {
                        $fields = $range.Fields
                        for ($fieldIndex = 1; $fieldIndex -le $fields.Count; $fieldIndex++) {
                            $field = $null
                            $codeRange = $null
                            try {
                                $field = $fields.Item($fieldIndex)
                                $codeRange = $field.Code
                                $code = [string]$codeRange.Text
                                $command = ''
                                if ($code -match '^\s*([A-Za-z]+)') { $command = $Matches[1].ToUpperInvariant() }
                                if ($safeCommands -notcontains $command) { continue }
                                $record = [ordered]@{
                                    story_type = $storyType; story_index = $chainIndex
                                    field_index = $fieldIndex; command = $command
                                    code = $code.Trim(); status = 'pending'; update_result = $null
                                }
                                if ($field.Locked) { $record.status = 'locked_skipped' }
                                else {
                                    try {
                                        $record.update_result = [bool]$field.Update()
                                        $record.status = if ($record.update_result) { 'updated' } else { 'update_returned_false' }
                                    } catch {
                                        $record.status = 'update_error'
                                        $record.error = $_.Exception.Message
                                    }
                                }
                                $records.Add($record)
                            } finally {
                                Release-ComReference $codeRange
                                Release-ComReference $field
                            }
                        }
                    } finally { Release-ComReference $fields }
                    $next = $range.NextStoryRange
                    Release-ComReference $range
                    $range = $next
                }
            } finally { Release-ComReference $range }
        }
    } finally { Release-ComReference $stories }
    return $records.ToArray()
}

function Get-DocumentMetrics($Document) {
    $Document.Repaginate()
    $inlineShapes = $null
    $shapes = $null
    $sections = $null
    $tables = $null
    $maths = $null
    $inlineTypes = [ordered]@{}
    $floatingTypes = [ordered]@{}
    $tableMetrics = [Collections.Generic.List[object]]::new()
    $inlineShapeMetrics = [Collections.Generic.List[object]]::new()
    $oleObjects = [Collections.Generic.List[object]]::new()
    $inlinePictures = 0
    $floatingPictures = 0
    try {
        $sections = $Document.Sections
        $tables = $Document.Tables
        $maths = $Document.OMaths
        $inlineShapes = $Document.InlineShapes
        $shapes = $Document.Shapes
        for ($index = 1; $index -le $tables.Count; $index++) {
            $table = $null
            $tableRange = $null
            $startRange = $null
            $endRange = $null
            $rows = $null
            $columns = $null
            $metricErrors = [Collections.Generic.List[string]]::new()
            $metric = [ordered]@{
                index = $index
                start_page = $null
                end_page = $null
                spans_pages = $null
                row_count = $null
                column_count = $null
                metadata_errors = @()
            }
            try {
                $table = $tables.Item($index)
                try { $rows = $table.Rows; $metric.row_count = [int]$rows.Count }
                catch { $metricErrors.Add('Read row count: ' + $_.Exception.Message) }
                try { $columns = $table.Columns; $metric.column_count = [int]$columns.Count }
                catch { $metricErrors.Add('Read column count: ' + $_.Exception.Message) }
                try {
                    $tableRange = $table.Range
                    $startPosition = [int]$tableRange.Start
                    $endPosition = [Math]::Max($startPosition, [int]$tableRange.End - 1)
                    # Duplicate ranges before collapsing; the table/document range is never edited.
                    $startRange = $tableRange.Duplicate
                    $endRange = $tableRange.Duplicate
                    $startRange.SetRange($startPosition, $startPosition)
                    $endRange.SetRange($endPosition, $endPosition)
                    $metric.start_page = [int]$startRange.Information(3)
                    $metric.end_page = [int]$endRange.Information(3)
                    $metric.spans_pages = $metric.start_page -ne $metric.end_page
                } catch { $metricErrors.Add('Read first/last page: ' + $_.Exception.Message) }
            } finally {
                Release-ComReference $columns
                Release-ComReference $rows
                Release-ComReference $endRange
                Release-ComReference $startRange
                Release-ComReference $tableRange
                Release-ComReference $table
            }
            $metric.metadata_errors = $metricErrors.ToArray()
            $tableMetrics.Add($metric)
        }
        for ($index = 1; $index -le $inlineShapes.Count; $index++) {
            $shape = $null
            $ole = $null
            $shapeRange = $null
            $pageRange = $null
            try {
                $shape = $inlineShapes.Item($index)
                $type = [int]$shape.Type
                $key = [string]$type
                if (-not $inlineTypes.Contains($key)) { $inlineTypes[$key] = 0 }
                $inlineTypes[$key]++
                if ($type -in @(3, 4)) { $inlinePictures++ }
                $shapeMetric = [ordered]@{
                    index = $index
                    type = $type
                    page_number = $null
                    width_pt = [double]$shape.Width
                    height_pt = [double]$shape.Height
                    metadata_errors = @()
                }
                try {
                    $shapeRange = $shape.Range
                    $pageRange = $shapeRange.Duplicate
                    $position = [int]$shapeRange.Start
                    $pageRange.SetRange($position, $position)
                    $shapeMetric.page_number = [int]$pageRange.Information(3)
                } catch { $shapeMetric.metadata_errors = @('Read page number: ' + $_.Exception.Message) }
                $inlineShapeMetrics.Add($shapeMetric)
                if ($type -in @(1, 2)) {
                    $progId = $null
                    try { $ole = $shape.OLEFormat; $progId = [string]$ole.ProgID } catch { }
                    $oleObjects.Add([ordered]@{ placement = 'inline'; index = $index; prog_id = $progId })
                }
            } finally {
                Release-ComReference $pageRange
                Release-ComReference $shapeRange
                Release-ComReference $ole
                Release-ComReference $shape
            }
        }
        for ($index = 1; $index -le $shapes.Count; $index++) {
            $shape = $null
            $ole = $null
            try {
                $shape = $shapes.Item($index)
                $type = [int]$shape.Type
                $key = [string]$type
                if (-not $floatingTypes.Contains($key)) { $floatingTypes[$key] = 0 }
                $floatingTypes[$key]++
                if ($type -in @(11, 13)) { $floatingPictures++ }
                if ($type -in @(7, 10)) {
                    $progId = $null
                    try { $ole = $shape.OLEFormat; $progId = [string]$ole.ProgID } catch { }
                    $oleObjects.Add([ordered]@{ placement = 'floating'; index = $index; prog_id = $progId })
                }
            } finally { Release-ComReference $ole; Release-ComReference $shape }
        }
        return [ordered]@{
            opened_successfully = $true
            opened_read_only = [bool]$Document.ReadOnly
            save_format = [int]$Document.SaveFormat
            compatibility_mode = [int]$Document.CompatibilityMode
            page_count = [int]$Document.ComputeStatistics(2)
            section_count = [int]$sections.Count
            table_count = [int]$tables.Count
            tables = $tableMetrics.ToArray()
            omath_count = [int]$maths.Count
            inline_shape_count = [int]$inlineShapes.Count
            inline_shapes = $inlineShapeMetrics.ToArray()
            floating_shape_count = [int]$shapes.Count
            inline_picture_count = $inlinePictures
            floating_picture_count = $floatingPictures
            inline_shape_types = $inlineTypes
            floating_shape_types = $floatingTypes
            ole_objects = $oleObjects.ToArray()
            equation_ole_object_count = @($oleObjects | Where-Object { $_.prog_id -match '(?i)(Equation|MathType|DSMT)' }).Count
            object_count_scope = 'Word Document collections; grouped shapes and header/footer objects may not be included'
            page_number_scope = 'Range.Information(3): physical page number from the document start, ignoring section numbering restarts'
        }
    } finally {
        Release-ComReference $maths
        Release-ComReference $shapes
        Release-ComReference $inlineShapes
        Release-ComReference $tables
        Release-ComReference $sections
    }
}

function Assert-BinaryDoc([string]$Path) {
    $stream = [IO.File]::OpenRead($Path)
    try {
        $signature = [byte[]]::new(8)
        if ($stream.Read($signature, 0, 8) -ne 8 -or [BitConverter]::ToString($signature) -ne 'D0-CF-11-E0-A1-B1-1A-E1') {
            throw 'The legacy DOC does not have the expected compound-file signature.'
        }
    } finally { $stream.Dispose() }
}

function Publish-StagedFile([string]$StagedPath, [string]$Destination) {
    if ([IO.File]::Exists($Destination)) {
        if (-not $Force) { throw "Output already exists; use -Force to replace it: $Destination" }
        # Replace the destination directory entry rather than truncating a possibly linked file.
        # Use a real backup pathname: some PowerShell/.NET binders turn $null into an empty string.
        $backupPath = Join-Path ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Destination))) ('.word-publish-backup-' + [Guid]::NewGuid().ToString('N'))
        [IO.File]::Replace($StagedPath, $Destination, $backupPath)
        # Delete only this operation's known backup, and only after Replace completed successfully.
        if ([IO.File]::Exists($backupPath)) { [IO.File]::Delete($backupPath) }
    } else { [IO.File]::Move($StagedPath, $Destination) }
}

$sourceFile = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
if ([IO.Path]::GetExtension($sourceFile) -ine '.docx') { throw 'SourcePath must be the working DOCX.' }
if ([string]::IsNullOrWhiteSpace($OutputStem) -or $OutputStem -in @('.', '..') -or
    $OutputStem.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
    $OutputStem -match '\.(docx?|json)$' -or $OutputStem -ne $OutputStem.TrimEnd(' ', '.')) {
    throw 'OutputStem must be a filename stem without directory separators, extension, or trailing spaces/dots.'
}
if ($OutputStem -match '^(?i)(template_source|template_reference)$') { throw 'Template source/reference filenames are protected.' }
$outputPath = Get-NormalizedPath $OutputDirectory
$docxPath = Join-Path $outputPath ($OutputStem + '.docx')
$docPath = Join-Path $outputPath ($OutputStem + '.doc')
$reportPath = Join-Path $outputPath ($OutputStem + '_save_check.json')
$protectedFiles = [Collections.Generic.List[object]]::new()
$protectedPaths = @($sourceFile)
foreach ($protectedPath in ($protectedPaths | Select-Object -Unique)) {
    foreach ($target in @($docxPath, $docPath, $reportPath)) {
        if (Test-SamePath $protectedPath $target) { throw "Output would overwrite a protected source: $protectedPath" }
    }
    $protectedFiles.Add([ordered]@{
        path = $protectedPath
        sha256_before = (Get-FileHash -LiteralPath $protectedPath -Algorithm SHA256).Hash
        sha256_after = $null
        unchanged = $null
    })
}
foreach ($target in @($docxPath, $docPath, $reportPath)) {
    if ((Test-Path -LiteralPath $target) -and -not $Force) { throw "Output already exists; use -Force to replace it: $target" }
}
[IO.Directory]::CreateDirectory($outputPath) | Out-Null
$stagingDirectory = Join-Path $outputPath ('.word-finalize-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($stagingDirectory) | Out-Null
$stagedDocx = Join-Path $stagingDirectory ($OutputStem + '.docx')
$stagedDoc = Join-Path $stagingDirectory ($OutputStem + '.doc')
$word = $null
$documents = $null
$document = $null
$ownedInstance = $false
$settingsChanged = $false
$oldAlerts = $null
$oldSecurity = $null
$failure = $null
$warnings = [Collections.Generic.List[string]]::new()
$cleanupErrors = [Collections.Generic.List[string]]::new()
$report = [ordered]@{
    started_utc = [DateTime]::UtcNow.ToString('o')
    completed_utc = $null
    status = 'running'
    source = $sourceFile
    outputs = [ordered]@{ docx = $docxPath; doc = $docPath; save_check = $reportPath }
    engine = "$hostLabel native COM ($ComProgId) / SaveAs2"
    word_version = $null
    initial_document_count = $null
    source_metrics = $null
    source_omml = $null
    field_updates = @()
    docx_check = $null
    doc_check = $null
    docx_omml = $null
    formula_compatibility = $null
    optional_document_compatibility_check_disabled = $null
    output_hashes = $null
    protected_files = $protectedFiles
    warnings = @()
    cleanup_errors = @()
    cleanup_succeeded = $null
    error = $null
}

try {
    $report.source_omml = Get-OmmlAudit $sourceFile
    $word = New-Object -ComObject $ComProgId
    $documents = $word.Documents
    $report.initial_document_count = [int]$documents.Count
    if ($documents.Count -ne 0) {
        throw "New $ComProgId contains existing documents (COM attached to an instance in use). Close that window and rerun; refusing to change settings, open files, or quit it."
    }
    $ownedInstance = $true
    $report.word_version = "$($word.Name) $($word.Version) build $(try { $word.Build } catch { '?' })"
    $oldAlerts = $word.DisplayAlerts
    try { $oldSecurity = $word.AutomationSecurity } catch { $oldSecurity = $null }   # WPS may not expose it
    $settingsChanged = $true
    $word.DisplayAlerts = 0
    if ($null -ne $oldSecurity) { $word.AutomationSecurity = 3 }
    $word.Visible = $false

    # ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False.
    $document = $documents.Open($sourceFile, $false, $true, $false)
    $report.source_metrics = Get-DocumentMetrics $document
    $report.field_updates = @(Update-InternalFields $document)
    $document.Repaginate()
    # Second pass resolves page-dependent fields after the first pagination pass.
    $report.field_updates += @(Update-InternalFields $document)
    $document.Repaginate()
    $report.optional_document_compatibility_check_disabled = Disable-OptionalDocumentCompatibilityCheck $document $warnings
    # Format 16 = DOCX, 0 = Word 97-2003 DOC. Word: the fifth SaveAs2 argument disables recent-file
    # registration; fall back to the two-argument form for hosts that reject it.
    foreach ($pair in @(@($stagedDocx, 16), @($stagedDoc, 0))) {
        try { $document.SaveAs2($pair[0], $pair[1], $false, [Type]::Missing, $false) }
        catch { $warnings.Add("SaveAs2 with 5 arguments failed on $ComProgId, used 2-argument form: $($_.Exception.Message)"); $document.SaveAs2($pair[0], $pair[1]) }
    }
    $document.Close(0)
    Release-ComReference $document
    $document = $null

    $document = $documents.Open($stagedDocx, $false, $true, $false)
    $report.docx_check = Get-DocumentMetrics $document
    $document.Close(0)
    Release-ComReference $document
    $document = $null
    $report.docx_omml = Get-OmmlAudit $stagedDocx
    # Word may report the reopened DOCX as XML document (12), even though SaveAs2 used default DOCX (16).
    if ($report.docx_check.save_format -notin @(12, 16)) { throw 'Reopened DOCX does not report a supported DOCX SaveFormat (12 or 16).' }
    if ($report.docx_omml.total_omath_count -ne $report.source_omml.total_omath_count) {
        throw 'Native OMML equation count changed during DOCX finalization. Final files were not published.'
    }

    Assert-BinaryDoc $stagedDoc
    $document = $documents.Open($stagedDoc, $false, $true, $false)
    $report.doc_check = Get-DocumentMetrics $document
    $document.Close(0)
    Release-ComReference $document
    $document = $null
    if ($report.doc_check.save_format -ne 0) { throw 'Reopened DOC does not report SaveFormat 0.' }

    $omathDelta = $report.doc_check.omath_count - $report.docx_check.omath_count
    $inlinePictureDelta = $report.doc_check.inline_picture_count - $report.docx_check.inline_picture_count
    $floatingPictureDelta = $report.doc_check.floating_picture_count - $report.docx_check.floating_picture_count
    $legacyAssessment = if ($report.source_omml.total_omath_count -eq 0) { 'no_source_native_omml' }
        elseif ($report.doc_check.omath_count -eq 0) { 'no_native_omath_objects_exposed_in_legacy_doc' }
        elseif ($omathDelta -lt 0) { 'native_omath_objects_reduced_in_legacy_doc' }
        else { 'word_reopens_legacy_doc_with_omath_objects; binary_storage_editability_not_proven' }
    $report.formula_compatibility = [ordered]@{
        docx_native_omml_preserved_by_count = $true
        docx_native_omml_status = if ($report.source_omml.total_omath_count -eq 0) { 'not_present_in_source' } else { 'preserved_by_count' }
        source_omml_count = $report.source_omml.total_omath_count
        docx_omml_count = $report.docx_omml.total_omath_count
        docx_main_document_omml_count = $report.docx_omml.main_document_omath_count
        docx_word_omath_count = $report.docx_check.omath_count
        doc_word_omath_count = $report.doc_check.omath_count
        doc_omath_count_delta = $omathDelta
        doc_inline_picture_count_delta = $inlinePictureDelta
        doc_floating_picture_count_delta = $floatingPictureDelta
        doc_equation_ole_object_count = $report.doc_check.equation_ole_object_count
        doc_assessment = $legacyAssessment
        legacy_doc_limit = 'DOC is binary, not an OMML package. Word may create picture previews/fallbacks or lose native equation editability. Image-count changes alone do not identify individual converted formulas.'
        editability_limit = 'Only native Word OMath objects are counted as native editable math. Pictures are not editable equations. An OLE ProgID is not proof that an equation editor is installed or usable.'
        mathtype_objects_created = $false
        equation_editor_installed_by_script = $false
    }
    if ($report.source_omml.total_omath_count -gt 0) {
        $warnings.Add('Legacy DOC equation editability is not guaranteed; retain DOCX as the native OMML master.')
    }
    foreach ($name in @('section_count', 'table_count', 'omath_count', 'inline_shape_count', 'floating_shape_count')) {
        if ($report.docx_check[$name] -ne $report.source_metrics[$name]) {
            $warnings.Add("DOCX $name changed: source=$($report.source_metrics[$name]); final=$($report.docx_check[$name]).")
        }
    }
    foreach ($name in @('page_count', 'section_count', 'table_count', 'inline_shape_count', 'floating_shape_count')) {
        if ($report.doc_check[$name] -ne $report.docx_check[$name]) {
            $warnings.Add("Legacy DOC $name differs: DOCX=$($report.docx_check[$name]); DOC=$($report.doc_check[$name]).")
        }
    }
    if (@($report.field_updates | Where-Object { $_.status -in @('update_error', 'update_returned_false') }).Count -gt 0) {
        $warnings.Add('Some internal fields did not update successfully; inspect field_updates in this report.')
    }
    foreach ($item in $protectedFiles) {
        if ((Get-FileHash -LiteralPath $item.path -Algorithm SHA256).Hash -ne $item.sha256_before) {
            throw "Protected source changed before publishing outputs: $($item.path)"
        }
    }
    Publish-StagedFile $stagedDocx $docxPath
    Publish-StagedFile $stagedDoc $docPath
    $report.output_hashes = [ordered]@{
        docx_sha256 = (Get-FileHash -LiteralPath $docxPath -Algorithm SHA256).Hash
        doc_sha256 = (Get-FileHash -LiteralPath $docPath -Algorithm SHA256).Hash
        docx_bytes = (Get-Item -LiteralPath $docxPath).Length
        doc_bytes = (Get-Item -LiteralPath $docPath).Length
    }
    $report.status = 'passed'
} catch {
    $failure = $_
    $report.status = 'failed'
    $report.error = [ordered]@{ message = $_.Exception.Message; type = $_.Exception.GetType().FullName; detail = $_.ScriptStackTrace }
} finally {
    if ($null -ne $document) {
        try { $document.Close(0) } catch { $cleanupErrors.Add('Close owned document: ' + $_.Exception.Message) }
        try { Release-ComReference $document } catch { $cleanupErrors.Add('Release owned document: ' + $_.Exception.Message) }
        $document = $null
    }
    if ($null -ne $word) {
        if ($ownedInstance -and $settingsChanged) {
            try { $word.DisplayAlerts = $oldAlerts } catch { $cleanupErrors.Add('Restore DisplayAlerts: ' + $_.Exception.Message) }
            if ($null -ne $oldSecurity) { try { $word.AutomationSecurity = $oldSecurity } catch { $cleanupErrors.Add('Restore AutomationSecurity: ' + $_.Exception.Message) } }
        }
        if ($ownedInstance) {
            try {
                if ($documents.Count -eq 0) { $word.Quit(0) }
                else { $cleanupErrors.Add('Owned Word instance now has open documents; did not quit or close any additional documents.') }
            } catch { $cleanupErrors.Add('Quit empty owned Word instance: ' + $_.Exception.Message) }
        }
    }
    try { Release-ComReference $documents } catch { $cleanupErrors.Add('Release Documents: ' + $_.Exception.Message) }
    try { Release-ComReference $word } catch { $cleanupErrors.Add('Release Word.Application: ' + $_.Exception.Message) }
    $documents = $null
    $word = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($ownedInstance) {
        # WPS keeps helper processes alive after Quit(); remove only automation instances this run started
        Start-Sleep -Seconds 2
        Get-CimInstance Win32_Process -Filter "Name='wps.exe' OR Name='WINWORD.EXE'" |
            Where-Object { $automationBefore -notcontains $_.ProcessId -and $_.CommandLine -match '/Automation|-Embedding' } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    }
    foreach ($item in $protectedFiles) {
        try {
            $item.sha256_after = (Get-FileHash -LiteralPath $item.path -Algorithm SHA256).Hash
            $item.unchanged = $item.sha256_after -eq $item.sha256_before
            if (-not $item.unchanged) {
                $report.status = 'failed'
                $cleanupErrors.Add('Protected source hash changed: ' + $item.path)
            }
        } catch {
            $report.status = 'failed'
            $cleanupErrors.Add('Verify protected source: ' + $_.Exception.Message)
        }
    }
    # Remove only the two known staged outputs; never recursively delete a directory.
    try {
        foreach ($path in @($stagedDocx, $stagedDoc)) {
            if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) }
        }
        if ([IO.Directory]::Exists($stagingDirectory)) { [IO.Directory]::Delete($stagingDirectory, $false) }
    } catch { $cleanupErrors.Add('Remove staging files/directory: ' + $_.Exception.Message) }
    $report.cleanup_succeeded = $cleanupErrors.Count -eq 0
    if (-not $report.cleanup_succeeded) { $warnings.Add('One or more cleanup/restore operations need review; see cleanup_errors.') }
    $report.completed_utc = [DateTime]::UtcNow.ToString('o')
    $report.warnings = $warnings.ToArray()
    $report.cleanup_errors = $cleanupErrors.ToArray()
}

# Emit a failure report too, including source-hash and cleanup evidence.
$reportTemporary = Join-Path $outputPath ('.word-save-check-' + [Guid]::NewGuid().ToString('N') + '.json')
try {
    [IO.File]::WriteAllText($reportTemporary, ($report | ConvertTo-Json -Depth 14), [Text.UTF8Encoding]::new($false))
    Publish-StagedFile $reportTemporary $reportPath
} catch {
    # Preserve diagnostic evidence if the final report cannot be published.
    Write-Warning ('Save-check report publication failed. Diagnostic temporary path: ' + $reportTemporary)
    throw
}
$report | ConvertTo-Json -Depth 14
if ($null -ne $failure) { throw $failure }
if ($report.status -ne 'passed') { throw "Finalization failed; see $reportPath" }
