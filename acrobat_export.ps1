<#
.SYNOPSIS
    Silently export PDF to Word/Excel/PPT via Adobe Acrobat Pro IAC (no mouse, no dialogs).

.DESCRIPTION
    Uses Acrobat's JavaScript bridge (PDDoc.GetJSObject -> saveAs), the same engine
    behind the GUI "Export To" menu. Requires Acrobat Pro (not Reader).

.PARAMETER PdfPath
    Source PDF path (required). Accepts pipeline input.

.PARAMETER Format
    Target format: docx (default) / doc / xlsx / pptx / rtf / txt / html.

.PARAMETER OutPath
    Output file path. Defaults to same folder as source, same base name.

.EXAMPLE
    .\acrobat_export.ps1 -PdfPath "E:\proj\data.pdf"
    .\acrobat_export.ps1 -PdfPath "E:\proj\data.pdf" -Format xlsx -OutPath "E:\proj\data.xlsx"
    Get-ChildItem D:\pdfs\*.pdf | ForEach-Object { .\acrobat_export.ps1 $_.FullName }
#>
param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('FullName')]
    [string]$PdfPath,
    [ValidateSet('docx', 'doc', 'xlsx', 'pptx', 'rtf', 'txt', 'html')]
    [string]$Format = 'docx',
    [string]$OutPath = ''
)

begin {
    $formatIds = @{
        docx = 'com.adobe.acrobat.docx'
        doc  = 'com.adobe.acrobat.doc'
        xlsx = 'com.adobe.acrobat.xlsx'
        pptx = 'com.adobe.acrobat.pptx'
        rtf  = 'com.adobe.acrobat.rtf'
        txt  = 'com.adobe.acrobat.txt'
        html = 'com.adobe.acrobat.html'
    }
    $flags = [Reflection.BindingFlags]::InvokeMethod -bor
             [Reflection.BindingFlags]::Public -bor
             [Reflection.BindingFlags]::Instance
    $comType = [System.__ComObject]

    function Convert-One {
        param([string]$Pdf, [string]$Fmt, [string]$Out)

        if (-not (Test-Path -LiteralPath $Pdf)) {
            Write-Output "FAIL|$Pdf|Source PDF not found"
            return
        }
        if ([string]::IsNullOrEmpty($Out)) {
            $dir = Split-Path -Parent $Pdf
            $base = [IO.Path]::GetFileNameWithoutExtension($Pdf)
            $Out = Join-Path $dir "$base.$Fmt"
        }
        $dirOut = Split-Path -Parent $Out
        if ($dirOut -and -not (Test-Path -LiteralPath $dirOut)) {
            New-Item -ItemType Directory -Path $dirOut -Force | Out-Null
        }
        if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }

        $avDoc = $null; $pdDoc = $null; $js = $null
        try {
            $avDoc = New-Object -ComObject AcroExch.AVDoc
            if (-not $avDoc.Open($Pdf, '')) { throw 'Acrobat could not open the PDF' }
            $pdDoc = $avDoc.GetPDDoc()
            $js = $pdDoc.GetJSObject()
            if ($null -eq $js) { throw 'JSObject unavailable (Acrobat Pro required, not Reader)' }

            # Core: invoke saveAs on the JS bridge object via reflection.
            $null = $comType.InvokeMember('saveAs', $flags, $null, $js, @($Out, $formatIds[$Fmt]))

            if (Test-Path -LiteralPath $Out) {
                $size = (Get-Item -LiteralPath $Out).Length
                Write-Output "OK|$Out|$size"
            } else {
                Write-Output "FAIL|$Out|No output produced (possibly blocked by Acrobat security policy)"
            }
        } catch {
            Write-Output "FAIL|$Pdf|$($_.Exception.Message)"
        } finally {
            # Only close documents WE opened; never quit Acrobat (user may have other docs).
            if ($null -ne $js)   { try { $null = $comType.InvokeMember('closeDoc', $flags, $null, $js, @($true)) } catch {} }
            if ($null -ne $pdDoc){ try { $pdDoc.Close() | Out-Null } catch {} }
            if ($null -ne $avDoc){ try { $avDoc.Close(1) | Out-Null } catch {} }
        }
    }
}
process {
    Convert-One -Pdf $PdfPath -Fmt $Format -Out $OutPath
}
