# Modify fields sent to it with proper word wrapping.
function wordwrap ($field, $maximumlinelength) {if ($null -eq $field) {return $null}
$breakchars = ',.;?!\/ '; $wrapped = @()
if (-not $maximumlinelength) {[int]$maximumlinelength = (100, $Host.UI.RawUI.WindowSize.Width | Measure-Object -Maximum).Maximum}
if ($maximumlinelength -lt 60) {[int]$maximumlinelength = 60}
if ($maximumlinelength -gt $Host.UI.RawUI.BufferSize.Width) {[int]$maximumlinelength = $Host.UI.RawUI.BufferSize.Width}
foreach ($line in $field -split "`n", [System.StringSplitOptions]::None) {if ($line -eq "") {$wrapped += ""; continue}
$remaining = $line
while ($remaining.Length -gt $maximumlinelength) {$segment = $remaining.Substring(0, $maximumlinelength); $breakIndex = -1
foreach ($char in $breakchars.ToCharArray()) {$index = $segment.LastIndexOf($char)
if ($index -gt $breakIndex) {$breakIndex = $index}}
if ($breakIndex -lt 0) {$breakIndex = $maximumlinelength - 1}
$chunk = $segment.Substring(0, $breakIndex + 1); $wrapped += $chunk; $remaining = $remaining.Substring($breakIndex + 1)}
if ($remaining.Length -gt 0 -or $line -eq "") {$wrapped += $remaining}}
return ($wrapped -join "`n")}

# Display a horizontal line.
function line ($colour, $length, [switch]$pre, [switch]$post, [switch]$double) {if (-not $length) {[int]$length = (100, $Host.UI.RawUI.WindowSize.Width | Measure-Object -Maximum).Maximum}
if ($length) {if ($length -lt 60) {[int]$length = 60}
if ($length -gt $Host.UI.RawUI.BufferSize.Width) {[int]$length = $Host.UI.RawUI.BufferSize.Width}}
if ($pre) {Write-Host ""}
$character = if ($double) {"="} else {"-"}
Write-Host -f $colour ($character * $length)
if ($post) {Write-Host ""}}

function help {# Inline help.
# Select content.
$scripthelp = Get-Content -Raw -Path $PSCommandPath; $sections = [regex]::Matches($scripthelp, "(?im)^## (.+?)(?=\r?\n)"); $selection = $null; $lines = @(); $wrappedLines = @(); $position = 0; $pageSize = 30; $inputBuffer = ""

function scripthelp ($section) {$pattern = "(?ims)^## ($([regex]::Escape($section)).*?)(?=^##|\z)"; $match = [regex]::Match($scripthelp, $pattern); $lines = $match.Groups[1].Value.TrimEnd() -split "`r?`n", 2; if ($lines.Count -gt 1) {$wrappedLines = (wordwrap $lines[1] 100) -split "`n", [System.StringSplitOptions]::None}
else {$wrappedLines = @()}
$position = 0}

# Display Table of Contents.
while ($true) {cls; Write-Host -f cyan "$(Get-ChildItem (Split-Path $PSCommandPath) | Where-Object { $_.FullName -ieq $PSCommandPath } | Select-Object -ExpandProperty BaseName) Help Sections:`n"

if ($sections.Count -gt 7) {$half = [Math]::Ceiling($sections.Count / 2)
for ($i = 0; $i -lt $half; $i++) {$leftIndex = $i; $rightIndex = $i + $half; $leftNumber  = "{0,2}." -f ($leftIndex + 1); $leftLabel   = " $($sections[$leftIndex].Groups[1].Value)"; $leftOutput  = [string]::Empty

if ($rightIndex -lt $sections.Count) {$rightNumber = "{0,2}." -f ($rightIndex + 1); $rightLabel  = " $($sections[$rightIndex].Groups[1].Value)"; Write-Host -f cyan $leftNumber -n; Write-Host -f white $leftLabel -n; $pad = 40 - ($leftNumber.Length + $leftLabel.Length)
if ($pad -gt 0) {Write-Host (" " * $pad) -n}; Write-Host -f cyan $rightNumber -n; Write-Host -f white $rightLabel}
else {Write-Host -f cyan $leftNumber -n; Write-Host -f white $leftLabel}}}

else {for ($i = 0; $i -lt $sections.Count; $i++) {Write-Host -f cyan ("{0,2}. " -f ($i + 1)) -n; Write-Host -f white "$($sections[$i].Groups[1].Value)"}}

# Display Header.
line yellow 100
if ($lines.Count -gt 0) {Write-Host  -f yellow $lines[0]}
else {Write-Host "Choose a section to view." -f darkgray}
line yellow 100

# Display content.
$end = [Math]::Min($position + $pageSize, $wrappedLines.Count)
for ($i = $position; $i -lt $end; $i++) {Write-Host -f white $wrappedLines[$i]}

# Pad display section with blank lines.
for ($j = 0; $j -lt ($pageSize - ($end - $position)); $j++) {Write-Host ""}

# Display menu options.
line yellow 100; Write-Host -f white "[↑/↓]  [PgUp/PgDn]  [Home/End]  |  [#] Select section  |  [Q] Quit  " -n; if ($inputBuffer.length -gt 0) {Write-Host -f cyan "section: $inputBuffer" -n}; $key = [System.Console]::ReadKey($true)

# Define interaction.
switch ($key.Key) {'UpArrow' {if ($position -gt 0) { $position-- }; $inputBuffer = ""}
'DownArrow' {if ($position -lt ($wrappedLines.Count - $pageSize)) { $position++ }; $inputBuffer = ""}
'PageUp' {$position -= 30; if ($position -lt 0) {$position = 0}; $inputBuffer = ""}
'PageDown' {$position += 30; $maxStart = [Math]::Max(0, $wrappedLines.Count - $pageSize); if ($position -gt $maxStart) {$position = $maxStart}; $inputBuffer = ""}
'Home' {$position = 0; $inputBuffer = ""}
'End' {$maxStart = [Math]::Max(0, $wrappedLines.Count - $pageSize); $position = $maxStart; $inputBuffer = ""}

'Enter' {if ($inputBuffer -eq "") {"`n"; return}
elseif ($inputBuffer -match '^\d+$') {$index = [int]$inputBuffer
if ($index -ge 1 -and $index -le $sections.Count) {$selection = $index; $pattern = "(?ims)^## ($([regex]::Escape($sections[$selection-1].Groups[1].Value)).*?)(?=^##|\z)"; $match = [regex]::Match($scripthelp, $pattern); $block = $match.Groups[1].Value.TrimEnd(); $lines = $block -split "`r?`n", 2
if ($lines.Count -gt 1) {$wrappedLines = (wordwrap $lines[1] 100) -split "`n", [System.StringSplitOptions]::None}
else {$wrappedLines = @()}
$position = 0}}
$inputBuffer = ""}

default {$char = $key.KeyChar
if ($char -match '^[Qq]$') {"`n"; return}
elseif ($char -match '^\d$') {$inputBuffer += $char}
else {$inputBuffer = ""}}}}}

function recycle {# A public module to mimic most features of the Remove-Item cmdlet, but with safe, Recycle Bin support.

# Act like a cmdlet, with -whatif and -confirm support, and user confirmation dependent on user settings.
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]

# Accepts -fullname, -recurse, -force
param([Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)] [Alias('FullName')] [Object[]]$Path, [switch]$Recurse, [switch]$Force, [switch]$Help)

# Add the Visual Basic library.
begin {Add-Type -AssemblyName Microsoft.VisualBasic}

# Resolve every path.
process {foreach ($item in $Path) {if ($item -is [string]) {$resolvedPaths = Get-Item -LiteralPath $item -ErrorAction SilentlyContinue

if ((-not $Path) -and (-not $help)) {Write-Host -f Cyan "`nUsage: Recycle [Path or Piped Object] <-Recurse> <-Force> <-WhatIf> <-Confirm> <-Help>`n"; return}

# External call to help.
if ($help) {help; return}

# Error-catching for unresolved paths.
if (-not $resolvedPaths) {$resolvedPaths = Get-Item -Path $item -ErrorAction SilentlyContinue}
if (-not $resolvedPaths) {Write-Warning "Path not found: $item"; continue}}

# Accepts piped objects, such as those sent from Get-Item or Get-ChildItem, etc.
elseif ($item -is [System.IO.FileSystemInfo]) {$resolvedPaths = $item}

# Error-catching for unsupported types like hash tables and Registry keys.
else {Write-Warning "Unsupported input type: $($item.GetType().Name)"; continue}

# Extracts full names for each legitimate object identified.
foreach ($resolved in @($resolvedPaths)) {$targetPath = $resolved.FullName; $isDirectory = $resolved.PSIsContainer

# Error-checking if the object is a directory, but -recurse was not specified.
if ($isDirectory -and -not $Recurse) {Write-Warning "Directory found but -Recurse not specified: $targetPath"; continue}

# Support for -whatif through PSCmdlet and then directory delete when -recurse is specified and supports -force.
if ($PSCmdlet.ShouldProcess($targetPath, 'Send to Recycle Bin')) {try {if ($isDirectory) {$cancelOption = if ($Force) {[Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException}
else {[Microsoft.VisualBasic.FileIO.UICancelOption]::DoNothing}
[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($targetPath, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin, $cancelOption)}

# File deletion, which doesn't require -recurse, but requires -force for read-only files.
elseif (-not $isDirectory) {$fileInfo = Get-Item -LiteralPath $targetPath
if (-not $Force) {if ($fileInfo.Attributes -band [System.IO.FileAttributes]::ReadOnly) {Write-Warning "File '$targetPath' is read-only. Use -Force to override."; continue}}
else {if ($fileInfo.Attributes -band [System.IO.FileAttributes]::ReadOnly) {$fileInfo.Attributes = $fileInfo.Attributes -bxor [System.IO.FileAttributes]::ReadOnly}}
[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($targetPath, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)}
Write-Verbose "Sent to Recycle Bin: $targetPath"}

# Error-catching
catch {Write-Warning "Failed to delete '$targetPath': $_"}}}}}}

Export-ModuleMember -Function recycle

# Helptext.
<#
## Recycle
This module adds Recycle Bin capabilities to PowerShell.

By default, PowerShell's Remove-Item function deletes files and objects completely, without a way to recover them. This module attempts to correct that by allowing files and directories to be deleted to the Windows Recycle Bin, instead.

Note: This module does import the Visual Basic library in order to accomplish this.
## Usage
Usage: Recycle [Path or Piped Object] <-Recurse> <-Force> <-WhatIf> <-Confirm> <-Help>

• Accepts piped objects, such as those sent from Get-Item or Get-ChildItem, for example: 
	Get-ChildItem C:\Temp | Recycle 
	Get-Item C:\Temp\file1.txt, C:\TMP | Recycle
	Get-Item C:\Temp\file1.txt, C:\TMP, C:\SomeDirectoryTree | Recycle -Recurse

• Since piped objects are allowed, error-checking is included to make sure the objects passed to this function are either files or directories.

• As a safety measure, the -Recurse option must be used if the object passed to it is a directory.

• Supports the -WhatIf switch to demonstrate what would happen for each action.

• Suppports the -Confirm switch as extra protection for activities deemed to have a higher risk impact.

• The -Force switch will exit the function with an exception if an interrupt is encountered. Otherwise, interrupts fail silently.

• Also, when deleting Read-Only files the -Force switch is required.
## License
MIT License

Copyright (c) 2025 Craig Plath

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
##>
