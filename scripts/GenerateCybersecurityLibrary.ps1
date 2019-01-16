Set-StrictMode -Version Latest

Import-Module -Name .\CybersecurityLibrary.psm1

$libraryPath = "$env:userprofile\Desktop\cybersecuritylibrary {0:yyyyMMdd}" -f [DateTime]::Now
$libraryJsonPath = "$libraryPath\library.json"
$librarySitePath = "$env:userprofile\Documents\GitHub\nsacyber.github.io"

$newUnprotectedPubs = Get-UnprotectedCybersecurityLibrary -Path $libraryPath
$newProtectedPubs = Get-ProtectedCybersecurityLibrary -Path $libraryPath -MetadataOnly # metadata only is used because authentication is required to download protected documents
$oldPubs = Get-ArchivedCybersecurityLibrary -Path $libraryPath

$allPubs = New-Object System.Collections.Generic.List[pscustomobject]
$allPubs.AddRange($newUnprotectedPubs)
$allPubs.AddRange($newProtectedPubs)
$allPubs.AddRange($oldPubs)

$json = $allPubs | Sort-Object -Property {$_.Converted.Date} -Descending | ConvertTo-Json -Depth 4
$json | Out-File -FilePath $libraryJsonPath -Encoding Unicode -NoNewline -Force


$markdown = Get-CybersecurityLibraryMarkdown -Path $libraryJsonPath

$markdown | Out-File -FilePath "$librarySitePath\publications.md" -Encoding UTF8 -NoNewline -Force 

Compress-Archive -Path "$libraryPath\*" -DestinationPath "$libraryPath.zip" -Verbose:$false