Set-StrictMode -Version Latest

Function Get-Intersection() {
    [CmdletBinding()]
    [OutputType([string[]])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='An array of string objects used as a reference for comparison')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ReferenceObject,

        [Parameter(Mandatory=$true, HelpMessage='An array of string objects compared to the reference objects')]
        [ValidateNotNullOrEmpty()]
        [string[]]$DifferenceObject
    )

    $result = [string[]]@(Compare-Object $ReferenceObject $DifferenceObject -PassThru -IncludeEqual -ExcludeDifferent)

    if($null -eq $result) {
        $result = [string[]]@()
    }

    return ,$result
}

Function Invoke-PageDownload() {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='The URL of the page to download')]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $uri = $Url

    $params = @{
        Uri = $Url;
        Method = 'Get';
        UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36';
    }

    $proxyUri = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy($uri)

    $ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

    if(([string]$proxyUri) -ne $uri) {
        $response = Invoke-WebRequest @params -Proxy $proxyUri -ProxyUseDefaultCredentials -UseBasicParsing
    } else {
        $response = Invoke-WebRequest @params -UseBasicParsing
    }

    $statusCode = $response.StatusCode

    if ($statusCode -eq 200) {
       $page = $response.Content

        return ($page | Out-String)
    } else {
        throw 'Request failed with status code $statusCode'
    }
}

Function Invoke-FileDownload() {
    [CmdletBinding()]
    [OutputType([void])]
    Param (
        [Parameter(Mandatory=$true, HelpMessage='URL of a file to download')]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url,

        [Parameter(Mandatory=$true, HelpMessage='The path to download the file to')]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $uri = $Url

    $params = @{
        Uri = $uri;
        Method = 'GET';
        ContentType = 'text/plain';
        UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.110 Safari/537.36';
    }

    $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

    $ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

    $proxyUri = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy($uri)

    if(([string]$proxyUri) -ne $uri) {
        $response = Invoke-WebRequest @params -Proxy $proxyUri -ProxyUseDefaultCredentials -UseBasicParsing
    } else {
        $response = Invoke-WebRequest @params -UseBasicParsing
    }

    $statusCode = $response.StatusCode

    if ($statusCode -eq 200) {
        $bytes = $response.Content
    } else {
        throw "Request failed with status code $statusCode"
    }

    Set-Content -Path $Path -Value $bytes -Encoding Byte -Force

    if(-not(Test-Path -Path $Path)) {
        throw "failed to download to $Path"
    }
}

Function Convert-Title() {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='The title')]
        [ValidateNotNullOrEmpty()]
        [string]$Title
    )

    $convertedTitle = $Title.Replace('(IAA)','').Replace('IAA','').Replace('(SDN)','').Replace('(CSfC)','').Replace('(SNMP)','').Replace('&amp;','and').Replace("$([char]0x00AE)1",'').Replace("$([char]0x00AE)2",'').Replace("$([char]0x2122)",'').Replace("$([char]0x00AE)",'').Replace('(NMP-I)','').Replace(' / ','').Replace('- Unclassified','').Replace('JCMA-Findings-and-Trends','Joint COMSEC Monitoring Activity Findings and Trends').Replace('Third-Party-Services-Your-Risk-Picture-Just-Got-a-Lot-More-Complex','Third Party Services Your Risk Picture Just Got a Lot More Complex').Replace('– Version 1','').Replace('(NDI)','').Replace('IADs',"IAD's").Replace('(CNSA)','').Replace('CGS All Files Zipped','Community Gold Standard 1.1.1 files').Replace('(ICS)','').Replace('Page','').Replace('(HMP)','').Replace('Out-of the Box','Out-of-the-Box').Replace('-old','').Replace('&quot;','"').Replace('Top10','Top 10').Replace('-VAS','VAS').Replace('CNSA','Commercial National Security Algorithm').Replace('(NSS)','').Replace('(CIRA)','').Replace('(XSS)','').Replace('(CUPS)','').Replace('(CUCME)','').Replace('"','').Replace('(ECDSA)','').Replace('WIDS','Wireless Intrusion Detection System').Replace("$([char]0x2018)","'").Replace("$([char]0x2019)","'").Replace('NSS','National Security Systems').Replace('NSCAP','National Security Cyber Assistance Program').Replace('CIRA','Cyber Incident Response Assistance').Replace('VPN','Virtual Private Network').Replace(' VA ',' Vulnerability Assessment ').Replace('WLAN','Wireless Local Area Network').Replace(' Tech ',' Technology ').Replace('SCAP','Security Content Automation Protocol').Replace('PDF','Portable Document Format').Replace('DNS','Domain Name System').Replace('IBM','').Replace('NIST','National Institute of Standards and Technology').Replace('FIPS','Federal Information Processing Standard').Replace('NIAP','National Information Assurance Partnership').Replace(' AD ',' Active Directory ').Replace('Cisco ASA','Cisco Adaptive Security Appliance').Replace('HBSS','Host Based Security System').Replace('- RSA','RSA').Replace('  ',' ').Trim()

    if ($convertedTitle.EndsWith(':')) {
        $convertedTitle = $convertedTitle[0..($convertedTitle.Length-2)] -join ''
    }

    return $convertedTitle
}

Function Get-Abstract() {
    [CmdletBinding()]
    [OutputType([string])]
    Param (
        [Parameter(Mandatory=$true, HelpMessage='Content of a page')]
        [ValidateNotNullOrEmpty()]
        [string]$Content
    )

    $startToken='<span class="metadataBlockText" itemprop="description">'
    $endToken='</span>'

    $startTokenIndex=$Content.IndexOf($startToken)
    $startUrlIndex=$startTokenIndex+$startToken.Length

    $endTokenIndex=$Content.IndexOf($endToken, $startTokenIndex)
    $endUrlIndex=$endTokenIndex-1

    if ($startTokenIndex -eq -1 -or $endTokenIndex -eq -1) {
        return ''
    }

    $description = $Content.Substring(($startUrlIndex), ($endUrlIndex + 1 - $startUrlIndex))
    $description = $description.Replace('<br /><br />', "`n")
    $description = $description.Replace('<br />',' ')
    $description = $description.Replace("$([char]0x00A0)","$([char]0x0020)")
    return $description
}

Function Convert-Abstract() {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='The abstract')]
        [ValidateNotNullOrEmpty()]
        [string]$Abstract
    )

    $convertedAbstract = $Abstract.Replace('privilege2','privilege').Replace('Th is','This').Replace('beacuse','Because')

    return $convertedAbstract

}

Function Convert-Description() {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='The description')]
        [ValidateNotNullOrEmpty()]
        [string]$Description
    )
    #$([char]0x000A)
    $convertedDescription = $Description.Replace('[1]','').Replace('[2]','').Replace("$([char]0x00AE)1",'').Replace("$([char]0x00AE)2",'').Replace("$([char]0x00AE)",'').Replace('65-bit','64-bit').Replace("$([char]0x2018)","'").Replace("$([char]0x2019)","'").Replace('&amp;','and').Replace('(U) Mitigations','Mitigations').Replace('PowerShellTM','PowerShell').Replace("$([char]0x201C)",'"').Replace("$([char]0x201D)",'"').Replace('bu er over ow','buffer overflow').Replace(".`n",".$([char]0x000D)").Replace("`n",' ').Replace('.9','.').Replace('CTR-U-OO-802243-16','').Replace("$([char]0x2122)",'').Replace('64-bit6','64-bit').Replace('[1,2,3]','').Replace('NPM-1','NPM-I').Replace('CandA','C&amp;A').Replace('(1)','').Replace('(2)','').Replace('(TSG)','(TCG)').Replace('e.g.,','e.g.').Replace('Exercise(CDX)','Exercise (CDX)').Replace('  ',' ')

    return $convertedDescription
}

Function Get-Category() {
    [CmdletBinding()]
    [OutputType([string])]
    Param (
        [Parameter(Mandatory=$true, HelpMessage='URL')]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url
    )

    $segments = $Url.Segments
    $categorySegments = $segments[3..($segments.Length-2)]
    $category = ($categorySegments -join '')

    return $category.Substring(0, $category.Length-1)
}

Function Convert-Category() {
    [CmdletBinding()]
    [OutputType([string])]
    Param (
        [Parameter(Mandatory=$true, HelpMessage='URL')]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url
    )

    $segments = $Url.Segments
    $categorySegments = $segments[3..($segments.Length-2)]

    $convertedCategory = $categorySegments -join ' > '
    $convertedCategory = $convertedCategory.Replace('/','').Replace('-',' ').Trim()

    $textInfo = (New-Object System.Globalization.CultureInfo 'en-US',$false).TextInfo
    $convertedCategory = $textInfo.ToTitleCase($convertedCategory)
    $convertedCategory = $convertedCategory.Replace('Alerts','').Replace('Tech','Technical').Replace('Ias ','IA Symposium ').Replace('Ia ','IA ').Replace('Faq','FAQ').Replace('Cgs','Community Gold Standard').Trim()

    return $convertedCategory
}

Function Get-DirectLink() {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='URL')]
        [ValidateNotNullOrEmpty()]
        [string]$Content
    )

    $startToken='FilePath='
    $endToken='&WpKes='

    $startTokenIndex=$Content.IndexOf($startToken)
    $startUrlIndex=$startTokenIndex+$startToken.Length

    $endTokenIndex=$Content.IndexOf($endToken, $startTokenIndex)
    $endUrlIndex=$endTokenIndex-1

    if ($startTokenIndex -eq -1 -or $endTokenIndex -eq -1) {
        return ''
    }

    # embedded link https://apps.nsa.gov/iaarchive/customcf/openAttachment.cfm?FilePath=/iad/library/ia-advisories-alerts/assets/public/upload/Drupal-Unauthenticated-Remote-Code-Execution-Vulnerability-CVE-2018-7600.pdf&WpKes=aF6woL7fQp3dJiZH3JsnngkER6tJNBTxwuZKY3

    # direct link: https://apps.nsa.gov/iaarchive/library/ia-advisories-alerts/assets/public/upload/Drupal-Unauthenticated-Remote-Code-Execution-Vulnerability-CVE-2018-7600.pdf

    $directRelativeLink = $Content.Substring(($startUrlIndex), ($endUrlIndex + 1 - $startUrlIndex))
    $directLink = '{0}{1}' -f 'https://apps.nsa.gov/iaarchive/',$directRelativeLink.Replace('/iad/', '')

    return $directLink
}

Function Get-SafeFileName() {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='URL')]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url,

        [Parameter(Mandatory=$true, HelpMessage='The title')]
        [ValidateNotNullOrEmpty()]
        [string]$Title
    )
    $safeTitle = $Title

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $filenameChars = $Title.ToCharArray()

    $presentInvalidChars = Get-Intersection -ReferenceObject $invalidChars -DifferenceObject $filenameChars

    if($presentInvalidChars.Length -gt 0) {
        $sanitized = $Title.Split($invalidChars, [System.StringSplitOptions]::RemoveEmptyEntries);
        $safeTitle = ($sanitized -join ' ').Replace('  ',' ')
    }

    #$safeTitle = $safeTitle.Replace("$([char]0x2014)",'-') # did not work
    #$safeTitle = $safeTitle.Replace("—",'-') # this is 0x2014 but it causes a script parsing error
    #$safeTitle = $safeTitle.Replace('—','-')
    #$safeTitle = $safeTitle.Replace((0x2014 -as [char]),'-') # did not work
    #$safeTitle = $safeTitle -replace '—','-' # did not work
    #$safeTitle = $safeTitle -replace 'u2014','-' # did not work
    $safeTitle = $safeTitle -replace '[^\p{L}\p{Nd}\s]', '' # works
    $safeTitle = $safeTitle.Replace("'",'').Replace('(','').Replace(')','').Replace('  ',' ')

    $originalFile = [System.IO.FileInfo]$Url.Segments[-1]

    $safeFile = '{0}{1}' -f $safeTitle,$originalFile.Extension

    return $safeFile
}

Function Get-UniqueFilePath() {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='Path')]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory=$true, HelpMessage='File')]
        [ValidateNotNullOrEmpty()]
        [string]$File
    )

    $uniqueFilePath = '{0}\{1}' -f $Path,$File

    if (Test-Path -Path $uniqueFilePath -PathType Leaf) {
        $extension = ([System.IO.FileInfo]$File).Extension
        $name = ([System.IO.FileInfo]$File).BaseName

        $count = 0

        while(Test-Path -Path $uniqueFilePath -PathType Leaf) {
            $count++

            $uniqueFilePath = '{0}\{1}-{2}{3}' -f $Path,$name,$count,$extension
        }
    }

    return $uniqueFilePath
}

Function Compare-CybersecurityLibrary() {
    [CmdletBinding()]
    [OutputType([void])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='ReferenceObject')]
        [ValidateNotNullOrEmpty()]
        [pscustomobject]$ReferenceObject,

        [Parameter(Mandatory=$true, HelpMessage='DifferenceObject')]
        [ValidateNotNullOrEmpty()]
        [pscustomobject]$DifferenceObject
    )

    Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject -Property 'Documents'
}

Function Get-UnprotectedCybersecurityLibrary() {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='Path to save the archive to')]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory=$false, HelpMessage='Only retrieve metadata')]
        [ValidateNotNullOrEmpty()]
        [switch]$MetadataOnly
    )

    if (-not(Test-Path -Path $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory | Out-Null
    }

    $documents = New-Object System.Collections.Generic.List[pscustomobject]

    $page = Invoke-PageDownload -Url 'https://www.nsa.gov/what-we-do/cybersecurity'

    # V1: <a href="/what-we-do/cybersecurity/assets/files/professional-resources/csa-drupal-remote-code-execution-vulnerability-cve.pdf?v=1">Advisory: Drupal Unauthenticated Remote Code Execution Vulnerability (April 2018)</a>
    # V2: <a href="/Portals/70/documents/what-we-do/cybersecurity/professional-resources/csi-uefi-advantages-over-legacy-mode.pdf?v=1">Info Sheet: UEFI Advantages Over Legacy Mode (March 2018)</a>
    $startToken = '<li><a href="/Portals/70/documents/what-we-do/cybersecurity/professional-resources/'
    $endToken = '</li>'

    $startUnprotectedIndex = 0;
    $startProtectedIndex = 0;
    $startIndex = $page.IndexOf($startToken, [StringComparison]::CurrentCultureIgnoreCase)

    while($startIndex -ne -1) {
        $endIndex = $page.IndexOf($endToken, $startIndex, [StringComparison]::CurrentCultureIgnoreCase)

        if ($endIndex -ne -1) {

            $line = $page[$startIndex..($endIndex+$endToken.Length)] -join ''

            if($line -match '.*<a href=(?<RelativeLink>.*)>(?<Title>.*).*</a>.*') {
                $originalLink = 'https://www.nsa.gov{0}' -f $matches.RelativeLink.Replace('"','')
                $directLink = $originalLink.Replace('?v=1','')

                $originalTitle = $matches.Title

                if($originalTitle -match '(?<Category>.*):(?<Title>.*)\((?<Date>[0-9A-Za-z ]+)\).*') {
                    $convertedTitle = Convert-Title -Title $matches.Title.Trim()
                    $originalCategory = $matches.Category.Trim()
                    $convertedCategory = $originalCategory.Trim()
                    $originalDate = $matches.Date.Trim()
                    $convertedDate = [DateTime]::ParseExact($originalDate, 'MMMM yyyy', [System.Globalization.CultureInfo]::CurrentCulture)
                } else {
                    Write-Warning -Message ('{0} did not match typical title pattern' -f $originalTitle)
                    $convertedTitle = Convert-Title -Title $originalTitle
                    $originalCategory = ''
                    $convertedCategory = ''
                    $originalDate = ''
                    $convertedDate = ''
                }

                $originalFile = ([System.Uri]$directLink).Segments[-1]

                $safeFile = Get-SafeFileName -Url $directLink -Title $convertedTitle

                $uniqueFilePath = ''

                if (-not($MetadataOnly)) {
                    $uniqueFilePath = Get-UniqueFilePath -Path $Path -File $safeFile

                    if (('{0}\{1}' -f $Path,$safeFile) -ne $uniqueFilePath) {
                        Write-Warning -Message ('{0}\{1} already exists so filename was changed to {2}' -f $Path,$safeFile,$uniqueFilePath)
                    }
                }

                if (-not($MetadataOnly)) {
                    Invoke-FileDownload -Url $directLink -Path $uniqueFilePath
                }

                $original = [pscustomobject]@{
                    Title = $originalTitle;
                    Description = '';
                    Abstract = '';
                    Date = $originalDate;
                    Link = $originalLink;
                    File = $originalFile;
                    Category = $originalCategory;
                }

                $converted = [pscustomobject]@{
                    Title = $convertedTitle;
                    Description = '';
                    Abstract = '';
                    Date = $convertedDate;
                    Link = $directLink;
                    File = $safeFile;
                    Category = $convertedCategory;
                }

                $hash = ''
                $size = 0

                if (-not($MetadataOnly)) {
                    $hash = (Get-FileHash -Path $uniqueFilePath -Algorithm SHA256).Hash
                    $size = (Get-Item -Path $uniqueFilePath).Length
                }

                $document = [pscustomobject]@{
                    Original = $original;
                    Converted = $converted;
                    DirectLink = $directLink;
                    DowloadedAs = $uniqueFilePath.Replace($Path + '\',''); # sanitize potentially sensitive path
                    SHA256 = $hash;
                    Size = $size;
                    Source = 'Current';
                    Protected = $false;
                }

                $documents.Add($document)
            } else {
                throw "regex failed on line $line"
            }

            $startIndex = $page.IndexOf($startToken, $endIndex+$endToken.Length, [StringComparison]::CurrentCultureIgnoreCase)
        }
    }

    return ,$documents
}

Function Get-ProtectedCybersecurityLibrary() {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='Path to save the archive to')]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory=$false, HelpMessage='Only retrieve metadata')]
        [ValidateNotNullOrEmpty()]
        [switch]$MetadataOnly
    )

    if (-not(Test-Path -Path $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory | Out-Null
    }

    $documents = New-Object System.Collections.Generic.List[pscustomobject]

    $page = Invoke-PageDownload -Url 'https://www.nsa.gov/what-we-do/cybersecurity'

    # <a href="https://apps.nsa.gov/PartnerLibrary/CSA_Cybersecurity_Advisory_Wordpress_Symposium.pdf">Advisory: WordPress Plugin "WP Symposium" Remote Code Execution CVE-2014-10021: (June 2018)</a>
    $startToken = '<li><a href="https://apps.nsa.gov/PartnerLibrary/'
    $endToken = '</li>'

    $startUnprotectedIndex = 0;
    $startProtectedIndex = 0;
    $startIndex = $page.IndexOf($startToken, [StringComparison]::CurrentCultureIgnoreCase)

    while($startIndex -ne -1) {
        $endIndex = $page.IndexOf($endToken, $startIndex, [StringComparison]::CurrentCultureIgnoreCase)

        if ($endIndex -ne -1) {

            $line = $page[$startIndex..($endIndex+$endToken.Length)] -join ''

            if($line -match '.*<a href=(?<Link>.*)>(?<Title>.*).*</a>.*') {
                $originalLink = $matches.Link.Replace('"','')
                $directLink = $matches.Link.Replace('"','')

                $originalTitle = $matches.Title

                if($originalTitle -match '(?<Category>\w+)(:){1}(?<Title>.*)(:){0,1}\((?<Date>[0-9A-Za-z ]+)\).*') {
                    $convertedTitle = Convert-Title -Title $matches.Title.Trim()
                    $originalCategory = $matches.Category.Trim()
                    $convertedCategory = $originalCategory.Trim()
                    $originalDate = $matches.Date.Trim()
                    $convertedDate = [DateTime]::ParseExact($originalDate, 'MMMM yyyy', [System.Globalization.CultureInfo]::CurrentCulture)
                } else {
                    Write-Warning -Message ('{0} did not match typical title pattern' -f $originalTitle)
                    $convertedTitle = Convert-Title -Title $originalTitle
                    $originalCategory = ''
                    $convertedCategory = ''
                    $originalDate = ''
                    $convertedDate = ''
                }

                $originalFile = ([System.Uri]$directLink).Segments[-1]

                $safeFile = Get-SafeFileName -Url $directLink -Title $convertedTitle

                $uniqueFilePath = ''

                if (-not($MetadataOnly)) {
                    $uniqueFilePath = Get-UniqueFilePath -Path $Path -File $safeFile

                    if (('{0}\{1}' -f $Path,$safeFile) -ne $uniqueFilePath) {
                        Write-Warning -Message ('{0}\{1} already exists so filename was changed to {2}' -f $Path,$safeFile,$uniqueFilePath)
                    }
                }

                if (-not($MetadataOnly)) {
                    Invoke-FileDownload -Url $directLink -Path $uniqueFilePath
                }

                $original = [pscustomobject]@{
                    Title = $originalTitle;
                    Description = '';
                    Abstract = '';
                    Date = $originalDate;
                    Link = $originalLink;
                    File = $originalFile;
                    Category = $originalCategory;
                }

                $converted = [pscustomobject]@{
                    Title = $convertedTitle;
                    Description = '';
                    Abstract = '';
                    Date = $convertedDate;
                    Link = $directLink;
                    File = $safeFile;
                    Category = $convertedCategory;
                }

                $hash = ''
                $size = 0

                if (-not($MetadataOnly)) {
                    $hash = (Get-FileHash -Path $uniqueFilePath -Algorithm SHA256).Hash
                    $size = (Get-Item -Path $uniqueFilePath).Length
                }

                $document = [pscustomobject]@{
                    Original = $original;
                    Converted = $converted;
                    DirectLink = $directLink;
                    DowloadedAs = $uniqueFilePath.Replace($Path + '\',''); # sanitize potentially sensitive path
                    SHA256 = $hash;
                    Size = $size;
                    Source = 'Current';
                    Protected = $true;
                }

                $documents.Add($document)
            } else {
                throw "regex failed on line $line"
            }

            $startIndex = $page.IndexOf($startToken, $endIndex+$endToken.Length, [StringComparison]::CurrentCultureIgnoreCase)
        }
    }

    return ,$documents
}

Function Get-ArchivedCybersecurityLibrary() {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='Path to save the archive to')]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory=$false, HelpMessage='Only retrieve metadata')]
        [ValidateNotNullOrEmpty()]
        [switch]$MetadataOnly
    )

    if (-not(Test-Path -Path $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory | Out-Null
    }

    $documents = New-Object System.Collections.Generic.List[pscustomobject]

    $rss = Invoke-PageDownload -Url 'https://apps.nsa.gov/iaarchive/library/index.cfm?xml=IAD%20Library,RSS2.0-Custom'

    $xml = ''

    try {
        $xml = [xml]$rss
    } catch {
        throw 'RSS feed cannot be converted to XML'
    }

    $lastBuildDate = [DateTime]::ParseExact($xml.rss.channel.lastBuildDate, 'ddd, dd MMM yyyy HH:mm:ss +0000', [System.Globalization.CultureInfo]::CurrentCulture)

    $xml.rss.channel.item | ForEach-Object {
        $originalTitle = $_.title
        $originalDescription = $_.description
        $originalDate = $_.pubDate
        $originalLink = $_.link

        $convertedTitle = Convert-Title -Title $_.title
        $convertedDescription = Convert-Description -Description $_.description
        $convertedDate = [DateTime]::ParseExact($_.pubDate, 'ddd, dd MMM yyyy HH:mm:ss +0000', [System.Globalization.CultureInfo]::CurrentCulture)
        $convertedLink = $_.link.Replace('http://','https://')
        #http://www.iad.gov/iad/library/supporting-documents/faq/identity-theft-threat-and-mitigations.cfm
        #https://apps.nsa.gov/iaarchive/library/supporting-documents/faq/identity-theft-threat-and-mitigations.cfm

        $convertedLink = $convertedLink.Replace('https://www.iad.gov/iad/','https://apps.nsa.gov/iaarchive/')

        $content = Invoke-PageDownload -Url $convertedLink

        $directLink = Get-DirectLink -Content $content

        if ($directLink -eq '') {
            Write-Warning -Message ('Unable to get direct file link from {0}' -f $convertedLink)
            return
        }

        $abstract = Get-Abstract -Content $content

        #if ([System.Web.HttpUtility]::HtmlDecode($_.description) -ne $abstract) {
        #    Write-Warning -Message ('RSS feed description and web page abstract do not match for {0}' -f $convertedLink)
        #    Write-Warning -Message ('RSS: {0}' -f [System.Web.HttpUtility]::HtmlDecode($_.description))
        #    Write-Warning -Message ('Web: {0}{1}' -f $abstract,[System.Environment]::NewLine)
        #}

        $convertedAbstract = Convert-Description -Description $abstract
        $convertedAbstract = Convert-Abstract -Abstract $convertedAbstract
        Add-Type -AssemblyName System.Web
        if ([System.Web.HttpUtility]::HtmlDecode($convertedDescription) -ne $convertedAbstract) {
            Write-Warning -Message ('RSS feed description and web page abstract do not match for {0}' -f $convertedLink)
            Write-Warning -Message ('RSS: {0}' -f [System.Web.HttpUtility]::HtmlDecode($convertedDescription))
            Write-Warning -Message ('Web: {0}{1}' -f $convertedAbstract,[System.Environment]::NewLine)
        }

        $originalFile = ([System.Uri]$directLink).Segments[-1]

        $safeFile = Get-SafeFileName -Url $directLink -Title $convertedTitle

        $uniqueFilePath = ''

        if (-not($MetadataOnly)) {
            $uniqueFilePath = Get-UniqueFilePath -Path $Path -File $safeFile

            if (('{0}\{1}' -f $Path,$safeFile) -ne $uniqueFilePath) {
                Write-Warning -Message ('{0}\{1} already exists so filename was changed to {2}' -f $Path,$safeFile,$uniqueFilePath)
            }
        }

        $category = Get-Category -Url $convertedLink

        $convertedCategory = Convert-Category -Url $convertedLink

        if (-not($MetadataOnly)) {
            Invoke-FileDownload -Url $directLink -Path $uniqueFilePath
        }

        $original = [pscustomobject]@{
            Title = $originalTitle;
            Description = $originalDescription;
            Abstract = $abstract;
            Date = $originalDate;
            Link = $originalLink;
            File = $originalFile;
            Category = $category;
        }

        $converted = [pscustomobject]@{
            Title = $convertedTitle;
            Description = $convertedDescription;
            Abstract = $convertedAbstract;
            Date = $convertedDate;
            Link = $convertedLink;
            File = $safeFile;
            Category = $convertedCategory;
        }

        $hash = ''
        $size = 0

        if (-not($MetadataOnly)) {
            $hash = (Get-FileHash -Path $uniqueFilePath -Algorithm SHA256).Hash
            $size = (Get-Item -Path $uniqueFilePath).Length
        }

        $document = [pscustomobject]@{
            Original = $original;
            Converted = $converted;
            DirectLink = $directLink;
            DowloadedAs = $uniqueFilePath.Replace($Path + '\',''); # sanitize potentially sensitive path
            SHA256 = $hash;
            Size = $size;
            Source = 'Archive';
            Protected = $false;
        }

        $documents.Add($document)
    }

    return ,$documents
}

Function Get-CybersecurityLibraryMarkdown() {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory=$true, HelpMessage='Path to library JSON file')]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $shortcuts = @{}

    $content = Get-Content -Path $Path -Raw

    $library = $content | ConvertFrom-Json

    $librarySummary = @($library | ForEach-Object {
        $publicationSummary = [pscustomobject]@{
            Title = $_.Converted.Title;
            Description = $_.Converted.Description;
            Abstract = $_.Converted.Abstract
            Date = $_.Converted.Date;
            Link = $_.Converted.Link;
            Category = $_.Converted.Category;
            Hash = $_.SHA256
            Size = $_.Size;
            Source = $_.Source;
            Protected = $_.Protected;
        }
       $publicationSummary
    } | Sort-Object -Property Date -Descending)

    $shortcuts = @{}

    $pageBuilder = New-Object System.Text.StringBuilder
    $tocBuilder = New-Object System.Text.StringBuilder
    $contentBuilder = New-Object System.Text.StringBuilder

    $headerTemplate = @"
# NSA Cybersecurity publications

This page lists NSA Cybersecurity publications.

* Current NSA Cybersecurity publications can be found under the **Resources for Cybersecurity Professionals** section at <https://www.nsa.gov/what-we-do/cybersecurity/>
* Archived NSA Information Assurance and Information Assurance Directorate publications can be found at <https://apps.nsa.gov/iaarchive/library> (formerly  <https://www.iad.gov>)

A zip file containing publications from both pages can be downloaded from <https://github.com/nsacyber/nsacyber.github.io/releases/latest>

\* notes when authorization is required to access a publication.

## Table of Contents

| Title | Location | Date | Size |
| --- | --- | --- | --- |
"@

    [void]$tocBuilder.AppendLine($headerTemplate)

    $entryTemplate = @"
### {0}

* Abstract: {1}
* Date: {2:MM/dd/yyyy}
* Link: <{3}>
* Category: {4}
* SHA256: {5}
* Size: {6:n0}{7}
* Location: {8}
* Access Controlled: {9}

Return to the [Table of Contents](#table-of-contents).

"@

    $rowTemplate = '| [{0}]({1}) ([more...](#{2})){3} | {4} | {5:MMM yyyy} | {6:n0}{7} |'

    $librarySummary | ForEach-Object {
        $shortcut = $_.Title.ToLower() -replace '[^\p{L}\p{Nd}\s\-]','' -replace ' ','-'

        $count = 0
        $uniqueShortcut = $shortcut

        while($shortcuts.ContainsKey($uniqueShortcut)){
            $count++
            $uniqueShortcut = '{0}-{1}' -f $shortcut,$count
        }

        $shortcuts.Add($uniqueShortcut, '')

        if ($_.Protected) {
            $row = $rowTemplate -f $_.Title,$_.Link,$uniqueShortcut,'*',$_.Source,$_.Date,'',''
        } else {
            $row = $rowTemplate -f $_.Title,$_.Link,$uniqueShortcut,'',$_.Source,$_.Date,[Math]::Round(($_.Size/1KB)),'KB'
        }

        [void]$tocBuilder.AppendLine($row)

        if ($_.Protected) {
            $entry = $entryTemplate -f $_.Title,$_.Abstract,$_.Date,$_.Link,$_.Category,'','','',$_.Source,$_.Protected
        } else {
            $entry = $entryTemplate -f $_.Title,$_.Abstract,$_.Date,$_.Link,$_.Category,$_.Hash,[Math]::Round(($_.Size/1KB)),'KB',$_.Source,$_.Protected
        }
        [void]$contentBuilder.AppendLine($entry)
    }

    [void]$pageBuilder.AppendLine($tocBuilder.ToString())
    [void]$pageBuilder.AppendLine('')
    [void]$pageBuilder.AppendLine('## Publications')
    [void]$pageBuilder.AppendLine('')
    [void]$pageBuilder.AppendLine($contentBuilder.ToString())

    return $pageBuilder.ToString().Trim()
}