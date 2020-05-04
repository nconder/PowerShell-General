function Save-KBFile {
    <#
    .SYNOPSIS
        Downloads patches from Microsoft

    .DESCRIPTION
         Downloads patches from Microsoft

    .PARAMETER Name
        The KB name or number. For example, KB4057119 or 4057119.

    .PARAMETER Path
        The directory to save the file.

    .PARAMETER FilePath
        The exact file name to save to, otherwise, it uses the name given by the webserver

    .PARAMETER Architecture
        Defaults to x64. Can be x64, x86 or "All"

    .NOTES
        Props to https://keithga.wordpress.com/2017/05/21/new-tool-get-the-latest-windows-10-cumulative-updates/
        Adapted for dbatools by Chrissy LeMaire (@cl)
        Then adapted again for general use without dbatools
        See https://github.com/sqlcollaborative/dbatools/pull/5863 for screenshots

    .EXAMPLE
        PS C:\> Save-KBFile -Name KB4057119

        Downloads KB4057119 to the current directory. This works for SQL Server or any other KB.

    .EXAMPLE
        PS C:\> Save-KBFile -Name KB4057119, 4057114 -Path C:\temp

        Downloads KB4057119 and the x64 version of KB4057114 to C:\temp. This works for SQL Server or any other KB.

    .EXAMPLE
        PS C:\> Save-KBFile -Name KB4057114 -Architecture All -Path C:\temp

        Downloads the x64 version of KB4057114 and the x86 version of KB4057114 to C:\temp. This works for SQL Server or any other KB.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Name,
        [string]$Path = ".",
        [string]$FilePath,
        [ValidateSet("x64", "x86", "All")]
        [string]$Architecture = "x64"
    )
    begin {
        function Get-KBLink {
            param(
                [Parameter(Mandatory)]
                [string]$Name
            )
            $kb = $Name.Replace("KB", "")
            $results = Invoke-WebRequest -Uri "http://www.catalog.update.microsoft.com/Search.aspx?q=KB$kb"
            $kbids = $results.InputFields |
                Where-Object { $_.type -eq 'Button' -and $_.Value -eq 'Download' } |
                Select-Object -ExpandProperty  ID

            Write-Verbose -Message "$kbids"

            if (-not $kbids) {
                Write-Warning -Message "No results found for $Name"
                return
            }

            $guids = $results.Links |
                Where-Object ID -match '_link' |
                Where-Object { $_.OuterHTML -match ( "(?=.*" + ( $Filter -join ")(?=.*" ) + ")" ) } |
                ForEach-Object { $_.id.replace('_link', '') } |
                Where-Object { $_ -in $kbids }

            if (-not $guids) {
                Write-Warning -Message "No file found for $Name"
                return
            }

            foreach ($guid in $guids) {
                Write-Verbose -Message "Downloading information for $guid"
                $post = @{ size = 0; updateID = $guid; uidInfo = $guid } | ConvertTo-Json -Compress
                $body = @{ updateIDs = "[$post]" }
                $links = Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $body |
                    Select-Object -ExpandProperty Content |
                    Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" |
                    Select-Object -Unique

                if (-not $links) {
                    Write-Warning -Message "No file found for $Name"
                    return
                }

                foreach ($link in $links) {
                    $link.matches.value
                }
            }
        }
    }
    process {
        if ($Name.Count -gt 0 -and $PSBoundParameters.FilePath) {
            throw "You can only specify one KB when using FilePath"
        }

        foreach ($kb in $Name) {
            $links = Get-KBLink -Name $kb

            if ($links.Count -gt 1 -and $Architecture -ne "All") {
                $templinks = $links | Where-Object { $PSItem -match "$($Architecture)_" }
                if ($templinks) {
                    $links = $templinks
                } else {
                    Write-Warning -Message "Could not find architecture match, downloading all"
                }
            }

            foreach ($link in $links) {
                if (-not $PSBoundParameters.FilePath) {
                    $FilePath = Split-Path -Path $link -Leaf
                } else {
                    $Path = Split-Path -Path $FilePath
                }

                $file = "$Path$([IO.Path]::DirectorySeparatorChar)$FilePath"

                if ((Get-Command Start-BitsTransfer -ErrorAction Ignore)) {
                    Start-BitsTransfer -Source $link -Destination $file
                } else {
                    # Invoke-WebRequest is crazy slow for large downloads
                    Write-Progress -Activity "Downloading $FilePath" -Id 1
                    (New-Object Net.WebClient).DownloadFile($link, $file)
                    Write-Progress -Activity "Downloading $FilePath" -Id 1 -Completed
                }
                if (Test-Path -Path $file) {
                    Get-ChildItem -Path $file
                }
            }
        }
    }
}
