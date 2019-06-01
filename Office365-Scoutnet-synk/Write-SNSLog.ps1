# Default log file name.
$Script:SNSLogFilePath='scoutnetSync.log'

Export-ModuleMember -Variable SNSLogFilePath

function Write-SNSLog
{
<#
.Synopsis
   Write-SNSLog writes a message to a specified log file with the current time stamp.

.DESCRIPTION
   The Write-SNSLog function is designed to add logging capability to other scripts.
   In addition to writing output and/or verbose you can write to a log file for
   later debugging.
   The exported variable $SNSLogFilePath can be used to oweride the default logfile scoutnetSync.log.

.PARAMETER Message
   Message is the content that you wish to add to the log file.

.PARAMETER Level
   Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational)
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info")]
        [string]$Level="Info"
    )

    Process
    {
        if (-not [String]::IsNullOrWhiteSpace($script:SNSLogFilePath))
        {
            # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
            if (!(Test-Path $script:SNSLogFilePath))
            {
                Write-Verbose "Creating $script:SNSLogFilePath."
                New-Item $script:SNSLogFilePath -Force -ItemType File > $null
            }
        }
        else
        {
            Write-Warning "Log file path is empty. No log file will be created."
        }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Error $Message
                $LevelText = 'ERROR:'
                }
            'Warn' {
                Write-Warning $Message
                $LevelText = 'WARNING:'
                }
            'Info' {
                Write-Verbose $Message
                $LevelText = 'INFO:'
                }
            }

        if (-not [String]::IsNullOrWhiteSpace($script:SNSLogFilePath))
        {
            # Write log entry to $script:LogPath
            "$FormattedDate $LevelText $Message" | Out-File -FilePath $script:SNSLogFilePath -Append -Encoding UTF8
        }
    }
}
