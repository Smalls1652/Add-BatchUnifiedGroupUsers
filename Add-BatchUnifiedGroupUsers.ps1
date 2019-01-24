[cmdletbinding()]
param(
    [Parameter(Mandatory = $true)]$GroupName,
    [Parameter(Mandatory = $true)]$UserData,
    [Parameter(Mandatory = $true)][pscredential]$Credential,
    [int]$Interval = 1000,
    [switch]$Log,
    [string]$LogPath = ".\UnifiedGroupManager.$(Get-Date -Format "MM.dd.yyyy-HH.mm").log"

)

function logger {
    param(
        [string]$LogText
    )

    Write-Verbose $LogText

    if ($Log) {
        Write-Output "[$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")] $($LogText)" | Out-File -FilePath $LogPath -Append
    }
}

function startup {
    param(
        [PSCredential]$Credential
    )

    $ExchangeOnline = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection

    Import-PSSession $ExchangeOnline -DisableNameChecking -AllowClobber

    return $ExchangeOnline
}

function cleanup {
    param(
        $Session
    )

    Remove-PSSession -Session $Session
}

logger -LogText "Starting log at $(Get-Date -Format "MM.dd.yyyy-HH.mm")."

#Initializing the O365 session to run Exchange Online commands.
$ExchSess = startup -Credential $Credential

#Converting the supplied $UserData object to an ArrayList.
#Converting to an ArrayList is essential for removing objects from the list.
[System.Collections.ArrayList]$ConvertedUserData = $UserData
logger -LogText "Converted supplied user data to an ArrayList."

#Initializing batch counters
$i = 0
$z = $Interval
$AllCount = $ConvertedUserData | Measure-Object | Select-Object -ExpandProperty "Count"

#Setting $BatchFinished to $true, so that the correct logging text is provided on run.
$BatchFinished = $true

#Starting batch process.
logger -LogText "Starting the batch process."
while ($i -le $AllCount) {

    if ($z -gt $AllCount) {
        #If $z > $AllCount, we need to make sure $z = $AllCount. This is to prevent it going over the actual size of the $ConvertedUserData array.
        $z = $AllCount
    }

    if (($BatchFinished)) {
        #Logging the start of a new batch.
        logger -LogText "Starting range $($i) - $($z)..."
    }
    else {
        #Logging the restart of a batch after an error.
        logger -LogText "Restarting range $($i) - $($z)..."
    }

    try {

        #Adding the range of users to the Unified Office 365 group.
        Add-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links ($ConvertedUserData[$i..$z].Name) -Verbose -ErrorAction "Stop"

        logger -LogText "Range $($i) - $($z) finished."

        #Incrementating the batch counters and setting $BatchFinished to $true.
        $i = $i + $Interval
        $z = $z + $Interval
        $BatchFinished = $true
        $ErrorCounter = 0
    }
    catch [Exception] {
        $ErrorMessage = $_.Exception

        Switch ($ErrorMessage.HResult) {
            '-2146233087' {
                <#

                HResult Return: 2146233087
                
                Reason for Error: A user object in the list does not match a user in Office 365.

                Solution: Remove the user from the $ConvertedUserData array.
                
                #>
                $ErrorUser = $ErrorMessage.Message | Select-String '.* \"(.*)\".*'

                logger -LogText "An error occured. HRESULT: -2146233087 Reason: A user object in the list does not match a user in Office 365."

                $ConvertedUserData.Remove(($ConvertedUserData | Where-Object -Property "Name" -eq $ErrorUser.Matches[0].Groups[1].Value))

                logger -LogText "Removing $($ErrorUser.Matches[0].Groups[1].Value) from the UserData list."

                $AllCount = $ConvertedUserData | Measure-Object | Select-Object -ExpandProperty "Count"

                logger -LogText "UserData list size is now $($AllCount) objects."
            }
            Default {

                #Generic error return.
                logger -LogText "An unknown error occurred. See the next log message below for the error message returned."
                logger -LogText "$($ErrorMessage.Message)"
            }
        }

        if (($ErrorCounter -lt 10) -and ($ErrorCounter -gt 0)) {
            $ErrorCounter++
            logger -LogText "ErrorCounter = $($ErrorCounter)"
        }
        elseif ($ErrorCounter -eq 10) {
            logger -LogText "There has been 10 consecutive errors returned, breaking the batch loop."
            break
        }
        else {
            $ErrorCounter = 1
            logger -LogText "ErrorCounter = $($ErrorCounter)"
        }

        $BatchFinished = $false
    }
}

#Close out the Office 365 session.
cleanup -Session $ExchSess

#Log the script completion.
logger -LogText "Script has finished running."