Function Get-SpotifyPID {
    # Function for fetching process ID for Spotify.exe

    # Get all processes
    [System.Diagnostics.Process[]]$Processes = Get-Process -ProcessName 'Spotify'

    # Iterate processes until a process matches a path of "*\spotify.exe" and has a defined window title.
    ForEach ($Process in $Processes) {
        If ($Process.Path -Like '*\spotify.exe' -And $Process.MainWindowTitle -ne '') {

            # Get PID of matching process
            [System.Int32]$ProcessID = $Process.Id

            Break
        }
    }

    # Throw error if no instance is found
    If (!$ProcessID) {Throw 'No valid Spotify instance found.'}

    Return $ProcessID

}

Function Get-SongData {
    # Function for fetching currently playing track based on Spotify window title

    # Attempt to get the Spotify process by PID
    Try   {[System.Diagnostics.Process]$Process = Get-Process -Id ($GLOBAL:SpotifyPID)}
    Catch {
        # If a process with specified PID does not exist, try to find a new PID. Throw error if no new PID can be found.
        Write-Warning "The Spotify instance with PID $($GLOBAL:SpotifyPID) is no longer running. Trying to fetch new PID..."
        Try {
            [System.Int32]$ProcessID             = Get-SpotifyPID
            [System.Diagnostics.Process]$Process = Get-Process -Id $ProcessID
        } Catch {Throw $_}
        $GLOBAL:SpotifyPID = $ProcessID
    }

    # Split window title on the first ' - ' to separate artist and track name or set pause message
    If ($Process.MainWindowTitle -Like 'Spotify Premium') {[System.String[]]$SongData = @('Spotify Premium', 'Playback paused')}
    Else {[System.String[]]$SongData = ($Process.MainWindowTitle) -Split [RegEx]::Escape(' - '), 2}
    
    Return $SongData
}

Function Get-APISongData {
    # Function for fetching song data from the Spotify API

    [CmdletBinding()]

    Param ([Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential)

    # Assemble request parameter splat
    [System.Collections.Hashtable]$Request = @{
        Uri     = [System.Uri]'https://api.spotify.com/v1/me/player/queue'
        Method  = 'Get'
        Headers = [System.Collections.Hashtable]@{
            'Accept'        = 'application/json'
            'Content-Type'  = 'application/json'
            'Authorization' = "Bearer $($Credential.GetNetworkCredential().Password)"
        }
    }

    # Attempt API request, update timestamp for last request if successful
    Try     {[System.Object]$APIResponse = Invoke-RestMethod @Request -ErrorAction Stop}
    Catch   {Throw $_}
    Finally {$GLOBAL:LastRequest = Get-Date}
    [System.Void]$GLOBAL:LastRequest # To avoid false IntelliSense warning

    # Validate API return and generate parameter splat
    [System.Collections.Hashtable]$QueueData = @{}
    If (!$APIResponse.currently_playing.name) {$QueueData['CurrentArtist'], $QueueData['CurrentTrack'] = Get-SongData}
    Else {
        $QueueData['CurrentTrack']  = $APIResponse.currently_playing.name
        $QueueData['CurrentArtist'] = $APIResponse.currently_playing.artists.name -Join ', '
    }
    If ($APIResponse.queue[0].name) {
        $QueueData['NextTrack']    = $APIResponse.queue[0].name
        $QueueData['NextArtist']   = $APIResponse.queue[0].artists.name -Join ', '
        $QueueData['Queue1Track']  = $APIResponse.queue[1].name
        $QueueData['Queue1Artist'] = $APIResponse.queue[1].artists.name -Join ', '
        $QueueData['Queue2Track']  = $APIResponse.queue[2].name
        $QueueData['Queue2Artist'] = $APIResponse.queue[2].artists.name -Join ', '
    }
    
    # Return generated Spotify playback data
    Return New-PlaybackObject @QueueData
}

Function Get-ZGEHandle {
    # Function for fetching the ZGameEditor window handle

    [System.String]$HandlePath = "$($GLOBAL:WorkingDirectory)\ZGEHandle.whnd"
        
    # If a previous handle is stored, prompt user whether to use this or not
    If ([System.IO.File]::Exists($HandlePath)) {
        [System.IntPtr]$ZGEWindowHandle = [System.Int32][System.String](Import-CLIXML -Path $HandlePath) # Stinky Deserialized
        Do {
            Write-Host -ForegroundColor Green -NoNewline "Use previous handle? ($($ZGEWindowHandle)) [Y/N]: "
            [System.String]$Action = Read-Host

        } Until ($Action -Match [RegEx]'(?i)y|n')
        If ($Action -Like 'Y') {Return $ZGEWindowHandle}
        Else {
            Remove-Variable ZGEWindowHandle
            Remove-Item -Path $HandlePath -Force
        }
    }

    [System.Int32]$OriginalY = $Host.UI.RawUI.CursorPosition.Y

    Do {

        [System.String]$WaitMessage = 'Press ENTER, focus the ZGameEditor window within 5 seconds and wait for beeps, then come back.'
        Write-Host -NoNewline -ForegroundColor Green $WaitMessage
        [Void](Read-Host)
        [System.Console]::SetCursorPosition(0, $OriginalY)

        # Countdown
        For ([System.Int32]$SleepTime = 5; $SleepTime -gt 0; $SleepTime--) {
            Clear-ConsoleLine -Lines 1 -From $OriginalY
            Write-Host -ForegroundColor Green "Press ENTER, focus the ZGameEditor window within $($SleepTime.ToString()) seconds and wait for beeps, then come back."
            Start-Sleep -Seconds 1
        }
        Clear-ConsoleLine -Lines 1 -From $OriginalY

        # Get window handle of currently focused window
        [System.IntPtr]$ZGEWindowHandle = [Win32WindowActions]::GetForegroundWindow()

        # boop beep beep
        [System.Int32]$Pitch = 500
        For ($i = 0; $i -lt 3; $i++) {
            $Pitch += $i * 300
            [System.Console]::Beep($Pitch, 200)
        }
        Write-Host "Handle: $($ZGEWindowHandle.ToString())"

        # The ZGameEditor window handle does not show in Get-Process, thus allowing this basic check of whether the handle belongs to a different window. Retry if so.
        [System.Boolean]$IsInvalid = $False
        ForEach ($Handle in (Get-Process).MainWindowHandle) {
            If ($ZGEWindowHandle -eq [System.IntPtr]$Handle) {
                Clear-ConsoleLine -Lines 1 -From $OriginalY
                Write-Host -ForegroundColor Red 'The handle is invalid. Try again.'
                $IsInvalid = $True
                Start-Sleep -Seconds 2
                Break
            }
        }
        If ($IsInvalid) {Continue} Else {Break}

    } While ($True)

    $ZGEWindowHandle | Export-CLIXML -Path $HandlePath -ErrorAction SilentlyContinue

    Return $ZGEWindowHandle
}

Function Limit-String {
    # Function for limiting string length based on width in pixels

    [CmdletBinding()]

    Param (
        [Parameter(Position = 1)][System.String]$InputObject,
        [Parameter(Position = 2, Mandatory)][System.Int32]$Width
    )

    # Initialize required objects and perform measurement
    [System.Drawing.Font]$Font      = New-Object System.Drawing.Font('Roboto-Thin', 30)
    [System.String]$Out             = $Null
    [System.Drawing.Size]$InputSize = [System.Windows.Forms.TextRenderer]::MeasureText($InputObject, $Font)

    # Trim string if measured width is larger than specified limit
    If ($InputSize.Width -gt $Width) {

        # Loop until wifth is less than or equal to specified limit. Perform iterative measurements with '...' included
        Do {
            $InputObject = $InputObject.Substring(0, ($InputObject.Length - 2))
            $InputSize   = [System.Windows.Forms.TextRenderer]::MeasureText("$($InputObject)...", $Font)
        } While ($InputSize.Width -gt $Width)

        $Out = "$($InputObject)..."

    } Else {$Out = $InputObject}

    Return $Out
}

Function New-PlaybackObject {
    # Function for generating playback data object

    [CmdletBinding()]

    Param (
        [System.String]$CurrentTrack  = 'N/A',
        [System.String]$CurrentArtist = 'N/A',
        [System.String]$NextTrack     = 'N/A',
        [System.String]$NextArtist    = 'N/A',
        [System.String]$Queue1Track   = 'N/A',
        [System.String]$Queue1Artist  = 'N/A',
        [System.String]$Queue2Track   = 'N/A',
        [System.String]$Queue2Artist  = 'N/A'
    )

    # Assemble playback data object with limited strings
    [System.Collections.Hashtable]$PlaybackData = @{
        Current = [System.Collections.Hashtable]@{
            Track  = Limit-String -InputObject $CurrentTrack  -Width 610 # MAX 610
            Artist = Limit-String -InputObject $CurrentArtist -Width 765 # MAX 765
        }
        Next    = [System.Collections.Hashtable]@{
            Track  = Limit-String -InputObject $NextTrack     -Width 500 # MAX 500
            Artist = Limit-String -InputObject $NextArtist    -Width 530 # MAX 530
        }
        Queue1  = [System.Collections.Hashtable]@{
            Track  = Limit-String -InputObject $Queue1Track   -Width 500 # MAX 500
            Artist = Limit-String -InputObject $Queue1Artist  -Width 530 # MAX 530
        }
        Queue2  = [System.Collections.Hashtable]@{
            Track  = Limit-String -InputObject $Queue2Track   -Width 500 # MAX 500
            Artist = Limit-String -InputObject $Queue2Artist  -Width 530 # MAX 530
        }
    }

    Return $PlaybackData

}

Function Clear-ConsoleLine {
    Param (
        [System.Int32]$Lines = 1,
        [System.Int32]$From  = $Host.UI.RawUI.CursorPosition.Y
    )
    [System.String]$ClearString = ''
    For ([System.Int32]$i = 0; $i -lt $Host.UI.RawUI.BufferSize.Width - 2; $i++) {$ClearString += ' '}
    [System.Console]::SetCursorPosition(0, $From)
    For ([System.Int32]$i = 0; $i -lt $Lines; $i++) {Write-Host $ClearString}
    [System.Console]::SetCursorPosition(0, $From)
}

Function Get-SpotifyOAuth {
    # Function for fetching the Spotify OAuth token

    # Path for credential storage
    [System.String]$CredPath = "$($GLOBAL:WorkingDirectory)\SptAPI.cred"

    [System.Int32]$OriginalY = $Host.UI.RawUI.CursorPosition.Y 

    Do {
        
        # If no credential is stored, prompt user for a token, convert to PSCredential and write the converted object to disk if valid
        If (![System.IO.File]::Exists($CredPath)) {

            [System.Security.SecureString]$APIToken = Read-Host -Prompt 'Enter your Spotify API Token (OAuth)' -AsSecureString
            
            # Test token validity
            Try {
                [System.Management.Automation.PSCredential]$APICredential = New-Object System.Management.Automation.PSCredential -ArgumentList 'APIKEY', $APIToken
                [System.Void](Get-APISongData -Credential $APICredential)
                $APICredential | Export-CLIXML -Path $CredPath -ErrorAction SilentlyContinue
                Break
            } Catch {
                # Retry if invalid
                Write-Host -ForegroundColor Red 'Invalid token. Try again.'
                If ([System.IO.File]::Exists($CredPath)) {Remove-Item -Path $CredPath -Force}
            }

        } Else {

            # Read stored token and test its validity
            Try {
                [System.Management.Automation.PSCredential]$APICredential = Import-CLIXML -Path $CredPath
                [System.Void](Get-APISongData -Credential $APICredential)
                Break
            } Catch {
                # Prompt for new token if invalid
                Write-Host -ForegroundColor Red 'Token has likely expired. Enter new...'
                If ([System.IO.File]::Exists($CredPath)) {Remove-Item -Path $CredPath -Force}
            }

        }
        Start-Sleep -Seconds 2
        Clear-ConsoleLine -Lines 5 -From $OriginalY

    } While ($True)
    Write-Host -ForegroundColor Green 'Token validated.'
    Start-Sleep -Seconds 2
    Clear-ConsoleLine -Lines 5 -From $OriginalY

    # Update token creation time
    $GLOBAL:TokenCreationTime = Get-Date
    [System.Void]$GLOBAL:TokenCreationTime # To avoid false IntelliSense warning

    Return $APICredential
}

Function MAIN {

    # Define working directory. Create if it does not exist.
    [System.String]$ScriptName              = (Get-Item $PSCommandPath).BaseName
    [System.String]$GLOBAL:WorkingDirectory = "$($PSScriptRoot)\$($ScriptName)"
    If (![System.IO.Directory]::Exists($GLOBAL:WorkingDirectory)) {New-Item -Path $GLOBAL:WorkingDirectory -ItemType Directory}
    Set-Location $GLOBAL:WorkingDirectory

    # Load C# classes for Win32API interactions
    [System.String]$Declarations = (Get-Content -Path @('Win32API.cs', "*.class")) -Join "`n"
    [System.String[]]$Assemblies = @(
        "System",
        "System.Runtime.InteropServices",
        "System.Windows.Forms"
    )

    # Add C# classes
    Add-Type -Language CSharp -TypeDefinition $Declarations -ReferencedAssemblies $Assemblies
    [System.Void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [System.Void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    # Initialize objects
        [System.Version]$Version                                  = '23.1.3.6'         # Set-Clipboard "'$(Get-Date -Format "y.M.H.m")'"
    # Globals
        [System.DateTime]$GLOBAL:TokenCreationTime                = 0                  # API token creation time
        [System.DateTime]$GLOBAL:LastRequest                      = Get-Date           # Last API request timestamp
        [System.Int32]$GLOBAL:SpotifyPID                          = Get-SpotifyPID     # Spotify process ID
    # ZGameEditor/FL Studio
        [System.IntPtr]$ZGEWindowHandle                           = Get-ZGEHandle      # ZGameEditor window handle
    # API Interaction
        [System.Management.Automation.PSCredential]$APICredential = Get-SpotifyOAuth   # Spotify API token <- Defines TokenCreationTime!
        [System.Double]$RequestInterval                           = 45                 # Interval between each API request in seconds
        [System.Double]$PreviousRequest                           = 0                  # Seconds since previous API request
    # API token expiration reminder
        [System.Double]$TokenAge                                  = 0                  # API token age in seconds <- Is calculated on every iteration in main loop!
        [System.Double]$TokenWarn                                 = 4000               # Warning threshold for API token age in seconds
        [System.Int32]$TokenWarnInterval                          = 10                 # Warning trigger interval in number of iterations (interval x refresh rate)
        [System.Int32]$IterationsToWarning                        = $TokenWarnInterval # Warning trigger countdown (starts equal to warning interval)
        [System.Double]$TokenExpires                              = 3600               # API token TTL in seconds
    # Other
        [System.Int32]$FailedRequests                             = 0                  # Failed requests counter
        [System.Int32]$FailuresToTolerate                         = 3                  # Threshold for next-up overwrite due to API error
        [System.Boolean]$Waiting                                  = $False             # Wait flag
        [System.Boolean]$AutoFocus                                = $True              # Automatically focus ZGameEditor (Does nothing atm)
        [System.Double]$RefreshRate                               = 2                  # Refresh rate in seconds
        [System.String[]]$LiveTrack                               = @('', '')          # Live track data
        [System.String[]]$TextTemplate                            = @(                 # Output template with ??DUMMYVALUES???
            'Now playing',
            '- Wow! It''s better than Boofy!',
            "$([Char]0xe6)fg$([Char]0xf8)f Live Procedural Visualizer$([Char]0x2122) :)",
            'Ping/DM/Yell for request',
            '??TRACK??',
            '??ARTIST??',
            '',
            '',
            'Next up',
            '??NTRACK??',
            '??NARTIST??',
            '??Q1TRACK??',
            '??Q1ARTIST??',
            '??Q2TRACK??',
            '??Q2ARTIST??',
            'Stay hydrated, astronaut'
        )
        [System.String]$StoredOutput = $TextTemplate -Join "`n" # Currently displayed output

    # Display startup message
    Clear-Host
    Write-Host -ForegroundColor Cyan     "ZGEViz/Spotify Synchronizer`n"
    Write-Host -ForegroundColor DarkCyan "Version:          $Version"
    Write-Host -ForegroundColor DarkCyan "ZGE Handle:       $ZGEWindowHandle"
    Write-Host -ForegroundColor DarkCyan "Spotify PID:      $($GLOBAL:SpotifyPID)"
    Write-Host -ForegroundColor DarkCyan "Request Interval: $RequestInterval"
    Write-Host -ForegroundColor Yellow   '-----------------------------------------------------------------'
    For ([System.Int32]$SleepTime = 10; $SleepTime -ge 0; $SleepTime--) {
        [System.String]$pChar = "$($SleepTime.ToString())"
        [System.Console]::SetCursorPosition(0, $Host.UI.RawUI.CursorPosition.Y)
        Write-Host -NoNewline -ForegroundColor Green "All set! Waiting $pChar seconds before starting..."
        Start-Sleep -Seconds 1
    }
    [System.Console]::SetCursorPosition(0, $Host.UI.RawUI.CursorPosition.Y)
    Write-Host $("All set! Waiting $pChar seconds before starting..." -Replace '.*', ' ')

    Do {
        # Refresh token age and time since last API request
        $PreviousRequest = (New-TimeSpan -Start $GLOBAL:LastRequest       -End (Get-Date)).TotalSeconds
        $TokenAge        = (New-TimeSpan -Start $GLOBAL:TokenCreationTime -End (Get-Date)).TotalSeconds
        If ($TokenAge -le $_TokenAge) {$IterationsToWarning = $TokenWarnInterval}

        # Check token age and alert if appropriate
        If ($TokenAge -ge $TokenWarn) {
            If ($TokenAge -ge $TokenExpires -And $IterationsToWarning -eq 1) {
                For ($i = 0; $i -lt 5; $i++) {[System.Console]::Beep(1000, 200)}
                Write-Host -ForegroundColor Red "WARNING: API token expiration imminent!"
                $IterationsToWarning = $TokenWarnInterval
            } ElseIf ($IterationsToWarning -eq 1) {
                For ($i = 0; $i -lt 2; $i++) {[System.Console]::Beep(250, 200)}
                Write-Host -ForegroundColor DarkYellow "WARNING: API token expires in ~$([System.Math]::Round(($TokenExpires - $TokenAge) / 60)) minutes!"
                $IterationsToWarning = $TokenWarnInterval
            } Else {$IterationsToWarning--}
        }

        # Get currently playing track and compare to live track
        [System.String[]]$CurrentTrack = Get-SongData
        If ([System.String]($CurrentTrack -Join ' - ') -ne [System.String]($LiveTrack -Join ' - ') -And !$Waiting) {[System.Boolean]$Update = $True}
        Else {[System.Boolean]$Update = $False}

        # Update data if track has changed or time since last request has exceeded the interval
        If ($PreviousRequest -ge $RequestInterval -Or $Update) {

            # Attempt to get next track in the queue
            Try {
                If ($Update) {Write-Host -NoNewline 'New track playing. Refreshing playback data... '}
                [System.Collections.Hashtable]$Player = Get-APISongData -Credential $APICredential
                If ($Update) {Write-Host -ForegroundColor Green 'OK'}
                $FailedRequests = 0
                Remove-Variable ErrorData, ErrorCode, ErrorMessage -ErrorAction SilentlyContinue
            } Catch [Microsoft.PowerShell.Commands.HttpResponseException] {
                $FailedRequests += 1
                [System.Object]$ErrorData    = $_.ErrorDetails.Message | ConvertFrom-Json
                [System.String]$ErrorCode    = $ErrorData.error.status
                [System.String]$ErrorMessage = $ErrorData.error.message
            } Catch {
                $FailedRequests += 1
                [System.String]$ErrorCode    = '0'
                [System.String]$ErrorMessage = ($Error[0].Exception.Message -Split "`n")[0]
            }
            If ($ErrorCode) {
                # Beep boop galore and assign new API token if the current one has expired
                If (($ErrorCode -eq '401') -Or ($ErrorCode -eq '403')) {
                    Write-Host -ForegroundColor Red "`nAPI request failed: $ErrorCode - $ErrorMessage"
                    For ($i = 0; $i -lt 3; $i++) {[System.Console]::Beep(500, 200)}
                    [System.Management.Automation.PSCredential]$APICredential = Get-SpotifyOAuth
                    Continue
                } Else {

                    # Display error
                    [System.Collections.Hashtable]$PlaybackObjectParams = @{
                        CurrentTrack  = $CurrentTrack[1]
                        CurrentArtist = $CurrentTrack[0]
                        NextTrack     = 'SPOTIFY API ERROR'
                        NextArtist    = "$ErrorCode - $ErrorMessage"
                    }
                    Write-Host -ForegroundColor Red "`nAPI request failed: $ErrorCode - $ErrorMessage"

                    [System.Collections.Hashtable]$Player = New-PlaybackObject @PlaybackObjectParams
                }
            }

            # Update Output
            [System.String]$Output = $TextTemplate -Join "`n"
            $Output = $Output.Replace('??TRACK??', $Player.Current.Track).Replace('??ARTIST??', $Player.Current.Artist)
            $Output = $Output.Replace('??NTRACK??', $Player.Next.Track).Replace('??NARTIST??', $Player.Next.Artist)
            $Output = $Output.Replace('??Q1TRACK??', $Player.Queue1.Track).Replace('??Q1ARTIST??', $Player.Queue1.Artist)
            $Output = $Output.Replace('??Q2TRACK??', $Player.Queue2.Track).Replace('??Q2ARTIST??', $Player.Queue2.Artist)
        }

        # Update display if output has changed and the ZGameEditor window is in focus
        If ($StoredOutput -ne $Output -And $FailedRequests -NotIn 1..$FailuresToTolerate) {
            If ($AutoFocus) {
                # [Win32WindowActions]::SetForegroundWindow($ZGEWindowHandle) # Disabled - Doesn't focus input field, just window
                #If ([Win32WindowActions]::IsIconic($ZGEWindowHandle)) {[Win32WindowActions]::ShowWindow($ZGEWindowHandle, 9)} # This might not be needed
            }
            If ([Win32WindowActions]::GetForegroundWindow() -eq $ZGEWindowHandle) {
                Write-Host -ForegroundColor Green "Applying playback changes:"
                Write-Host -ForegroundColor Green "`tTrack: $($Player.Current.Track)"
                Write-Host -ForegroundColor Green "`t       $($Player.Current.Artist)"
                Write-Host -ForegroundColor Green "`t-------------------------------------------------------------"
                Write-Host -ForegroundColor Green "`tNext:  $($Player.Next.Track)"
                Write-Host -ForegroundColor Green "`t       $($Player.Next.Artist)"

                Set-Clipboard -Value $Output

                # CTRL + A
                [Win32KeyEvent]::Sendkey(0x11, 0) # CTRL key down
                Start-Sleep -Milliseconds 100
                [Win32KeyEvent]::Sendkey(0x41, 0) # A key down
                Start-Sleep -Milliseconds 50
                [Win32KeyEvent]::Sendkey(0x41, 2) # A key up
                Start-Sleep -Milliseconds 100
                [Win32KeyEvent]::Sendkey(0x11, 2) # CTRL key up
                Start-Sleep -Milliseconds 200

                # CTRL + V
                [Win32KeyEvent]::Sendkey(0x11, 0) # CTRL key down
                Start-Sleep -Milliseconds 100
                [Win32KeyEvent]::Sendkey(0x56, 0) # V key down
                Start-Sleep -Milliseconds 50
                [Win32KeyEvent]::Sendkey(0x56, 2) # V key up
                Start-Sleep -Milliseconds 100
                [Win32KeyEvent]::Sendkey(0x11, 2) # CTRL key up        

                $StoredOutput = $Output
                $LiveTrack    = $CurrentTrack
                $Waiting      = $False

            }

            ElseIf (!$Waiting) {
                Write-Host -ForegroundColor Yellow 'Song has updated. Waiting for ZGE focus...'
                $Waiting = $True
            }
        }

        Start-Sleep -Seconds $RefreshRate
        [System.Double]$_TokenAge = $TokenAge

    } While ($True)
}
Main
