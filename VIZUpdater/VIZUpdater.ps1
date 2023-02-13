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

    [System.String]$HandlePath  = "$($GLOBAL:WorkingDirectory)\ZGEHandle.whnd"
    [System.String]$FLPIDPath   = "$($GLOBAL:WorkingDirectory)\FLPID.dat"
    [System.Int32]$CurrentFLPID = (Get-Process -Name 'FL64' -ErrorAction SilentlyContinue).Id
    If (!$CurrentFLPID) {
        [System.Int32]$CurrentY = $Host.UI.RawUI.CursorPosition.Y
        Do {
            [System.Console]::SetCursorPosition(0, $CurrentY)
            Clear-ConsoleLine -Lines 2
            Write-Host -ForegroundColor Red 'FL Studio is not running. Launch and prepare FL Studio, then press any key to continue.'
            [System.Void](Read-Host)
            [System.Int32]$CurrentFLPID = (Get-Process -Name 'FL64' -ErrorAction SilentlyContinue).Id
        } Until ($CurrentFLPID)
    }
        
    # Use previous window handle if current FL Studio PID matches stored PID.
    If ([System.IO.File]::Exists($HandlePath)) {
        [System.IntPtr]$ZGEWindowHandle = [System.Int32][System.String](Import-CLIXML -Path $HandlePath) # Many typecasts because stinky deserialized
        If ([System.IO.File]::Exists($FLPIDPath)) {
            [System.Int32]$StoredFLPID = Import-CLIXML -Path $FLPIDPath
            If ($CurrentFLPID -eq $StoredFLPID) {
                Write-Host -ForegroundColor Green "Using previous ZGE window handle ($($ZGEWindowHandle.ToString()))"
                Return $ZGEWindowHandle
            }
        }
        $CurrentFLPID | Export-CLIXML -Path $FLPIDPath -Force # Overwrite previously stored PID and delete stored handle and variable if current and stored PIDs don't match
        Remove-Variable -Name ZGEWindowHandle
        Remove-Item -Path $HandlePath -Force
    }

    [System.Int32]$OriginalY = $Host.UI.RawUI.CursorPosition.Y

    [System.Collections.Hashtable]$Regular = @{
        ForegroundColor = [System.ConsoleColor]::DarkCyan
        NoNewline       = $True
    }
    [System.Collections.Hashtable]$Highlight = @{
        ForegroundColor = [System.ConsoleColor]::DarkYellow
        NoNewline       = $True
    }
    [System.Collections.Hashtable]$Critical = @{
        ForegroundColor = [System.ConsoleColor]::Red
        NoNewline       = $True
    }

    Do {

        Write-Host -ForegroundColor Cyan '** Register ZGameEditor window handle **'
        Write-Host -ForegroundColor Cyan '----------------------------------------'

        Write-Host @Regular '1: In ZGameEditor - Select the '
        Write-Host @Highlight 'Add content'
        Write-Host @Regular " tab`n"

        Write-Host @Regular '2: In '
        Write-Host @Highlight 'Add content'
        Write-Host @Regular ' - Select the '
        Write-Host @Highlight 'Text'
        Write-Host @Regular " tab`n"

        Write-Host @Regular '3: Press '
        Write-Host @Highlight 'ENTER'
        Write-Host @Regular ' in '
        Write-Host @Highlight "this window`n"

        Write-Host @Regular '4: Within '
        Write-Host @Critical '5 seconds'
        Write-Host @Regular ' - Click the '
        Write-Host @Highlight 'Text field'
        Write-Host @Regular ' under the Text tab opened in steps 1-2, and '
        Write-Host @Highlight "wait for audio signal`n"

        [System.Void](Read-Host)
        [System.Int32]$CurrentY = $Host.UI.RawUI.CursorPosition.Y
        For ([System.Int32]$Tick = 5; $Tick -gt 0; $Tick--) {
            Write-Host -ForegroundColor Green          "  Focus ZGameEditor... ($($Tick))"
            Write-Host -NoNewline -ForegroundColor Red '  ~~~~~~~~~~~~~~~~~~~~~~~~'
            Start-Sleep -Seconds 1
            [System.Console]::SetCursorPosition(0, $CurrentY)
            Clear-ConsoleLine -Lines 2
        }

        # Get window handle of currently focused window
        [System.IntPtr]$ZGEWindowHandle = [Win32WindowActions]::GetForegroundWindow()

        # boop beep beep
        [System.Int32]$Pitch = 500
        For ($i = 0; $i -lt 3; $i++) {
            $Pitch += $i * 300
            [System.Console]::Beep($Pitch, 200)
        }

        # The ZGameEditor window handle does not show in Get-Process, thus allowing this basic check of whether the handle belongs to a different window. Retry if so.
        [System.Boolean]$IsInvalid = $False
        ForEach ($Handle in (Get-Process).MainWindowHandle) {
            If ($ZGEWindowHandle -eq [System.IntPtr]$Handle) {
                Write-Host -ForegroundColor Red 'The handle is not correct. Try again.'
                $IsInvalid = $True
                Start-Sleep -Seconds 3
                [System.Console]::SetCursorPosition(0, $OriginalY)
                Clear-ConsoleLine -Lines 10
                Break
            }
        }
        If ($IsInvalid) {Continue} Else {Break}

    } While ($True)

    Write-Host -ForegroundColor Green "ZGameEditor window handle: $($ZGEWindowHandle.ToString())"

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

Function Write-PinnedHeader {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [System.Version]$Version,
        [Parameter(Mandatory = $True)]
        [System.IntPtr]$ZGEWindowHandle,
        [Parameter(Mandatory = $True)]
        [System.Int32]$RequestInterval,
        [Parameter(Mandatory = $True)]
        [System.Double]$TokenAge,
        [Parameter(Mandatory = $True)]
        [System.Double]$TokenExpires
    )
    [System.TimeSpan]$Runtime = New-TimeSpan -Start $GLOBAL:StartupTime -End (Get-Date)
    [System.Collections.Hashtable]$VibeTime = @{
        Hours   = [System.String][System.Math]::Floor($Runtime.TotalHours)
        Minutes = [System.String]$Runtime.Minutes
    }
    [System.String]$VibeString = ''
    If ($VibeTime.Hours -gt 0) {
        $VibeString = "for $($VibeTime.Hours) hour"
        If ($VibeTime.Hours -ne 1) {$VibeString += 's'}
    }
    ElseIf ($VibeTime.Minutes -gt 0) {
        $VibeString = "for $($VibeTime.Minutes) minute"
        If ($VibeTime.Minutes -ne 1) {$VibeString += 's'}
    }
    If ($VibeTime.Minutes -gt 0 -And $VibeTime.Hours -gt 0) {
        $VibeString += " and $($VibeTime.Minutes) minute"
        If ($VibeTime.Minutes -ne 1) {$VibeString += 's'}
    }
    $VibeString += '!'
    [System.Int32]$AgeMins    = $TokenAge / 60
    [System.Int32]$ExpireMins = $TokenExpires / 6000
    [System.ConsoleColor]$TokenColor = Switch ($AgeMins) {
        {$_ -ge ($ExpireMins * 100)} {[System.ConsoleColor]::Red; Break}        # >100% elapsed (expired)
        {$_ -ge ($Expiremins * 90)}  {[System.ConsoleColor]::DarkYellow; Break} # >90% elapsed
        {$_ -ge ($ExpireMins * 75)}  {[System.ConsoleColor]::Yellow; Break}     # >75% elapsed
        Default                      {[System.ConsoleColor]::DarkCyan; Break}   # Normal
    }
    [System.Int32]$CurrentY = $Host.UI.RawUI.CursorPosition.Y
    [System.Console]::SetCursorPosition(0, 0)
    Clear-ConsoleLine -Lines 7
    Write-Host -ForegroundColor Cyan                "ZGEViz/Spotify Synchronizer  -  Straight vibin' $($VibeString)`n"
    Write-Host -ForegroundColor DarkCyan            "Version:          $Version"
    Write-Host -ForegroundColor DarkCyan            "ZGE Handle:       $ZGEWindowHandle"
    Write-Host -ForegroundColor DarkCyan            "Spotify PID:      $($GLOBAL:SpotifyPID)"
    Write-Host -ForegroundColor DarkCyan            "Request Interval: $RequestInterval"
    Write-Host -NoNewline -ForegroundColor DarkCyan "API token age:    "
    Write-Host -ForegroundColor $TokenColor                           "$AgeMins minutes"
    Clear-ConsoleLine -FillChar '-'
    [System.Console]::SetCursorPosition(0, $CurrentY)
}

Function Clear-ConsoleLine {
    Param (
        [System.Int32]$Lines   = 1,
        [System.Int32]$From    = $Host.UI.RawUI.CursorPosition.Y,
        [System.Char]$FillChar = ' '
    )
    # Store initial cursor Y position
    [System.Int32]$InitialY = $Host.UI.RawUI.CursorPosition.Y

    # Generate string of given character with length equal to the horizontal buffer size, then repeat for specified number of lines
    [System.String]$ClearLine = -Join ([System.String]$FillChar * $Host.UI.RawUI.BufferSize.Width)
    [System.String[]]$ClearString = $ClearLine * $Lines

    # Position cursor and clear lines
    [System.Console]::SetCursorPosition(0, $From)
    Write-Host ($ClearString -Join "`n")

    # Revert to initial cursor position
    If (!$PSBoundParameters.ContainsKey('FillChar')) {[System.Console]::SetCursorPosition(0, $InitialY)}
}

Function Get-SpotifyOAuth {
    Param (
        [System.String]$Set
    )
    # Function for fetching the Spotify OAuth token

    # Path for credential storage
    [System.String]$CredPath = "$($GLOBAL:WorkingDirectory)\SptAPI.cred"

    [System.Int32]$OriginalY = $Host.UI.RawUI.CursorPosition.Y

    [System.Boolean]$LoadedFromPrevious = $False

    Do {
        
        # If no credential is stored, prompt user for a token, convert to PSCredential and write the converted object to disk if valid
        If (![System.IO.File]::Exists($CredPath) -Or $Set) {

            If ($Set) {[System.Security.SecureString]$APIToken = ConvertTo-SecureString -String $Set -AsPlainText -Force}
            Else {
                [System.Console]::CursorVisible = $True
                [System.Security.SecureString]$APIToken = Read-Host -Prompt 'Enter your Spotify API Token (OAuth)' -AsSecureString
                [System.Console]::CursorVisible = $False
            }
            
            # Test token validity
            Try {
                [System.Management.Automation.PSCredential]$APICredential = New-Object System.Management.Automation.PSCredential -ArgumentList 'APIKEY', $APIToken
                [System.Void](Get-APISongData -Credential $APICredential)
                $APICredential | Export-CLIXML -Path $CredPath -ErrorAction SilentlyContinue
                Break
            } Catch {
                If ($Set) {
                    Write-Host -ForegroundColor Red 'Invalid token. Enter manually.'
                    Remove-Variable -Name Set
                } Else {Write-Host -ForegroundColor Red 'Invalid token. Try again.'}
                # Retry if invalid
                If ([System.IO.File]::Exists($CredPath)) {Remove-Item -Path $CredPath -Force}
            }

        } Else {

            # Read stored token and test its validity
            Try {
                [System.Management.Automation.PSCredential]$APICredential = Import-CLIXML -Path $CredPath
                [System.Void](Get-APISongData -Credential $APICredential)
                # If token's validity test doesn't throw an error, read its creating timestamp
                [System.DateTime]$GLOBAL:TokenCreationTime = Import-CLIXML -Path "$($GLOBAL:WorkingDirectory)\TokenCreation.dat"
                $LoadedFromPrevious = $True
                Break
            } Catch {
                # Prompt for new token if invalid
                Write-Host -ForegroundColor Red 'Token has likely expired. Enter new...'
                If ([System.IO.File]::Exists($CredPath)) {Remove-Item -Path $CredPath -Force}
            }

        }
        Start-Sleep -Seconds 3
        [System.Console]::SetCursorPosition(0, $OriginalY)
        Clear-ConsoleLine -Lines 5

    } While ($True)
    If ($Set) {Write-Host -ForegroundColor Green 'Refreshed API token.'}
    Else      {Write-Host -ForegroundColor Green 'API token validated.'}
    Start-Sleep -Seconds 2

    # Update token creation time if a new token has been created
    If (!$LoadedFromPrevious) {
        $GLOBAL:TokenCreationTime = Get-Date
        $GLOBAL:TokenCreationTime | Export-CLIXML -Path "$($GLOBAL:WorkingDirectory)\TokenCreation.dat" -Force
    }
    [System.Void]$GLOBAL:TokenCreationTime # To avoid false IntelliSense warning

    Return $APICredential
}

Function Invoke-ZGameVizUpdater {

    # Hide console cursor
    [System.Console]::CursorVisible = $False

    # Define working directory. Create if it does not exist.
    [System.String]$ScriptName              = (Get-Item $PSCommandPath).BaseName
    [System.String]$GLOBAL:WorkingDirectory = "$($PSScriptRoot)\$($ScriptName)"
    If (![System.IO.Directory]::Exists($GLOBAL:WorkingDirectory)) {New-Item -Path $GLOBAL:WorkingDirectory -ItemType Directory}
    Set-Location $GLOBAL:WorkingDirectory

    # Load C# classes for Win32API interactions
    [System.String]$Declarations = (Get-Content -Path @('Win32API.cs', "*.class")) -Join "`n"
    [System.String[]]$Assemblies = @(
        "System",
        #"System.Drawing", # Disabled because Win32WindowActions might not be needed
        #"System.Drawing.Primitives", # Disabled because Win32WindowActions might not be needed
        "System.Runtime.InteropServices",
        "System.Windows.Forms"
    )

    # Add C# classes
    Add-Type -Language CSharp -TypeDefinition $Declarations -ReferencedAssemblies $Assemblies
    [System.Void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [System.Void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    # Initialize objects
        [System.Version]$Version                                  = '23.2.13.730'      # Set-Clipboard "'$(Get-Date -Format "y.M.d.Hmm")'"
        [System.DateTime]$GLOBAL:StartupTime                      = Get-Date           # Startup timestamp
        [System.Void]$GLOBAL:StartupTime                                               # Because IntelliSense is silly
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
        [System.Double]$TokenWarn                                 = 3100               # Warning threshold for API token age in seconds (Default: 3100)
        [System.Int32]$TokenWarnInterval                          = 10                 # Warning trigger interval in number of iterations (interval x refresh rate)
        [System.Int32]$IterationsToWarning                        = $TokenWarnInterval # Warning trigger countdown (starts equal to warning interval)
        [System.Double]$TokenExpires                              = 3600               # API token TTL in seconds (Default: 3600)
        [System.Boolean]$MuteWarnings                             = $True              # Mutes token expiration warnings
    # Other
        [System.Int32]$FailedRequests                             = 0                  # Failed requests counter
        [System.Int32]$FailuresToTolerate                         = 3                  # Threshold for next-up overwrite due to API error
        [System.Boolean]$Waiting                                  = $False             # Wait flag
        [System.Boolean]$AutoFocus                                = $True              # Automatically focus ZGameEditor (Does nothing atm)
        [System.Double]$RefreshRate                               = 2                  # Refresh rate in seconds
        [System.String[]]$LiveTrack                               = @('', '')          # Live track data
        [System.Int32]$StartupDelay                               = 10                 # Seconds to wait before starting execution
        [System.Int32]$ReminderIndex                              = 0                  # Rolling reminder array index
        [System.String[]]$Reminders = @(                                               # Rolling reminders
            'Stay hydrated, astronaut :)',
            'Don''t forget to hydrate :)',
            'Got your drink? :)'
        ) # Rolling reminders
        [System.Collections.Hashtable]$PinnedHeader = @{                               # Splat object for Write-PinnedHeader
            Version         = $Version
            ZGEWindowHandle = $ZGEWindowHandle
            RequestInterval = $RequestInterval
            TokenAge        = $TokenAge
            TokenExpires    = $TokenExpires
        } # Splat object for Write-PinnedHeader
        [System.String[]]$TextTemplate = @(                                            # Output template with ??DUMMYVALUES???
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
            '??REMINDER??'
        ) # Output template with ??DUMMYVALUES???
        [System.String]$StoredOutput                          = $TextTemplate -Join "`n" # Currently displayed output
        [System.Management.Automation.ScriptBlock]$AsyncOAuth = { # ScriptBlock for asynchronous authentication job
            [System.Management.Automation.ScriptBlock]$AsyncFocus = {
                $Sig = '[DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);'
                $Focus = Add-Type -MemberDefinition $Sig -Name WindowAPI -PassThru
                Start-Sleep -Seconds 1
                [System.IntPtr]$Handle = (Get-Process -Name 'pwsh' | Where-Object {$_.MainWindowTitle -eq 'Spotify API token soon expired'}).MainWindowHandle
                [System.Void]$Focus::SetForegroundWindow($Handle)
            }
            [System.Void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            While (!$Valid) {
                [ThreadJob.ThreadJob]$FocusJob = Start-ThreadJob -ScriptBlock $AsyncFocus -Name 'Focus'
                [System.String]$Text = [Microsoft.VisualBasic.Interaction]::InputBox('Enter new Spotify auth token', 'Spotify API token soon expired', 'Enter token...')
                Remove-Job -Id $FocusJob.Id -Force
                Remove-Variable -Name $FocusJob
                If ($Text -Match 'BQ[a-zA-Z0-9_\-]{174}') {$Valid = $True}
                Else {[System.Void]([Microsoft.VisualBasic.Interaction]::MsgBox('The token is invalid. Try again.', 16, 'Invalid API token'))}
            }
            Write-Output $Text
        } # ScriptBlock for asynchronous authentication job
        
    # Display startup message
    Clear-Host
    Write-Host -ForegroundColor Cyan     "ZGEViz/Spotify Synchronizer  -  Just started vibin'`n"
    Write-Host -ForegroundColor DarkCyan "Version:          $Version"
    Write-Host -ForegroundColor DarkCyan "ZGE Handle:       $ZGEWindowHandle"
    Write-Host -ForegroundColor DarkCyan "Spotify PID:      $($GLOBAL:SpotifyPID)"
    Write-Host -ForegroundColor DarkCyan "Request Interval: $RequestInterval"
    Write-Host -ForegroundColor DarkCyan "API token age:    ---"
    Clear-ConsoleLine -FillChar '-'
    For ([System.Int32]$SleepTime = $StartupDelay; $SleepTime -ge 0; $SleepTime--) {
        [System.String]$pChar = "$($SleepTime.ToString())"
        [System.Console]::SetCursorPosition(0, $Host.UI.RawUI.CursorPosition.Y)
        Write-Host -NoNewline -ForegroundColor Green "All set! Waiting $pChar seconds before starting..."
        Start-Sleep -Seconds 1
    }
    Clear-ConsoleLine

    Do {
        # Check asynchronous auth job state if it has been started
        If ($OAuthJob) {
            # Update job object
            $OAuthJob = Get-Job -Id $OAuthJob.Id
            # If job has completed successfully, get job output, generate new API credential and remove job and corresponding object
            If ($OAuthJob.State -eq 'Completed') {
                [System.String]$OAuthJobResults = Receive-Job -Id $OAuthJob.Id
                Remove-Job -Id $OAuthJob.Id -Force
                If ($OAuthJobResults) {[System.Management.Automation.PSCredential]$APICredential = Get-SpotifyOAuth -Set $OAuthJobResults}
                Remove-Variable -Name OAuthJob
            }
        }

        # Refresh token age and time since last API request
        $PreviousRequest = (New-TimeSpan -Start $GLOBAL:LastRequest       -End (Get-Date)).TotalSeconds
        $TokenAge        = (New-TimeSpan -Start $GLOBAL:TokenCreationTime -End (Get-Date)).TotalSeconds
        $PinnedHeader['TokenAge'] = $TokenAge
        If ($TokenAge -le $_TokenAge) {$IterationsToWarning = $TokenWarnInterval}

        # Check token age and alert if appropriate
        If ($TokenAge -ge $TokenWarn) {
            If (!$OAuthJob) {[ThreadJob.ThreadJob]$OAuthJob = Start-ThreadJob -ScriptBlock $AsyncOAuth -Name AsyncOAuth}
            If ($TokenAge -ge $TokenExpires -And $IterationsToWarning -eq 1) {
                If (!$MuteWarnings) {For ($i = 0; $i -lt 5; $i++) {[System.Console]::Beep(1000, 200)}}
                Write-Host -ForegroundColor Red "WARNING: API token expiration imminent!"
                $IterationsToWarning = $TokenWarnInterval
            } ElseIf ($IterationsToWarning -eq 1) {
                If (!$MuteWarnings) {For ($i = 0; $i -lt 2; $i++) {[System.Console]::Beep(250, 200)}}
                Write-Host -ForegroundColor DarkYellow "WARNING: API token expires in ~$([System.Math]::Round(($TokenExpires - $TokenAge) / 60)) minutes!"
                $IterationsToWarning = $TokenWarnInterval
            } Else {$IterationsToWarning--}
            Write-PinnedHeader @PinnedHeader
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
            Write-PinnedHeader @PinnedHeader

            If ($Update) {
                # Advance rolling reminder. Index loops back to 0 if greater than or equal to count (prevent index out of range)
                $ReminderIndex++
                If ($ReminderIndex -ge $Reminders.Count) {$ReminderIndex = 0}
            }

            # Update Output
            [System.String]$Output = $TextTemplate -Join "`n"
            $Output = $Output.Replace('??TRACK??', $Player.Current.Track).Replace('??ARTIST??', $Player.Current.Artist)
            $Output = $Output.Replace('??NTRACK??', $Player.Next.Track).Replace('??NARTIST??', $Player.Next.Artist)
            $Output = $Output.Replace('??Q1TRACK??', $Player.Queue1.Track).Replace('??Q1ARTIST??', $Player.Queue1.Artist)
            $Output = $Output.Replace('??Q2TRACK??', $Player.Queue2.Track).Replace('??Q2ARTIST??', $Player.Queue2.Artist)
            $Output = $Output.Replace('??REMINDER??', $Reminders[$ReminderIndex])
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
                Write-PinnedHeader @PinnedHeader
                # Get current clipboard content
                [System.Object]$Clipboard = Get-Clipboard -Raw

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
                Start-Sleep -Milliseconds 100

                # Restore clipboard
                Set-Clipboard -Value $Clipboard

                $StoredOutput = $Output
                $LiveTrack    = $CurrentTrack
                $Waiting      = $False

            }

            ElseIf (!$Waiting) {
                Write-Host -ForegroundColor Yellow 'Song has updated. Waiting for ZGE focus...'
                $Waiting = $True
            }
        }

        Write-PinnedHeader @PinnedHeader

        Start-Sleep -Seconds $RefreshRate

        [System.Double]$_TokenAge = $TokenAge

    } While ($True)
}
Invoke-ZGameVizUpdater
