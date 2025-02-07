#requires -Version 5
# This module is a collection of functions that can be used to manage Windows processes / Services and other windows related tasks

#region ProcessRelatives
# Process Relatives help to get the parent and child processes of a given process
# Thanks to Guy Leech for the Get-ProcessRelatives function that I just modified to be a part of this module
Class ProcessDetail {
    [String]    $Name
    [UInt32]    $Id
    [UInt32]    $ParentId
    [Int32]     $SessionId
    [String]    $Owner
    [DateTime]  $CreationDate
    [String]    $CommandLine
    [String]    $Service
    ProcessDetail($Process) {
        $this.Name = $Process.ProcessName
        $this.Id = $Process.ProcessID
        $this.ParentId = $Process.ParentId
        $this.SessionId = $Process.SessionId
        $this.CreationDate = $Process.CreationDate
        $this.CommandLine = $Process.CommandLine
    }
}
Function Get-ProcessRelatives {
    [CmdletBinding(DefaultParameterSetName='byName')]
    <#
        .SYNOPSIS
            Get parent and child processes details and recurse

        .DESCRIPTION
            Uses win32_process. Level 0 processes are those specified via parameter, positive levels are parent processes & negative levels are child processes
            Child processes are not guaranteed to be directly below their parent - check process id and parent process id

        .PARAMETER name
            A regular expression to match the name(s) of the process(es) to retrieve.

        .PARAMETER id
            The ID(s) of the process(es) to retrieve.

        .PARAMETER indentMultiplier
            The multiplier for the indentation level. Default is 1.

        .PARAMETER indenter
            The character(s) used for indentation. Default is a space.

        .PARAMETER unknownProcessName
            The placeholder name for unknown processes. Default is '<UNKNOWN>'.

        .PARAMETER properties
            The properties to retrieve for each process

        .PARAMETER quiet
            Suppresses warning output if specified.

        .PARAMETER norecurse
            Prevents recursion through processes if specified.

        .PARAMETER noIndent
            Disables creatring indented name if specified.

        .PARAMETER noChildren
            Excludes child processes from the output if specified.

        .PARAMETER noOwner
            Excludes the owner from the output if specified which speeds up the script.

        .PARAMETER sessionId
            Process all processes passed via -id or -name regardless of session if * is passed (default)
            Only process processes passed via -id or -name if they are in the same session as the script if -1 is passed
            Only process processes passed via -id or -name if they are in the same session as the value passed if it is a positive integer

        .EXAMPLE
        Get-ProcessRelatives -id 12345

        Get parent and child processes of the running process with pid 12345

        .EXAMPLE
        Get-ProcessRelatives -name notepad.exe,winword.exe -properties *

        Get parent and child processes of all running processes of notepad and winword, outputting all win32_process properties & added ones

        .EXAMPLE
        Get-ProcessRelatives -name powershell.exe -sessionid -1

        Get parent and child processes of powershll.exe processes running in the same session as the script

        .NOTES
            Modification History:

            2024/09/13  @guyrleech  Script born
            2024/09/16  @guyrleech  First release
            2024/09/17  @SwissPowerShell integrated in SPS-WinTools module
    #>
    Param(
        [Parameter(
            ValueFromPipelineByPropertyName,
            ParameterSetName='byName',
            Position=1,
            HelpMessage='The name(s) of the process(es) to retrieve.'
        )]
        [String[]] ${Name},
        [Parameter(
            ValueFromPipelineByPropertyName,
            ParameterSetName='byId',
            Mandatory,
            Position=1,
            HelpMessage='The ID(s) of the process(es) to retrieve.'
        )]
        [Int[]]     ${Id},
        [Parameter(HelpMessage='The session id to filter on.')]
        [String]    ${SessionId} = '*',
        [Parameter(HelpMessage='The multiplier for the indentation level.')]
        [Int]       ${IndentMultiplier} = 1,
        [Parameter(HelpMessage='The character(s) used for indentation.')]
        [String]    ${Indenter} = ' ',
        [Parameter(HelpMessage='The placeholder name for unknown processes.')]
        [String]    ${UnknownProcessName} = '<UNKNOWN>',
        [Parameter(HelpMessage='The properties to retrieve for each process')]
        [String[]]  ${Properties} = @( 'IndentedName' , 'ProcessId' , 'ParentProcessId' , 'Sessionid' , '-' , 'Owner' , 'CreationDate' , 'Level' , 'CommandLine' , 'Service' ),
        [Parameter(HelpMessage='Suppresses warning output if specified.')]
        [Switch]    ${Quiet},
        [Parameter(HelpMessage='Prevents recursion through processes if specified.')]
        [Switch]    ${Norecurse},
        [Parameter(HelpMessage='Disables creatring indented name if specified.')]
        [Switch]    ${NoIndent},
        [Parameter(HelpMessage='Excludes child processes from the output if specified.')]
        [Switch]    ${NoChildren},
        [Parameter(HelpMessage='Excludes the owner from the output if specified which speeds up the script.')]
        [Switch]    ${NoOwner}
    )
    BEGIN {
        Class NameComparer : System.Collections.Generic.IComparer[PSCustomObject]{
            NameComparer(){}
            [Int] Compare ([PSCustomObject]$X , [PSCustomObject]$Y) {
                ## cannot simply return difference directly since Compare must return int but uint32 could be bigger
                [Int64] $Difference = $X.ProcessId - $Y.ProcessId
                If ($Difference -eq 0){
                    return 0
                } ElseIf ($Difference -lt 0) {
                    return -1
                } Else {
                    return 1
                }
            }
        }
        Function Get-DirectRelativeProcessDetails {
            Param(
                [Int]       ${Id},
                [Int]       ${level} = 0,
                [DateTime]  ${Created},
                [Bool]      ${Children} = $false,
                [Switch]    ${Recurse},
                [Switch]    ${Quiet} ,
                [Switch]    ${FirstCall}
            )
            Write-Verbose -Message "Get-DirectRelativeProcessDetails pid $($Id) level $($Level)"
            $ProcessDetail = $null
            ## array is of win32_process objects where we order & search on process id
            [Int] $ProcessDetailIndex = $Processes.BinarySearch([PScustomobject]@{ProcessId = $Id}, $Comparer)
            If ($ProcessDetailIndex -ge 0){
                $ProcessDetail = $Processes[$ProcessDetailIndex]
            }
            ## else not found
            ## guard against pid re-use (do not need to check pid created after child process since could not exist before with same pid although can't guarantee that pid hasn't been reused since unless we check process auditing/sysmon)
            If (($null -ne $ProcessDetail) -and (($null -eq $Created) -or ((-not $children) -and ($processDetail.CreationDate -le $Created)) -or $Children)) {
                ## * means any session, -1 means session script is running in any other positive value is session id it process must be running in
                If (($SessionId -ne '*') -and ($FirstCall -eq $True)) {
                    If ($SessionIdAsInt -lt 0){
                        ## session for script only
                        If ($ProcessDetail.SessionId -ne $ThisSessionId){
                            $ProcessDetail = $null
                        }
                    }ElseIf ($SessionIdAsInt -ne $ProcessDetail.SessionId){
                        ## session id passed so check process is in this session
                        $ProcessDetail = $null
                    }
                }
                If (($null -ne $ProcessDetail) -and ($null -ne $ProcessDetail.ParentProcessId) -and ($ProcessDetail.ParentProcessId -gt 0)){
                    If ($Recurse -eq $True){
                        If ($Children){
                            $Processes | Where-Object ParentProcessId -eq $Id -PipelineVariable childProcess | ForEach-Object {
                                    Get-DirectRelativeProcessDetails -Id $ChildProcess.ProcessId -Level ($Level - 1) -Recurse -Children $true -Created $ProcessDetail.CreationDate -Quiet:$Quiet -Processes $Processes
                                }
                        }
                        If ($FirstCall -or (-not $Children)){
                            ## getting parents
                            Get-DirectRelativeProcessDetails -Id $ProcessDetail.ParentProcessId -Level ($Level + 1) -Children $false -Recurse  -Created $processDetail.CreationDate -Quiet:$Quiet -Processes $Processes
                        }
                    }

                    ## don't just look up svchost.exe as could be a service with it's own exe
                    [String] $Service = ($RunningServices[$ProcessDetail.ProcessId] | Select-Object -ExpandProperty Name) -join '/'

                    $Owner = $null
                    If  (-Not $noOwner){
                        If (-Not $processDetail.PSObject.Properties['Owner']){
                            $OwnerDetail = Invoke-CimMethod -InputObject $ProcessDetail -MethodName GetOwner -ErrorAction SilentlyContinue -Verbose:$false
                            If(($null -ne $OwnerDetail) -and ($OwnerDetail.ReturnValue -eq 0)){
                                $Owner = "$($OwnerDetail.Domain)\$($OwnerDetail.User)"
                            }
                            Add-Member -InputObject $ProcessDetail -MemberType NoteProperty -Name Owner -Value $Owner
                        }Else{
                            $Owner = $ProcessDetail.Owner
                        }
                    }
                    ## clone the process detail since may be used by another process being analysed and could be at a different level in that
                    ## clone() method not available in PS 7.x
                    $Clone = [CimInstance]::new($ProcessDetail)

                    Add-Member -InputObject $Clone -PassThru -NotePropertyMembers @{
                        ## return
                        Owner   = $Owner
                        Service = $Service
                        Level   = $Level
                        '-'     = $(if ($FirstCall) {'*'} Else {''})
                    }
                }
            ## else no parent or excluded based on session id
            } ElseIf ($FirstCall) {
                ## only warn on first call
                If (-not $Quiet){
                    Write-Warning "No process found for id $($Id)"
                }
            } ElseIf (-not $Quiet){
                ## TODO search process auditing/sysmon ?
                [pscustomobject]@{
                    Name = $UnknownProcessName
                    ProcessId = $Id
                    Level = $Level
                }
            }
        }

    }
    PROCESS {
        ## Get the session id of the current process
        [Int] $ThisSessionId = (Get-Process -id $Pid).SessionId
        # Convert the session id to an integer or null if the session Is is not a number
        $SessionIdAsInt = $SessionId -as [Int]
        ## use sorted array so can find quicker
        $Comparer = [NameComparer]::new()
        ## get all processes so quicker to find parents and children regardless of session id as only filter on session id of processes specified by paramter, not parent/child
        $Processes = [System.Collections.Generic.List[PSCustomObject]](Get-CimInstance -ClassName win32_process -Verbose:$false)
        ## sort so can binary search
        $Processes.Sort($Comparer)
        Write-Verbose -Message "Found: $($processes.Count) processes"
        # if the parameterset is byName then get the process id for the given name(s)
        if ($PSCmdlet.ParameterSetName -eq 'byName') {
            if ($Name.Count -eq 0) {
                # No Name parameter passed => get all processes
                $Name = @('.+')
            }
            $Id = @(ForEach ($ProcessName in $Name) {
                $Processes | Where-Object Name -Match $ProcessName | Select-Object -ExpandProperty ProcessId
            })
            if ($Id.Count -eq 0) {
                Throw $([System.Management.Automation.ErrorRecord]::new(
                        [System.Management.Automation.ItemNotFoundException]::new("No processes found for $($Name)"),
                        'NoProcessesFound',
                        [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                        $null
                    )
                )
            }
            Write-Verbose -Message "Found: $($Id.Count) pids for process(es) $($Name)"
        }
        # get all services as an hashtable so we can quickly look them up
        [HashTable] $RunningServices = @{}
        Get-CimInstance -ClassName win32_service -filter 'ProcessId > 0' -Verbose:$False| ForEach-Object {
            # could be multiple so store as array
            $Existing = $RunningServices[$_.ProcessId]
            If ($null -eq $Existing) {
                # the process is not already in the hashtable
                $RunningServices.Add($_.ProcessId, ([System.Collections.Generic.List[Object]]$_))
            } Else {
                # already in the hashtable have this pid
                $Existing.Add($_)
            }
        }
        Write-Verbose -Message "Found: $($RunningServices.Count) running service(s) pid(s)"
        [Array] $Results = @(
            ForEach ($ProcessId in $Id) {
                [Array] $Result = @(
                    Get-DirectRelativeProcessDetails -Id $ProcessId -Recurse:(-Not $Norecurse) -Quiet:$Quiet -Children (-Not $NoChildren) -FirstCall -Processes $Processes | Sort-Object -Property Level -Descending
                )
                # now we know how many levels we can indent so the topmost process has no ident - no point waiting for all results as some may not have still existing parents so don't know what level in relation to other processes
                if ((-Not $NoIndent) -and ($null -ne $Result) -and ($Result.Count -gt 1)) {
                    $LevelRange = $Result | Measure-Object -Maximum -Minimum -Property Level
                    ForEach ($Item in $Result) {
                        Add-Member -InputObject $Item -MemberType NoteProperty -Name IndentedName -Value ("$($Indenter * ($LevelRange.Maximum - $Item.Level) * $IndentMultiplier)$($Item.Name)")
                    }
                } Else {
                    # not indenting
                    $Properties[0] = 'Name'
                }
                $Result
            }
        )
    }
    END {
        $Results | Select-Object -Property $Properties # | Format-Table -AutoSize
    }
}
#endregion ProcessRelatives
#region GetCaller
# Get-Caller function to get the caller of the current command can be used to get the caller of a function or to return the right location when you throw an error
Function Get-Caller {
    <#
        .SYNOPSIS
        Get the caller of the current command.

        .DESCRIPTION
        This function captures and returns the caller of the current command using the Get-PSCallStack function and the $MyInvocation automatic variable.

        .PARAMETER None
        This function does not take any parameters.
    #>
    [CmdletBinding()]
    Param()

    # Capture the call stack
    $CallStack = Get-PSCallStack

    # Determine the caller
    If (($null -ne $CallStack) -and ($CallStack.Count -gt 1)) {
        $ParentFrame = $CallStack[1]
        $Caller = $ParentFrame.FunctionName
        $CallerScript = $ParentFrame.ScriptName
        $Line = $ParentFrame.ScriptLineNumber
    } Else {
        $Caller = $MyInvocation.InvocationName
        $CallerScript = $MyInvocation.ScriptName
        $Line = $MyInvocation.ScriptLineNumber
    }

    # If the caller script is a scriptblock, use the current script name instead
    If ($Caller.Trim() -eq '<ScriptBlock>') {
        $Caller = $CallerScript
    }

    # Return the caller information
    [PSCustomObject]@{
        Caller = $Caller
        Line   = $Line
    }
}
#endregion GetCaller
#region GetEnum
Function Get-Enum {
    <#
        .SYNOPSIS
        Get the values of an enumeration.

        .DESCRIPTION
        This function returns the values of an enumeration.

        .PARAMETER EnumType
        The enumeration type to get the values of.

        .EXAMPLE
        Get-Enum -EnumType [System.DayOfWeek]

        Get the values of the System.DayOfWeek enumeration.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, Position = 0,ValueFromPipeline)]
        ${EnumType},
        [Switch] ${Full}
    )
    # Apparently in certain condition [MyEnum] is recognized as a string and not as a type so we need to force cast it as a type
    $EnumTypeAsType = $EnumType -as [System.Type]
    if ($Full) {
        [Enum]::GetNames($EnumTypeAsType) | ForEach-Object {
            [PSCustomObject]@{
                Name = $_
                Value = $EnumType::$_.Value__
            }
        }
    }Else{
        [Enum]::GetValues($EnumTypeAsType)
    }
    # Get the values of the enumeration
    
}
#endregion GetEnum