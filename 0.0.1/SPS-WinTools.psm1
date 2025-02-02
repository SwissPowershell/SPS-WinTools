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
#region Active directory functions
# these function are based on https://github.com/techspence/ScriptSentry/blob/main/Invoke-ScriptSentry.ps1
Function Get-CurrentForest {
    <#
        .SYNOPSIS
        Get the current Active Directory Forest.

        .DESCRIPTION
        This function retrieves the current Active Directory Forest.
    #>
    [CmdletBinding()]
    param()
    Try {
        [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    } Catch [System.Management.Automation.MethodInvocationException] {
        Throw "Unable to get the current Active Directory Forest: $($_.Exception.Message)"
    }
}
Function Get-Domain {
    <#
        .SYNOPSIS
        Get the current Active Directory Domain or a specified domain.

        .DESCRIPTION
        This function retrieves the current Active Directory Domain or a specified domain.

        .PARAMETER Domain
        The domain to retrieve. If not specified, the current domain is retrieved.

        .PARAMETER Credential
        The credentials to use to retrieve the domain.

        .EXAMPLE
        Get-Domain

        Get the current Active Directory Domain.

        .EXAMPLE
        Get-Domain -Domain 'contoso.com'

        Get the domain 'contoso.com'.

        .EXAMPLE
        Get-Domain -Credential (Get-Credential)

        Get the current Active Directory Domain using alternate credentials.

        .NOTES
    #>
    [OutputType([System.DirectoryServices.ActiveDirectory.Domain])]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True)]
        [ValidateNotNullOrEmpty()]
        [String] ${Domain},
        [Parameter(Position = 1)]
        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        ${Credential} = [Management.Automation.PSCredential]::Empty
    )

    PROCESS {
        If ($PSBoundParameters['Credential']) {
            Write-Verbose '[Get-Domain] Using provided credentials for Get-Domain'
            If (-not $PSBoundParameters['Domain']) {
                # No domain provided but a credential was provided extract the credentials domain
                $Domain = $Credential.GetNetworkCredential().Domain
                Write-Verbose "[Get-Domain] Extracted domain '$($Domain)' from -Credential"
            }
            Try {
                [System.DirectoryServices.ActiveDirectory.DirectoryContext] $DomainContext = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::new('Domain', $Domain, $Credential.UserName, $Credential.GetNetworkCredential().Password)
            }Catch {
                Throw "Unable to create a DirectoryContext for domain '$($Domain)': $($_.Exception.Message)"
            }
            Try {
                [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }Catch {
                Write-Verbose "[Get-Domain] The domain '$($Domain)' does not exist, could not be contacted, there isn't an existing trust, or the specified credentials are invalid: $($_.Exception.Message)"
            }
        }ElseIf ($PSBoundParameters['Domain']) {
            Try {
                [System.DirectoryServices.ActiveDirectory.DirectoryContext] $DomainContext = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::new('Domain', $Domain)
            }Catch{
                Throw "Unable to create a DirectoryContext for the domain '$($Domain)': $($_.Exception.Message)"
            }
            Try {
                [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }Catch {
                Throw "[Get-Domain] The domain '$($Domain)' does not exist, could not be contacted, or there isn't an existing trust : $($_.Exception.Message)"
            }
        }Else {
            Try {
                [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            }Catch {
                Throw "[Get-Domain] Error retrieving the current domain: $($_.Exception.Message)"
            }
        }
    }
}
Function Get-DomainSearcher {
    <#
        .SYNOPSIS
        Get a DirectorySearcher object for a specified domain.

        .DESCRIPTION
        This function retrieves a DirectorySearcher object for a specified domain.

        .PARAMETER Domain
        The domain to retrieve. If not specified, the current domain is retrieved.

        .PARAMETER LDAPFilter
        The LDAP filter to use for the search.

        .PARAMETER Properties
        The properties to retrieve for each object.

        .PARAMETER SearchBase
        The search base to use for the search.

        .PARAMETER SearchBasePrefix
        The search base prefix to use for the search.

        .PARAMETER Server
        The domain controller to use for the search.

        .PARAMETER SearchScope
        The search scope to use for the search.

        .PARAMETER ResultPageSize
        The result page size to use for the search.

        .PARAMETER ServerTimeLimit
        The server time limit to use for the search.

        .PARAMETER SecurityMasks
        The security masks to use for the search.

        .PARAMETER Tombstone
        The tombstone to use for the search.

        .PARAMETER Credential
        The credentials to use to retrieve the domain.

        .EXAMPLE
        Get-DomainSearcher -Domain 'contoso.com' -LDAPFilter '(objectClass=user)' -Properties @('name', 'distinguishedName') -SearchBase 'DC=contoso,DC=com' -SearchBasePrefix 'CN=Users' -Server 'dc01.contoso.com' -SearchScope 'Subtree' -ResultPageSize 200 -ServerTimeLimit 120 -SecurityMasks 'Dacl' -Tombstone -Credential (Get-Credential)

        Get a DirectorySearcher object for the domain 'contoso.com' with the specified parameters.

        .NOTES

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('System.DirectoryServices.DirectorySearcher')]
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline = $True,Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String] ${Domain},

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String] ${LDAPFilter},

        [ValidateNotNullOrEmpty()]
        [String[]] ${Properties},

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String] ${SearchBase},

        [ValidateNotNullOrEmpty()]
        [String] ${SearchBasePrefix},

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String] ${Server},

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String] ${SearchScope} = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int] ${ResultPageSize} = 200,

        [ValidateRange(1, 10000)]
        [Int] ${ServerTimeLimit} = 120,

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String] ${SecurityMasks},

        [Switch] ${Tombstone},

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        ${Credential} = [Management.Automation.PSCredential]::Empty
    )

    PROCESS {
        If ($PSBoundParameters['Domain']) {
            If ($ENV:USERDNSDOMAIN -and ($ENV:USERDNSDOMAIN.Trim() -ne '')) {
                # see if we can grab the user DNS logon domain from environment variables
                $UserDomain = $ENV:USERDNSDOMAIN
                If ($ENV:LOGONSERVER -and ($ENV:LOGONSERVER.Trim() -ne '') -and $UserDomain) {
                    $Server = "$($ENV:LOGONSERVER -replace '\\','').$UserDomain"
                }
            }
        }ElseIf ($PSBoundParameters['Credential']) {
            # if not -Domain is specified, but -Credential is, try to retrieve the current domain name with Get-Domain
            $DomainObject = Get-Domain -Credential $Credential
            $Server = ($DomainObject.PdcRoleOwner).Name
            $Domain = $DomainObject.Name
        }ElseIf ($ENV:USERDNSDOMAIN -and ($ENV:USERDNSDOMAIN.Trim() -ne '')) {
            # see if we can grab the user DNS logon domain from environment variables
            $Domain = $ENV:USERDNSDOMAIN
            If ($ENV:LOGONSERVER -and ($ENV:LOGONSERVER.Trim() -ne '') -and $Domain) {
                $Server = "$($ENV:LOGONSERVER -replace '\\','').$Domain"
            }
        }Else {
            # otherwise, resort to Get-Domain to retrieve the current domain object
            write-verbose "Get-Domain"
            $DomainObject = Get-Domain
            $Server = ($DomainObject.PdcRoleOwner).Name
            $Domain = $DomainObject.Name
        }
        [System.Collections.Generic.List[String]] $SearchStringArray = @('LDAP://')
        If ($Server -and ($Server.Trim() -ne '')) {
            $SearchStringArray.Add($Server)
            If ($Domain) {
                $SearchStringArray.Add('/')
            }
        }

        If ($PSBoundParameters['SearchBasePrefix']) {
            $SearchStringArray.Add("$($SearchBasePrefix),")
        }

        If ($PSBoundParameters['SearchBase']) {
            If ($SearchBase -Match '^GC://') {
                # if we're searching the global catalog, get the path in the right format
                $DN = $SearchBase.ToUpper().Trim('/')
                $SearchStringArray = @()
            }Else {
                If ($SearchBase -match '^LDAP://') {
                    If ($SearchBase -match "LDAP://.+/.+") {
                        $SearchStringArray = @()
                        $DN = $SearchBase
                    }Else {
                        $DN = $SearchBase.SubString(7)
                    }
                }Else {
                    $DN = $SearchBase
                }
            }
        }Else {
            # transform the target domain name into a distinguishedName if an ADS search base is not specified
            If ($Domain -and ($Domain.Trim() -ne '')) {
                $DN = "DC=$($Domain.Replace('.', ',DC='))"
            }
        }

        $SearchStringArray.Add($DN)
        # Convert the collection of string to an array of string
        [String[]] $SearchString = $SearchStringArray
        Write-Verbose "[Get-DomainSearcher] search base: $($SearchString)"

        If ($Credential -ne [Management.Automation.PSCredential]::Empty) {
            Write-Verbose "[Get-DomainSearcher] Using alternate credentials for LDAP connection"
            # bind to the inital search object using alternate credentials
            [DirectoryServices.DirectoryEntry] $DomainObject = [DirectoryServices.DirectoryEntry]::New($SearchString, $Credential.UserName, $Credential.GetNetworkCredential().Password)
            [System.DirectoryServices.DirectorySearcher] $Searcher = [System.DirectoryServices.DirectorySearcher]::new($DomainObject)
        }Else {
            # bind to the inital object using the current credentials
            [System.DirectoryServices.DirectorySearcher] $Searcher = [System.DirectoryServices.DirectorySearcher]::new([ADSI]$SearchString)
        }

        $Searcher.PageSize = $ResultPageSize
        $Searcher.SearchScope = $SearchScope
        $Searcher.CacheResults = $False
        $Searcher.ReferralChasing = [System.DirectoryServices.ReferralChasingOption]::All

        if ($PSBoundParameters['ServerTimeLimit']) {
            $Searcher.ServerTimeLimit = $ServerTimeLimit
        }

        if ($PSBoundParameters['Tombstone']) {
            $Searcher.Tombstone = $True
        }

        if ($PSBoundParameters['LDAPFilter']) {
            $Searcher.Filter = $LDAPFilter
        }

        if ($PSBoundParameters['SecurityMasks']) {
            $Searcher.SecurityMasks = Switch ($SecurityMasks) {
                'Dacl' { [System.DirectoryServices.SecurityMasks]::Dacl ; BREAK}
                'Group' { [System.DirectoryServices.SecurityMasks]::Group ; BREAK}
                'None' { [System.DirectoryServices.SecurityMasks]::None ; BREAK}
                'Owner' { [System.DirectoryServices.SecurityMasks]::Owner ; BREAK}
                'Sacl' { [System.DirectoryServices.SecurityMasks]::Sacl ; BREAK}
                default {} # no security mask
            }
        }

        if ($PSBoundParameters['Properties']) {
            # handle an array of properties to load w/ the possibility of comma-separated strings
            $PropertiesToLoad = $Properties | ForEach-Object { $_.Split(',') }
            $Null = $Searcher.PropertiesToLoad.AddRange(($PropertiesToLoad))
        }
        Write-Output $Searcher
    }
}
Function Get-DomainGroupMember {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [OutputType('StrongView.GroupMember')]
    [CmdletBinding(DefaultParameterSetName = 'None')]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name', 'MemberDistinguishedName', 'MemberName')]
        [String[]] ${Identity},

        [ValidateNotNullOrEmpty()]
        [String] ${Domain},

        [Parameter(ParameterSetName = 'ManualRecurse')]
        [Switch] ${Recurse},

        [Parameter(ParameterSetName = 'RecurseUsingMatchingRule')]
        [Switch] ${RecurseUsingMatchingRule},

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String] ${LDAPFilter},

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String] ${SearchBase},

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String] ${Server},

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String] ${SearchScope} = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int] ${ResultPageSize} = 200,

        [ValidateRange(1, 10000)]
        [Int] ${ServerTimeLimit},

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String] ${SecurityMasks},

        [Switch] ${Tombstone},

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        ${Credential} = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $SearcherArguments = @{
            'Properties' = 'member,samaccountname,distinguishedname'
        }
        if ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        if ($PSBoundParameters['LDAPFilter']) { $SearcherArguments['LDAPFilter'] = $LDAPFilter }
        if ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        if ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        if ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        if ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        if ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        if ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        if ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }

        $ADNameArguments = @{}
        if ($PSBoundParameters['Domain']) { $ADNameArguments['Domain'] = $Domain }
        if ($PSBoundParameters['Server']) { $ADNameArguments['Server'] = $Server }
        if ($PSBoundParameters['Credential']) { $ADNameArguments['Credential'] = $Credential }
    }

    PROCESS {
        $GroupSearcher = Get-DomainSearcher @SearcherArguments
        If ($GroupSearcher) {
            If ($PSBoundParameters['RecurseUsingMatchingRule']) {
                $SearcherArguments['Identity'] = $Identity
                $SearcherArguments['Raw'] = $True
                $Group = Get-DomainGroup @SearcherArguments

                If (-not $Group) {
                    Write-Warning "[Get-DomainGroupMember] Error searching for group with identity: $($Identity)"
                }Else {
                    $GroupFoundName = $Group.properties.item('samaccountname')[0]
                    $GroupFoundDN = $Group.properties.item('distinguishedname')[0]

                    If ($PSBoundParameters['Domain']) {
                        $GroupFoundDomain = $Domain
                    }Else {
                        # if a domain isn't passed, try to extract it from the found group distinguished name
                        If ($GroupFoundDN) {
                            $GroupFoundDomain = $GroupFoundDN.SubString($GroupFoundDN.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        }
                    }
                    Write-Verbose "[Get-DomainGroupMember] Using LDAP matching rule to recurse on '$($GroupFoundDN)', only user accounts will be returned."
                    $GroupSearcher.filter = "(&(samAccountType=805306368)(memberof:1.2.840.113556.1.4.1941:=$($GroupFoundDN)))"
                    $GroupSearcher.PropertiesToLoad.AddRange(('distinguishedName'))
                    $Members = $GroupSearcher.FindAll() | ForEach-Object {$_.Properties.distinguishedname[0]}
                }
                $Null = $SearcherArguments.Remove('Raw')
            }Else {
                $IdentityFilter = ''
                $Filter = ''
                $Identity | Where-Object {$_} | ForEach-Object {
                    $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                    If ($IdentityInstance -match '^S-1-') {
                        $IdentityFilter += "(objectsid=$($IdentityInstance))"
                    }Elseif ($IdentityInstance -match '^CN=') {
                        $IdentityFilter += "(distinguishedname=$($IdentityInstance))"
                        If ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                            # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                            #   and rebuild the domain searcher
                            $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                            Write-Verbose "[Get-DomainGroupMember] Extracted domain '$($IdentityDomain)' from '$($IdentityInstance)'"
                            $SearcherArguments['Domain'] = $IdentityDomain
                            $GroupSearcher = Get-DomainSearcher @SearcherArguments
                            If (-not $GroupSearcher) {
                                Write-Warning "[Get-DomainGroupMember] Unable to retrieve domain searcher for '$($IdentityDomain)'"
                            }
                        }
                    }ElseIf ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                        $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                        $IdentityFilter += "(objectguid=$($GuidByteString))"
                    }ElseIf ($IdentityInstance.Contains('\')) {
                        $ConvertedIdentityInstance = $IdentityInstance.Replace('\28', '(').Replace('\29', ')') | Convert-ADName -OutputType Canonical
                        If ($ConvertedIdentityInstance) {
                            $GroupDomain = $ConvertedIdentityInstance.SubString(0, $ConvertedIdentityInstance.IndexOf('/'))
                            $GroupName = $IdentityInstance.Split('\')[1]
                            $IdentityFilter += "(samAccountName=$($GroupName))"
                            $SearcherArguments['Domain'] = $GroupDomain
                            Write-Verbose "[Get-DomainGroupMember] Extracted domain '$($GroupDomain)' from '$($IdentityInstance)'"
                            $GroupSearcher = Get-DomainSearcher @SearcherArguments
                        }
                    }Else {
                        $IdentityFilter += "(samAccountName=$($IdentityInstance))"
                    }
                }

                If ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                    $Filter += "(|$($IdentityFilter))"
                }

                If ($PSBoundParameters['LDAPFilter']) {
                    Write-Verbose "[Get-DomainGroupMember] Using additional LDAP filter: $($LDAPFilter)"
                    $Filter += "$($LDAPFilter)"
                }

                $GroupSearcher.Filter = "(&(objectCategory=group)$($Filter))"
                Write-Verbose "[Get-DomainGroupMember] Get-DomainGroupMember filter string: $($GroupSearcher.filter)"
                Try {
                    $Result = $GroupSearcher.FindOne()
                }Catch {
                    Write-Warning "[Get-DomainGroupMember] Error searching for group with identity '$($Identity)': $($_.Exception.Message)"
                    $Members = @()
                }

                $GroupFoundName = ''
                $GroupFoundDN = ''

                if ($Result) {
                    $Members = $Result.properties.item('member')

                    if ($Members.count -eq 0) {
                        # ranged searching, thanks @meatballs__ !
                        $Finished = $False
                        $Bottom = 0
                        $Top = 0

                        While (-not $Finished) {
                            $Top = $Bottom + 1499
                            $MemberRange="member;range=$($Bottom)-$($Top)"
                            $Bottom += 1500
                            $Null = $GroupSearcher.PropertiesToLoad.Clear()
                            $Null = $GroupSearcher.PropertiesToLoad.Add("$($MemberRange)")
                            $Null = $GroupSearcher.PropertiesToLoad.Add('samaccountname')
                            $Null = $GroupSearcher.PropertiesToLoad.Add('distinguishedname')

                            Try {
                                $Result = $GroupSearcher.FindOne()
                                $RangedProperty = $Result.Properties.PropertyNames -like "member;range=*"
                                $Members += $Result.Properties.item($RangedProperty)
                                $GroupFoundName = $Result.properties.item('samaccountname')[0]
                                $GroupFoundDN = $Result.properties.item('distinguishedname')[0]

                                If ($Members.count -eq 0) {
                                    $Finished = $True
                                }
                            }Catch [System.Management.Automation.MethodInvocationException] {
                                $Finished = $True
                            }
                        }
                    }Else {
                        $GroupFoundName = $Result.properties.item('samaccountname')[0]
                        $GroupFoundDN = $Result.properties.item('distinguishedname')[0]
                        $Members += $Result.Properties.item($RangedProperty)
                    }

                    If ($PSBoundParameters['Domain']) {
                        $GroupFoundDomain = $Domain
                    }Else {
                        # if a domain isn't passed, try to extract it from the found group distinguished name
                        If ($GroupFoundDN) {
                            $GroupFoundDomain = $GroupFoundDN.SubString($GroupFoundDN.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        }
                    }
                }
            }

            ForEach ($Member in $Members) {
                If ($Recurse -and $UseMatchingRule) {
                    $Properties = $_.Properties
                }Else {
                    $ObjectSearcherArguments = $SearcherArguments.Clone()
                    $ObjectSearcherArguments['Identity'] = $Member
                    $ObjectSearcherArguments['Raw'] = $True
                    $ObjectSearcherArguments['Properties'] = 'distinguishedname,cn,samaccountname,objectsid,objectclass'
                    $Object = Get-DomainObject @ObjectSearcherArguments
                    $Properties = $Object.Properties
                }

                If ($Properties) {
                    $GroupMember = New-Object PSObject
                    $GroupMember | Add-Member Noteproperty 'GroupDomain' $GroupFoundDomain
                    $GroupMember | Add-Member Noteproperty 'GroupName' $GroupFoundName
                    $GroupMember | Add-Member Noteproperty 'GroupDistinguishedName' $GroupFoundDN

                    If ($Properties.objectsid) {
                        $MemberSID = (([System.Security.Principal.SecurityIdentifier]::New($Properties.objectsid[0], 0).Value))
                    }Else {
                        $MemberSID = $Null
                    }

                    Try {
                        $MemberDN = $Properties.distinguishedname[0]
                        If ($MemberDN -match 'ForeignSecurityPrincipals|S-1-5-21') {
                            Try {
                                If (-not $MemberSID) {
                                    $MemberSID = $Properties.cn[0]
                                }
                                $MemberSimpleName = Convert-ADName -Identity $MemberSID -OutputType 'DomainSimple' @ADNameArguments

                                If ($MemberSimpleName) {
                                    $MemberDomain = $MemberSimpleName.Split('@')[1]
                                }Else {
                                    Write-Warning "[Get-DomainGroupMember] Error converting $MemberDN"
                                    $MemberDomain = $Null
                                }
                            }Catch {
                                Write-Warning "[Get-DomainGroupMember] Error converting $MemberDN"
                                $MemberDomain = $Null
                            }
                        }Else {
                            # extract the FQDN from the Distinguished Name
                            $MemberDomain = $MemberDN.SubString($MemberDN.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        }
                    }Catch {
                        $MemberDN = $Null
                        $MemberDomain = $Null
                    }

                    If ($Properties.samaccountname) {
                        # forest users have the samAccountName set
                        $MemberName = $Properties.samaccountname[0]
                    }Else {
                        # external trust users have a SID, so convert it
                        Try {
                            $MemberName = ConvertFrom-SID -ObjectSID $Properties.cn[0] @ADNameArguments
                        }Catch {
                            # if there's a problem contacting the domain to resolve the SID
                            $MemberName = $Properties.cn[0]
                        }
                    }

                    If ($Properties.objectclass -match 'computer') {
                        $MemberObjectClass = 'computer'
                    }ElseIf ($Properties.objectclass -match 'group') {
                        $MemberObjectClass = 'group'
                    }ElseIf ($Properties.objectclass -match 'user') {
                        $MemberObjectClass = 'user'
                    }Else {
                        $MemberObjectClass = $Null
                    }
                    $GroupMember | Add-Member Noteproperty 'MemberDomain' $MemberDomain
                    $GroupMember | Add-Member Noteproperty 'MemberName' $MemberName
                    $GroupMember | Add-Member Noteproperty 'MemberDistinguishedName' $MemberDN
                    $GroupMember | Add-Member Noteproperty 'MemberObjectClass' $MemberObjectClass
                    $GroupMember | Add-Member Noteproperty 'MemberSID' $MemberSID
                    $GroupMember.PSObject.TypeNames.Insert(0, 'StrongView.GroupMember')
                    Write-Output $GroupMember

                    # if we're doing manual recursion
                    If ($PSBoundParameters['Recurse'] -and $MemberDN -and ($MemberObjectClass -match 'group')) {
                        Write-Verbose "[Get-DomainGroupMember] Manually recursing on group: $($MemberDN)"
                        $SearcherArguments['Identity'] = $MemberDN
                        $Null = $SearcherArguments.Remove('Properties')
                        Get-DomainGroupMember @SearcherArguments
                    }
                }
            }
            $GroupSearcher.dispose()
        }
    }
}

Function Get-DomainUser {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '')]
    [OutputType('StrongView.User')]
    [OutputType('StrongView.User.Raw')]
    [CmdletBinding(DefaultParameterSetName = 'AllowDelegation')]
    Param(
            [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
            [Alias('DistinguishedName', 'SamAccountName', 'Name', 'MemberDistinguishedName', 'MemberName')]
            [String[]] ${Identity},
    
            [Switch] ${SPN},
    
            [Switch] ${AdminCount},
    
            [Parameter(ParameterSetName = 'AllowDelegation')]
            [Switch] ${AllowDelegation},
    
            [Parameter(ParameterSetName = 'DisallowDelegation')]
            [Switch] ${DisallowDelegation},
    
            [Switch] ${TrustedToAuth},
    
            [Alias('KerberosPreauthNotRequired', 'NoPreauth')]
            [Switch] ${PreauthNotRequired},
    
            [ValidateNotNullOrEmpty()]
            [String] ${Domain},
    
            [ValidateNotNullOrEmpty()]
            [Alias('Filter')]
            [String] ${LDAPFilter},
    
            [ValidateNotNullOrEmpty()]
            [String[]] ${Properties},
    
            [ValidateNotNullOrEmpty()]
            [Alias('ADSPath')]
            [String] ${SearchBase},
    
            [ValidateNotNullOrEmpty()]
            [Alias('DomainController')]
            [String] ${Server},
    
            [ValidateSet('Base', 'OneLevel', 'Subtree')]
            [String] ${SearchScope} = 'Subtree',
    
            [ValidateRange(1, 10000)]
            [Int] ${ResultPageSize} = 200,
    
            [ValidateRange(1, 10000)]
            [Int] ${ServerTimeLimit},
    
            [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
            [String] ${SecurityMasks},
    
            [Switch] ${Tombstone},
    
            [Alias('ReturnOne')]
            [Switch] ${FindOne},
    
            [Management.Automation.PSCredential]
            [Management.Automation.CredentialAttribute()]
            ${Credential} = [Management.Automation.PSCredential]::Empty,
    
            [Switch] ${Raw}
        )
    BEGIN {
        $SearcherArguments = @{}
        If ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        If ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        If ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        If ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        If ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        If ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        If ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        If ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        If ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        If ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $UserSearcher = Get-DomainSearcher @SearcherArguments
    }
    
    PROCESS {
        If ($UserSearcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                If ($IdentityInstance -match '^S-1-') {
                    $IdentityFilter += "(objectsid=$($IdentityInstance))"
                }ElseIf ($IdentityInstance -match '^CN=') {
                    $IdentityFilter += "(distinguishedname=$($IdentityInstance))"
                    If ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainUser] Extracted domain '$($IdentityDomain)' from '$($IdentityInstance)'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $UserSearcher = Get-DomainSearcher @SearcherArguments
                        If (-not $UserSearcher) {
                            Write-Warning "[Get-DomainUser] Unable to retrieve domain searcher for '$($IdentityDomain)'"
                        }
                    }
                }ElseIf ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                    $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                    $IdentityFilter += "(objectguid=$($GuidByteString))"
                }ElseIf ($IdentityInstance.Contains('\')) {
                    $ConvertedIdentityInstance = $IdentityInstance.Replace('\28', '(').Replace('\29', ')') | Convert-ADName -OutputType Canonical
                    If ($ConvertedIdentityInstance) {
                        $UserDomain = $ConvertedIdentityInstance.SubString(0, $ConvertedIdentityInstance.IndexOf('/'))
                        $UserName = $IdentityInstance.Split('\')[1]
                        $IdentityFilter += "(samAccountName=$($UserName))"
                        $SearcherArguments['Domain'] = $UserDomain
                        Write-Verbose "[Get-DomainUser] Extracted domain '$($UserDomain)' from '$($IdentityInstance)'"
                        $UserSearcher = Get-DomainSearcher @SearcherArguments
                    }
                }Else {
                    $IdentityFilter += "(samAccountName=$($IdentityInstance))"
                }
            }

            If ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$($IdentityFilter))"
            }

            If ($PSBoundParameters['SPN']) {
                Write-Verbose '[Get-DomainUser] Searching for non-null service principal names'
                $Filter += '(servicePrincipalName=*)'
            }
            If ($PSBoundParameters['AllowDelegation']) {
                Write-Verbose '[Get-DomainUser] Searching for users who can be delegated'
                # negation of "Accounts that are sensitive and not trusted for delegation"
                $Filter += '(!(userAccountControl:1.2.840.113556.1.4.803:=1048574))'
            }
            If ($PSBoundParameters['DisallowDelegation']) {
                Write-Verbose '[Get-DomainUser] Searching for users who are sensitive and not trusted for delegation'
                $Filter += '(userAccountControl:1.2.840.113556.1.4.803:=1048574)'
            }
            If ($PSBoundParameters['AdminCount']) {
                Write-Verbose '[Get-DomainUser] Searching for adminCount=1'
                $Filter += '(admincount=1)'
            }
            If ($PSBoundParameters['TrustedToAuth']) {
                Write-Verbose '[Get-DomainUser] Searching for users that are trusted to authenticate for other principals'
                $Filter += '(msds-allowedtodelegateto=*)'
            }
            If ($PSBoundParameters['PreauthNotRequired']) {
                Write-Verbose '[Get-DomainUser] Searching for user accounts that do not require kerberos preauthenticate'
                $Filter += '(userAccountControl:1.2.840.113556.1.4.803:=4194304)'
            }
            If ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainUser] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$($LDAPFilter)"
            }

            # build the LDAP filter for the dynamic UAC filter value
            $UACFilter | Where-Object {$_} | ForEach-Object {
                If ($_ -match 'NOT_.*') {
                    $UACField = $_.Substring(4)
                    $UACValue = [Int]($UACEnum::$UACField)
                    $Filter += "(!(userAccountControl:1.2.840.113556.1.4.803:=$UACValue))"
                }Else {
                    $UACValue = [Int]($UACEnum::$_)
                    $Filter += "(userAccountControl:1.2.840.113556.1.4.803:=$UACValue)"
                }
            }

            $UserSearcher.filter = "(&(samAccountType=805306368)$Filter)"
            Write-Verbose "[Get-DomainUser] filter string: $($UserSearcher.filter)"

            If ($PSBoundParameters['FindOne']) { $Results = $UserSearcher.FindOne() }Else { $Results = $UserSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                If ($PSBoundParameters['Raw']) {
                    # return raw result objects
                    $User = $_
                    $User.PSObject.TypeNames.Insert(0, 'StrongView.User.Raw')
                }Else {
                    $User = Convert-LDAPProperty -Properties $_.Properties
                    $User.PSObject.TypeNames.Insert(0, 'StrongView.User')
                }
                $User
            }
            If ($Results) {
                Try { $Results.dispose() }Catch { Write-Verbose "[Get-DomainUser] Error disposing of the Results object: $($_.Exception.Message)" }
            }
            $UserSearcher.dispose()
        }
    }
}
## TO DO : Refactor nexts functions
Function Get-DomainObject {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [OutputType('StrongView.ADObject')]
    [OutputType('StrongView.ADObject.Raw')]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('DistinguishedName', 'SamAccountName', 'Name', 'MemberDistinguishedName', 'MemberName')]
        [String[]] ${Identity},

        [ValidateNotNullOrEmpty()]
        [String] ${Domain},

        [ValidateNotNullOrEmpty()]
        [Alias('Filter')]
        [String] ${LDAPFilter},

        [ValidateNotNullOrEmpty()]
        [String[]] ${Properties},

        [ValidateNotNullOrEmpty()]
        [Alias('ADSPath')]
        [String] ${SearchBase},

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String] ${Server},

        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [String] ${SearchScope} = 'Subtree',

        [ValidateRange(1, 10000)]
        [Int] ${ResultPageSize} = 200,

        [ValidateRange(1, 10000)]
        [Int] ${ServerTimeLimit},

        [ValidateSet('Dacl', 'Group', 'None', 'Owner', 'Sacl')]
        [String] ${SecurityMasks},

        [Switch] ${Tombstone},

        [Alias('ReturnOne')]
        [Switch] ${FindOne},

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        ${Credential} = [Management.Automation.PSCredential]::Empty,

        [Switch] ${Raw}
    )

    BEGIN {
        $SearcherArguments = @{}
        If ($PSBoundParameters['Domain']) { $SearcherArguments['Domain'] = $Domain }
        If ($PSBoundParameters['Properties']) { $SearcherArguments['Properties'] = $Properties }
        If ($PSBoundParameters['SearchBase']) { $SearcherArguments['SearchBase'] = $SearchBase }
        If ($PSBoundParameters['Server']) { $SearcherArguments['Server'] = $Server }
        If ($PSBoundParameters['SearchScope']) { $SearcherArguments['SearchScope'] = $SearchScope }
        If ($PSBoundParameters['ResultPageSize']) { $SearcherArguments['ResultPageSize'] = $ResultPageSize }
        If ($PSBoundParameters['ServerTimeLimit']) { $SearcherArguments['ServerTimeLimit'] = $ServerTimeLimit }
        If ($PSBoundParameters['SecurityMasks']) { $SearcherArguments['SecurityMasks'] = $SecurityMasks }
        If ($PSBoundParameters['Tombstone']) { $SearcherArguments['Tombstone'] = $Tombstone }
        If ($PSBoundParameters['Credential']) { $SearcherArguments['Credential'] = $Credential }
        $ObjectSearcher = Get-DomainSearcher @SearcherArguments
    }

    PROCESS {
        If ($ObjectSearcher) {
            $IdentityFilter = ''
            $Filter = ''
            $Identity | Where-Object {$_} | ForEach-Object {
                $IdentityInstance = $_.Replace('(', '\28').Replace(')', '\29')
                If ($IdentityInstance -match '^S-1-') {
                    $IdentityFilter += "(objectsid=$($IdentityInstance))"
                }Elseif ($IdentityInstance -match '^(CN|OU|DC)=') {
                    $IdentityFilter += "(distinguishedname=$($IdentityInstance))"
                    If ((-not $PSBoundParameters['Domain']) -and (-not $PSBoundParameters['SearchBase'])) {
                        # if a -Domain isn't explicitly set, extract the object domain out of the distinguishedname
                        #   and rebuild the domain searcher
                        $IdentityDomain = $IdentityInstance.SubString($IdentityInstance.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
                        Write-Verbose "[Get-DomainObject] Extracted domain '$($IdentityDomain)' from '$($IdentityInstance)'"
                        $SearcherArguments['Domain'] = $IdentityDomain
                        $ObjectSearcher = Get-DomainSearcher @SearcherArguments
                        If (-not $ObjectSearcher) {
                            Write-Warning "[Get-DomainObject] Unable to retrieve domain searcher for '$($IdentityDomain)'"
                        }
                    }
                }Elseif ($IdentityInstance -imatch '^[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}$') {
                    $GuidByteString = (([Guid]$IdentityInstance).ToByteArray() | ForEach-Object { '\' + $_.ToString('X2') }) -join ''
                    $IdentityFilter += "(objectguid=$($GuidByteString))"
                }Elseif ($IdentityInstance.Contains('\')) {
                    $ConvertedIdentityInstance = $IdentityInstance.Replace('\28', '(').Replace('\29', ')') | Convert-ADName -OutputType Canonical
                    If ($ConvertedIdentityInstance) {
                        $ObjectDomain = $ConvertedIdentityInstance.SubString(0, $ConvertedIdentityInstance.IndexOf('/'))
                        $ObjectName = $IdentityInstance.Split('\')[1]
                        $IdentityFilter += "(samAccountName=$ObjectName)"
                        $SearcherArguments['Domain'] = $ObjectDomain
                        Write-Verbose "[Get-DomainObject] Extracted domain '$($ObjectDomain)' from '$IdentityInstance'"
                        $ObjectSearcher = Get-DomainSearcher @SearcherArguments
                    }
                }Elseif ($IdentityInstance.Contains('.')) {
                    $IdentityFilter += "(|(samAccountName=$IdentityInstance)(name=$IdentityInstance)(dnshostname=$IdentityInstance))"
                }Else {
                    $IdentityFilter += "(|(samAccountName=$IdentityInstance)(name=$IdentityInstance)(displayname=$IdentityInstance))"
                }
            }
            If ($IdentityFilter -and ($IdentityFilter.Trim() -ne '') ) {
                $Filter += "(|$IdentityFilter)"
            }

            If ($PSBoundParameters['LDAPFilter']) {
                Write-Verbose "[Get-DomainObject] Using additional LDAP filter: $LDAPFilter"
                $Filter += "$LDAPFilter"
            }

            # build the LDAP filter for the dynamic UAC filter value
            $UACFilter | Where-Object {$_} | ForEach-Object {
                If ($_ -match 'NOT_.*') {
                    $UACField = $_.Substring(4)
                    $UACValue = [Int]($UACEnum::$UACField)
                    $Filter += "(!(userAccountControl:1.2.840.113556.1.4.803:=$UACValue))"
                }Else {
                    $UACValue = [Int]($UACEnum::$_)
                    $Filter += "(userAccountControl:1.2.840.113556.1.4.803:=$UACValue)"
                }
            }

            If ($Filter -and $Filter -ne '') {
                $ObjectSearcher.filter = "(&$Filter)"
            }
            Write-Verbose "[Get-DomainObject] Get-DomainObject filter string: $($ObjectSearcher.filter)"

            If ($PSBoundParameters['FindOne']) { $Results = $ObjectSearcher.FindOne() }Else { $Results = $ObjectSearcher.FindAll() }
            $Results | Where-Object {$_} | ForEach-Object {
                If ($PSBoundParameters['Raw']) {
                    # return raw result objects
                    $Object = $_
                    $Object.PSObject.TypeNames.Insert(0, 'StrongView.ADObject.Raw')
                }Else {
                    $Object = Convert-LDAPProperty -Properties $_.Properties
                    $Object.PSObject.TypeNames.Insert(0, 'StrongView.ADObject')
                }
                $Object
            }
            If ($Results) {
                Try { $Results.dispose() }Catch { Write-Verbose "[Get-DomainObject] Error disposing of the Results object: $($_.Exception.Message)" }
            }
            $ObjectSearcher.dispose()
        }
    }
}

Function Convert-ADName {
    <#
    .SYNOPSIS
    
    Converts Active Directory object names between a variety of formats.
    
    Author: Bill Stewart, Pasquale Lantella  
    Modifications: Will Schroeder (@harmj0y)  
    License: BSD 3-Clause  
    Required Dependencies: None  
    
    .DESCRIPTION
    
    This function is heavily based on Bill Stewart's code and Pasquale Lantella's code (in LINK)
    and translates Active Directory names between various formats using the NameTranslate COM object.
    
    .PARAMETER Identity
    
    Specifies the Active Directory object name to translate, of the following form:
    
        DN                short for 'distinguished name'; e.g., 'CN=Phineas Flynn,OU=Engineers,DC=fabrikam,DC=com'
        Canonical         canonical name; e.g., 'fabrikam.com/Engineers/Phineas Flynn'
        NT4               domain\username; e.g., 'fabrikam\pflynn'
        Display           display name, e.g. 'pflynn'
        DomainSimple      simple domain name format, e.g. 'pflynn@fabrikam.com'
        EnterpriseSimple  simple enterprise name format, e.g. 'pflynn@fabrikam.com'
        GUID              GUID; e.g., '{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}'
        UPN               user principal name; e.g., 'pflynn@fabrikam.com'
        CanonicalEx       extended canonical name format
        SPN               service principal name format; e.g. 'HTTP/kairomac.contoso.com'
        SID               Security Identifier; e.g., 'S-1-5-21-12986231-600641547-709122288-57999'
    
    .PARAMETER OutputType
    
    Specifies the output name type you want to convert to, which must be one of the following:
    
        DN                short for 'distinguished name'; e.g., 'CN=Phineas Flynn,OU=Engineers,DC=fabrikam,DC=com'
        Canonical         canonical name; e.g., 'fabrikam.com/Engineers/Phineas Flynn'
        NT4               domain\username; e.g., 'fabrikam\pflynn'
        Display           display name, e.g. 'pflynn'
        DomainSimple      simple domain name format, e.g. 'pflynn@fabrikam.com'
        EnterpriseSimple  simple enterprise name format, e.g. 'pflynn@fabrikam.com'
        GUID              GUID; e.g., '{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}'
        UPN               user principal name; e.g., 'pflynn@fabrikam.com'
        CanonicalEx       extended canonical name format, e.g. 'fabrikam.com/Users/Phineas Flynn'
        SPN               service principal name format; e.g. 'HTTP/kairomac.contoso.com'
    
    .PARAMETER Domain
    
    Specifies the domain to use for the translation, defaults to the current domain.
    
    .PARAMETER Server
    
    Specifies an Active Directory server (domain controller) to bind to for the translation.
    
    .PARAMETER Credential
    
    Specifies an alternate credential to use for the translation.
    
    .EXAMPLE
    
    Convert-ADName -Identity "TESTLAB\harmj0y"
    
    harmj0y@testlab.local
    
    .EXAMPLE
    
    "TESTLAB\krbtgt", "CN=Administrator,CN=Users,DC=testlab,DC=local" | Convert-ADName -OutputType Canonical
    
    testlab.local/Users/krbtgt
    testlab.local/Users/Administrator
    
    .EXAMPLE
    
    Convert-ADName -OutputType dn -Identity 'TESTLAB\harmj0y' -Server PRIMARY.testlab.local
    
    CN=harmj0y,CN=Users,DC=testlab,DC=local
    
    .EXAMPLE
    
    $SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
    $Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm', $SecPassword)
    'S-1-5-21-890171859-3433809279-3366196753-1108' | Convert-ADNAme -Credential $Cred
    
    TESTLAB\harmj0y
    
    .INPUTS
    
    String
    
    Accepts one or more objects name strings on the pipeline.
    
    .OUTPUTS
    
    String
    
    Outputs a string representing the converted name.
    
    .LINK
    
    http://windowsitpro.com/active-directory/translating-active-directory-object-names-between-formats
    https://gallery.technet.microsoft.com/scriptcenter/Translating-Active-5c80dd67
    #>
    
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [OutputType([String])]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('Name', 'ObjectName')]
        [String[]] ${Identity},

        [ValidateSet('DN', 'Canonical', 'NT4', 'Display', 'DomainSimple', 'EnterpriseSimple', 'GUID', 'Unknown', 'UPN', 'CanonicalEx', 'SPN')]
        [String] ${OutputType},

        [ValidateNotNullOrEmpty()]
        [String] ${Domain},

        [ValidateNotNullOrEmpty()]
        [Alias('DomainController')]
        [String] ${Server},

        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        ${Credential} = [Management.Automation.PSCredential]::Empty
    )

    BEGIN {
        $NameTypes = @{
            'DN'                =   1  # CN=Phineas Flynn,OU=Engineers,DC=fabrikam,DC=com
            'Canonical'         =   2  # fabrikam.com/Engineers/Phineas Flynn
            'NT4'               =   3  # fabrikam\pflynn
            'Display'           =   4  # pflynn
            'DomainSimple'      =   5  # pflynn@fabrikam.com
            'EnterpriseSimple'  =   6  # pflynn@fabrikam.com
            'GUID'              =   7  # {95ee9fff-3436-11d1-b2b0-d15ae3ac8436}
            'Unknown'           =   8  # unknown type - let the server do translation
            'UPN'               =   9  # pflynn@fabrikam.com
            'CanonicalEx'       =   10 # fabrikam.com/Users/Phineas Flynn
            'SPN'               =   11 # HTTP/kairomac.contoso.com
            'SID'               =   12 # S-1-5-21-12986231-600641547-709122288-57999
        }

        # accessor functions from Bill Stewart to simplify calls to NameTranslate
        Function Invoke-Method {
            Param (
                [__ComObject] $Object,
                [String] $Method,
                $Parameters
            )
            $Output = $Null
            $Output = $Object.GetType().InvokeMember($Method, 'InvokeMethod', $NULL, $Object, $Parameters)
            Write-Output $Output
        }

        Function Get-Property {
            Param(
                [__ComObject] $Object,
                [String] $Property
            )
            $Object.GetType().InvokeMember($Property, 'GetProperty', $NULL, $Object, $NULL)
        }

        Function Set-Property {
            Param(
                [__ComObject] $Object,
                [String] $Property,
                $Parameters
            )
            [Void] $Object.GetType().InvokeMember($Property, 'SetProperty', $NULL, $Object, $Parameters)
        }

        # https://msdn.microsoft.com/en-us/library/aa772266%28v=vs.85%29.aspx
        If ($PSBoundParameters['Server']) {
            $ADSInitType = 2
            $InitName = $Server
        }ElseIf ($PSBoundParameters['Domain']) {
            $ADSInitType = 1
            $InitName = $Domain
        }ElseIf ($PSBoundParameters['Credential']) {
            $Cred = $Credential.GetNetworkCredential()
            $ADSInitType = 1
            $InitName = $Cred.Domain
        }Else {
            # if no domain or server is specified, default to GC initialization
            $ADSInitType = 3
            $InitName = $Null
        }
    }

    PROCESS {
        ForEach ($TargetIdentity in $Identity) {
            If (-not $PSBoundParameters['OutputType']) {
                If ($TargetIdentity -match "^[A-Za-z]+\\[A-Za-z ]+") {
                    $ADSOutputType = $NameTypes['DomainSimple']
                }Else {
                    $ADSOutputType = $NameTypes['NT4']
                }
            }Else {
                $ADSOutputType = $NameTypes[$OutputType]
            }

            $Translate = New-Object -ComObject NameTranslate

            If ($PSBoundParameters['Credential']) {
                Try {
                    $Cred = $Credential.GetNetworkCredential()

                    Invoke-Method -Object $Translate -Method 'InitEx' (
                        $ADSInitType,
                        $InitName,
                        $Cred.UserName,
                        $Cred.Domain,
                        $Cred.Password
                    )
                }Catch {
                    Write-Verbose "[Convert-ADName] Error initializing translation for '$($Identity)' using alternate credentials : $($_.Exception.InnerException.Message)"
                }
            }Else {
                Try {
                    $Null = Invoke-Method -Object $Translate -Method 'Init' -Parameters ($ADSInitType,$InitName)
                }Catch {
                    Write-Verbose "[Convert-ADName] Error initializing translation for '$($Identity)' : $($_.Exception.InnerException.Message)"
                }
            }

            # always chase all referrals
            Set-Property -Object $Translate -Property 'ChaseReferral' -Parameters (0x60)

            Try {
                # 8 = Unknown name type -> let the server do the work for us
                $Null = Invoke-Method -Object $Translate -Method 'Set' -Parameters (8, $TargetIdentity)
                Invoke-Method -Object $Translate -Method 'Get' -Parameters ($ADSOutputType)
            }Catch [System.Management.Automation.MethodInvocationException] {
                Write-Verbose "[Convert-ADName] Error translating '$($TargetIdentity)' : $($_.Exception.InnerException.Message)"
            }
        }
    }
}

## Have a look to https://github.com/PowerShellMafia/PowerSploit/tree/master