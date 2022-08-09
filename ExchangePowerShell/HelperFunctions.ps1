function global:Search-UnifiedLog
{
    <#
        .SYNOPSIS
            Use the Search-UnifiedLog function to search the unified audit log.
        .DESCRIPTION
            This function increases ResultSize to its maximum and uses SessionID in order to retrieve all entries.
        .PARAMETER UserIDs
            The UserIds parameter filters the log entries by the ID of the user who performed the action.
        .PARAMETER FreeText
            The FreeText parameter filters the log entries by the specified text string. If the value contains spaces, enclose the value in quotation marks (")
        .PARAMETER IPAddresses
            The IPAddresses parameter filters the log entries by the specified IP addresses. You specify multiple IP addresses separated by commas.
        .PARAMETER ObjectIds
            The ObjectIds parameter filters the log entries by object ID. The object ID is the target object that was acted upon, and depends on the RecordType and Operations values of the event. For example, for SharePoint operations, the object ID is the URL path to a file, folder, or site. For Azure Active Directory operations, the object ID is the account name or GUID value of the account.
        .PARAMETER Operations
            The Operations parameter filters the log entries by operation. The available values for this parameter depend on the RecordType value. For a list of the available values for this parameter, see Audited activities.
        .PARAMETER RecordType
            The RecordType parameter filters the log entries by record type.
        .PARAMETER SiteIds
            The SiteIds parameter filters the log entries by site ID. You can specify multiple values separated by commas.
        .PARAMETER StartDate
            The StartDate parameter specifies the start date of the date range. Entries are stored in the unified audit log in Coordinated Universal Time (UTC). If you specify a date/time value without a time zone, the value is in UTC.
        .PARAMETER EndDate
            The EndDate parameter specifies the end date of the date range. Entries are stored in the unified audit log in Coordinated Universal Time (UTC). If you specify a date/time value without a time zone, the value is in UTC.
        .PARAMETER SessionID
            The SessionId parameter specifies an ID you provide in the form of a string to identify a command (the cmdlet and its parameters) that will be run multiple times to return paged data. The SessionId can be any string value you choose and in this function a created GUID.
        .PARAMETER ResultSize
            The ResultSize parameter specifies the maximum number of results to return. The default value is 100, maximum is 5,000 (which is the default in this function).
        .PARAMETER Formatted
            The Formatted switch causes attributes that are normally returned as integers (for example, RecordType and Operation) to be formatted as descriptive strings. You don't need to specify a value with this switch.
        .EXAMPLE
            Search-UnifiedLog -StartDate 5/1/2018 -EndDate 5/2/2018
        .NOTES
            The function is using the Cmdlet Search-UnifiedAuditLog and set the parameter SessionCommand to ReturnLargeSet. In combination with SessionID and ResultSize all entries up to the maximum of 50,000 will be returned.
        .LINK
            https://docs.microsoft.com/powershell/module/exchange/search-unifiedauditlog?view=exchange-ps
    #>

    [CmdletBinding()]
    param(
        [System.String[]]
        $UserIDs,

        [System.String]
        $FreeText,

        [System.String[]]
        $IPAddresses,

        [System.String[]]
        $ObjectIds,

        [System.String[]]
        $Operations,

        [System.String]
        $RecordType,

        [System.String[]]
        $SiteIds,

        [System.DateTime]
        $StartDate = $((Get-Date).AddMonths(-1)),

        [System.DateTime]
        $EndDate = $(Get-Date),

        [System.String]
        $SessionID = $(([System.Guid]::NewGuid()).ToString()),

        [System.Int16]
        $ResultSize = '5000',

        [System.Management.Automation.SwitchParameter]
        $Formatted

    )

    begin
    {
        #$collection = [System.Collections.ArrayList]@()
        [System.Array]$collection = $null
        [System.Int16]$totalCount = 0
        [System.Array]$tempResult = $null
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $param = @{
            StartDate = $StartDate
            EndDate = $EndDate
            SessionId = $SessionID
            SessionCommand = 'ReturnLargeSet'
            ResultSize = $ResultSize
        }

        if (-not [System.String]::IsNullOrEmpty($UserIDs))
        {
            $param.Add('UserIds',$UserIDs)
        }
        if (-not [System.String]::IsNullOrEmpty($FreeText))
        {
            $param.Add('FreeText',$FreeText)
        }
        if (-not [System.String]::IsNullOrEmpty($IPAddresses))
        {
            $param.Add('IPAddresses',$IPAddresses)
        }
        if (-not [System.String]::IsNullOrEmpty($ObjectIds))
        {
            $param.Add('ObjectIds',$ObjectIds)
        }
        if (-not [System.String]::IsNullOrEmpty($Operations))
        {
            $param.Add('Operations',$Operations)
        }
        if (-not [System.String]::IsNullOrEmpty($RecordType))
        {
            $param.Add('RecordType',$RecordType)
        }
        if (-not [System.String]::IsNullOrEmpty($SiteIds))
        {
            $param.Add('SiteIds',$SiteIds)
        }
        if ($Formatted)
        {
            $param.Add('Formatted',$true)
        }
    }

    process
    {
        Write-Verbose "Start searching..."

        do
        {
            $tempResult = Search-UnifiedAuditLog @param
            if ($tempResult)
            {
                $collection += $tempResult
                Write-Verbose "TotalCount:$($collection[0].ResultCount) ResultIndex:$($tempResult.ResultIndex[-1]) Runtime:$($timer.Elapsed.ToString())"
            }
            else
            {
                Write-Verbose 'No records found!'
            }
        }
        until( $(if ($tempResult){ $tempResult.ResultIndex[-1] -ge $tempResult.ResultCount[-1]} else { return $true}) )
    }

    end
    {
        $collection | Sort-Object CreationDate
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Get-MessageTraceFull
{
    <#
        .SYNOPSIS
            Use the Get-MessageTraceFull function to search MessageTrace.
        .DESCRIPTION
            This function increases PageSize to its maximum of 5,000 and Page to 1. Use the Get-MessageTrace cmdlet to trace messages as they pass through the cloud-based organization. You can use this cmdlet to search message data for the last 10 days. If you run this cmdlet without any parameters, only data from the last 48 hours is returned.
        .PARAMETER EndDate
            The EndDate parameter specifies the end date of the date range.
        .PARAMETER Expression
            This parameter is reserved for internal Microsoft use.
        .PARAMETER FromIP
            The FromIP parameter filters the results by the source IP address. For incoming messages, the value of FromIP is the public IP address of the SMTP email server that sent the message. For outgoing messages from Exchange Online, the value is blank.
        .PARAMETER MessageId
            The MessageId parameter filters the results by the Message-ID header field of the message. This value is also known as the Client ID. The format of the Message-ID depends on the messaging server that sent the message. The value should be unique for each message. However, not all messaging servers create values for the Message-ID in the same way. Be sure to include the full Message ID string (which may include angle brackets) and enclose the value in quotation marks (for example, "d9683b4c-127b-413a-ae2e-fa7dfb32c69d@DM3NAM06BG401.Eop-nam06.prod.protection.outlook.com").
        .PARAMETER MessageTraceId
            The MessageTraceId parameter can be used with the recipient address to uniquely identify a message trace and obtain more details. A message trace ID is generated for every message that's processed by the system.
        .PARAMETER PageSize
            The PageSize parameter specifies the maximum number of entries per page. Valid input for this parameter is an integer between 1 and 5000. The default value is 1000.
        .PARAMETER ProbeTag
            This parameter is reserved for internal Microsoft use.
        .PARAMETER RecipientAddress
            The RecipientAddress parameter filters the results by the recipient's email address. You can specify multiple values separated by commas.
        .PARAMETER SenderAddress
            The SenderAddress parameter filters the results by the sender's email address. You can specify multiple values separated by commas.
        .PARAMETER StartDate
            The StartDate parameter specifies the start date of the date range.
        .PARAMETER Status
            The Status parameter filters the results by the delivery status of the message.
        .PARAMETER ToIP
            The ToIP parameter filters the results by the destination IP address. For outgoing messages, the value of ToIP is the public IP address in the resolved MX record for the destination domain. For incoming messages to Exchange Online, the value is blank.
        .EXAMPLE
            Get-MessageTraceFull -SenderAddress john@contoso.com -StartDate 06/13/2018 -EndDate 06/15/2018
        .NOTES
            The function uses the Cmdlet Get-MessageTrace and is doing paging for you in order to retrieve up to the maximum 1,000,000 entries.
        .LINK
            https://docs.microsoft.com/powershell/module/exchange/get-messagetrace?view=exchange-ps
    #>
    [CmdletBinding()]
    param(
        [System.DateTime]
        $EndDate,

        [System.Linq.Expressions.Expression]
        $Expression,

        [System.String]
        $FromIP,

        [System.String[]]
        $MessageId,

        [System.GUID]
        $MessageTraceId,

        [System.Int32]
        $PageSize = '5000',

        [System.String]
        $ProbeTag,

        [System.String[]]
        $RecipientAddress,

        [System.String[]]
        $SenderAddress,

        [System.DateTime]
        $StartDate,

        [System.String[]]
        [ValidateSet('None', 'GettingStatus', 'Failed', 'Pending', 'Delivered', 'Expanded', 'Quarantined', 'FilteredAsSpam')]
        $Status,

        [System.String]
        $ToIP

    )

    begin
    {
        $collection = [System.Collections.ArrayList]@()
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        [System.Boolean]$haveMore = $true
        $param = @{}
        [System.Int16]$PageCounter = '0'

        if (-not [System.String]::IsNullOrEmpty($EndDate))
        {
            $param.Add('EndDate',$EndDate)
        }
        if (-not [System.String]::IsNullOrEmpty($Expression))
        {
            $param.Add('Expression',$Expression)
        }
        if (-not [System.String]::IsNullOrEmpty($FromIP))
        {
            $param.Add('FromIP',$FromIP)
        }
        if (-not [System.String]::IsNullOrEmpty($MessageId))
        {
            $param.Add('MessageId',$MessageId)
        }
        if (-not [System.String]::IsNullOrEmpty($MessageTraceId))
        {
            $param.Add('MessageTraceId',$MessageTraceId)
        }
        if (-not [System.String]::IsNullOrEmpty($PageSize))
        {
            $param.Add('PageSize',$PageSize)
        }
        if (-not [System.String]::IsNullOrEmpty($ProbeTag))
        {
            $param.Add('ProbeTag',$ProbeTag)
        }
        if (-not [System.String]::IsNullOrEmpty($RecipientAddress))
        {
            $param.Add('RecipientAddress',$RecipientAddress)
        }
        if (-not [System.String]::IsNullOrEmpty($SenderAddress))
        {
            $param.Add('SenderAddress',$SenderAddress)
        }
        if (-not [System.String]::IsNullOrEmpty($StartDate))
        {
            $param.Add('StartDate',$StartDate)
        }
        if (-not [System.String]::IsNullOrEmpty($Status))
        {
            $param.Add('Status',$Status)
        }
        if (-not [System.String]::IsNullOrEmpty($ToIP))
        {
            $param.Add('ToIP',$ToIP)
        }
        $param.Add('Page',[System.Int16]'1')
    }

    process
    {
        while ($haveMore)
        {
            $tempResult = $null
            $tempResult = Get-MessageTrace @param
            $collection += $tempResult
            Write-Verbose "TotalCount:$($collection.Count) Page:$($param.Page) Runtime:$($timer.Elapsed.ToString()) ResultCount:$($tempResult.Count)"

            if ($tempResult.Count -eq $PageSize)
            {
                Write-Verbose "Increasing Page number"
                $param.Page++
            }
            else
            {
                $haveMore = $false
            }
        }
    }

    end
    {
        $collection | Sort-Object Received
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }

}

function global:Prompt
{
    <#
        .SYNOPSIS
            The function customizes your PowerShell window.
        .DESCRIPTION
            The function customize your PowerShell window based on your connection: Either EXO or SCC.
    #>
    Import-Module -Name ExchangeOnlineManagement -Force
    [System.Boolean]$elevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')
    $connectionContexts = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetAllConnectionContexts() | Where-Object {-not [System.String]::IsNullOrEmpty($_.PowerShellCredentials)}
    foreach ($context in $connectionContexts)
    {
        switch ($context.ConnectionUri)
        {
            'https://outlook.office365.com'                 {$ConnectedTo = 'EXO'}
            'https://ps.compliance.protection.outlook.com'  {$ConnectedTo = 'SCC'}
        }
        $connectString += "Connected to $($ConnectedTo) as $($context.PowerShellCredentials.UserName) IsRpsSession:$($context.IsRpsSession) ExoModuleVersion:$($context.ExoModuleVersion) RoutingHint:$($context.RoutingHint)"
    }

    $Host.UI.RawUI.WindowTitle = (Get-Date $([System.DateTime]::Now.ToUniversalTime()) -UFormat '%y/%m/%d %R').Tostring() + " UTC $($connectString) ProcessID:$PID Elevated:$elevated"
    Write-Host '[' -NoNewline
    Write-Host (Get-Date $([System.DateTime]::Now.ToUniversalTime()) -UFormat '%T')-NoNewline
    Write-Host ' UTC]:' -NoNewline
    Write-Host (Split-Path (Get-Location) -Leaf) -NoNewline
    return "> "
}
Prompt

function global:Get-ManagedFolderAssistantLog
{
    <#
        .SYNOPSIS
            This function retrieves ECL for a given mailbox.
        .DESCRIPTION
            This function retrieves and format the MFA log for a given mailbox. It uses the Cmdlet Export-MailboxDiagnosticLogs for this.
        .PARAMETER Identity
            The Identity parameter specifies that mailbox that contains the diagnostics logs that you want to view. You can use any value that uniquely identifies the mailbox.
        .EXAMPLE
            Get-ManagedFolderAssistantLog -Identity ingo@bla.com | Select-Object -ExpandProperty ecl
        .LINK
            https://docs.microsoft.com/powershell/module/exchange/export-mailboxdiagnosticlogs?view=exchange-ps
            https://ingogegenwarth.wordpress.com/2017/11/20/advanced-cal/
            https://timmcmic.wordpress.com/2019/02/13/office-365-tracking-last-run-times-of-the-managed-folder-assistance-exchange-life-cycle/
    #>
    [CmdletBinding()]
    param(
        [System.String[]]
        $Identity
    )

    begin
    {
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $collection = [System.Collections.ArrayList]@()
    }

    process
    {
        foreach ($ID in $Identity)
        {
            Write-Verbose "Processing $($ID)..."
            $data = New-Object -TypeName PSObject
            $data | add-member -type NoteProperty -Name Identity -Value $ID
            $data | add-member -type NoteProperty -Name Ecl -Value $(([xml](Export-MailboxDiagnosticLogs -Identity $ID -ExtendedProperties).MailboxLog).Properties.MailboxTable.Property | ? name -match 'elc' | ? name -NotMatch 'funnel')
            $collection.Add($data) | Out-Null
        }
    }

    end
    {
        $collection
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Get-QuarantineMessageFull
{
    <#
        .SYNOPSIS
            Use the Get-QuarantineMessageFull function to retrieve all quarantined messages.
        .DESCRIPTION
            
        .PARAMETER Direction
            The Direction parameter filters the results by incoming or outgoing messages. Valid are Inbound and Outbound.
        .PARAMETER Domain
            The Domain parameter filters the results by sender or recipient domain. You can specify multiple domain values separated by commas.
        .PARAMETER EndExpiresDate
            The EndExpiresDate parameter specifies the latest messages that will automatically be deleted from the quarantine. Use this parameter with the StartExpiresDate parameter.
        .PARAMETER EndReceivedDate
            The EndReceivedDate parameter specifies the latest messages to return in the results. Use this parameter with the StartReceivedDate parameter.
        .PARAMETER Identity
            The Identity parameter specifies the quarantined message that you want to view. The value is a unique quarantined message identifier in the format GUID1\GUID2 (for example c14401cf-aa9a-465b-cfd5-08d0f0ca37c5\4c2ca98e-94ea-db3a-7eb8-3b63657d4db7).
        .PARAMETER MessageId
            The MessageId parameter filters the results by the Message-ID header field of the message. This value is also known as the Client ID. The format of the Message-ID depends on the messaging server that sent the message. The value should be unique for each message. However, not all messaging servers create values for the Message-ID in the same way. Be sure to include the full Message ID string (which may include angle brackets) and enclose the value in quotation marks (for example, "<d9683b4c-127b-413a-ae2e-fa7dfb32c69d@DM3NAM06BG401.Eop-nam06.prod.protection.outlook.com>").
        .PARAMETER MyItems
            The MyItems switch filters the results by messages where you (the user that's running the command) are the recipient. You don't need to specify a value with this switch.
        .PARAMETER Page
            The Page parameter specifies the page number of the results you want to view. Valid input for this parameter is an integer between 1 and 1000. The default value is 1.
        .PARAMETER PageSize
            The PageSize parameter specifies the maximum number of entries per page. Valid input for this parameter is an integer between 1 and 1000. The default value is 100.
        .PARAMETER QuarantineTypes
            The QuarantineTypes parameter filters the results by what caused the message to be quarantined.
        .PARAMETER RecipientAddress
            The RecipientAddress parameter filters the results by the recipient's email address. You can specify multiple values separated by commas.
        .PARAMETER Reported
            The Reported parameter filters the results by messages that have already been reported as false positives.
        .PARAMETER SenderAddress
            The SenderAddress parameter filters the results by the sender's email address. You can specify multiple values separated by commas.
        .PARAMETER StartExpiresDate
            The StartExpiresDate parameter specifies the earliest messages that will automatically be deleted from the quarantine. Use this parameter with the EndExpiresDate parameter.
        .PARAMETER StartReceivedDate
            The StartReceivedDate parameter specifies the earliest messages to return in the results. Use this parameter with the EndReceivedDate parameter.
        .PARAMETER Subject
            The Subject parameter filters the results by the subject field of the message. If the value contains spaces, enclose the value in quotation marks (").
        .PARAMETER Type
            The Type parameter filters the results by what caused the message to be quarantined.
        .EXAMPLE
            Get-QuarantineMessageFull -StartReceivedDate 06/13/2016 -EndReceivedDate 06/15/2016
        .NOTES
            The function increases PageSize to its maximum of 1,000 and is doing paging until all entries have been retrieved.
        .LINK
            https://docs.microsoft.com/powershell/module/exchange/get-quarantinemessage?view=exchange-ps
    #>
    [CmdletBinding()]
    param(
        [System.String]
        $Direction,

        [System.String[]]
        $Domain,

        [System.DateTime]
        $EndExpiresDate,

        [System.DateTime]
        $EndReceivedDate,

        [System.String]
        $Identity,

        [System.String]
        $MessageId,

        [System.Management.Automation.SwitchParameter]
        $MyItems,

        [System.Int32]
        $Page = '1',

        [System.Int32]
        $PageSize = '1000',

        [System.String[]]
        [ValidateSet('Bulk', 'HighConfPhish', 'Phish', 'Spam', 'Malware', 'TransportRule')]
        $QuarantineTypes,

        [System.String[]]
        $RecipientAddress,

        [System.Boolean]
        $Reported,

        [System.String[]]
        $SenderAddress,

        [System.DateTime]
        $StartExpiresDate,

        [System.DateTime]
        $StartReceivedDate,

        [System.String]
        $Subject,

        [System.String]
        [ValidateSet('Bulk', 'HighConfPhish', 'Phish', 'Spam', 'TransportRule')]
        $Type

    )

    begin
    {
        $collection = [System.Collections.ArrayList]@()
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        [System.Boolean]$haveMore = $true
        $param = @{}

        if (-not [System.String]::IsNullOrEmpty($Direction))
        {
            $param.Add('Direction',$Direction)
        }
        if (-not [System.String]::IsNullOrEmpty($Domain))
        {
            $param.Add('Domain',$($Domain -join ','))
        }
        if (-not [System.String]::IsNullOrEmpty($EndExpiresDate))
        {
            $param.Add('EndExpiresDate',$EndExpiresDate)
        }
        if (-not [System.String]::IsNullOrEmpty($EndReceivedDate))
        {
            $param.Add('EndReceivedDate',$EndReceivedDate)
        }
        if (-not [System.String]::IsNullOrEmpty($Identity))
        {
            $param.Add('Identity',$Identity)
        }
        if (-not [System.String]::IsNullOrEmpty($MyItems))
        {
            $param.Add('MyItems',$MyItems)
        }
        if (-not [System.String]::IsNullOrEmpty($Page))
        {
            $param.Add('Page',$Page)
        }
        if (-not [System.String]::IsNullOrEmpty($PageSize))
        {
            $param.Add('PageSize',$PageSize)
        }
        if (-not [System.String]::IsNullOrEmpty($QuarantineTypes))
        {
            $param.Add('QuarantineTypes',$($QuarantineTypes -join ','))
        }
        if (-not [System.String]::IsNullOrEmpty($RecipientAddress))
        {
            $param.Add('RecipientAddress',$($RecipientAddress -join ','))
        }
        if (-not [System.String]::IsNullOrEmpty($Reported))
        {
            $param.Add('Reported',$Reported)
        }
        if (-not [System.String]::IsNullOrEmpty($SenderAddress))
        {
            $param.Add('SenderAddress',$($SenderAddress -join ','))
        }
        if (-not [System.String]::IsNullOrEmpty($StartExpiresDate))
        {
            $param.Add('StartExpiresDate',$StartExpiresDate)
        }
        if (-not [System.String]::IsNullOrEmpty($StartReceivedDate))
        {
            $param.Add('StartReceivedDate',$StartReceivedDate)
        }
        if (-not [System.String]::IsNullOrEmpty($Subject))
        {
            $param.Add('Subject',$Subject)
        }
        if (-not [System.String]::IsNullOrEmpty($Type))
        {
            $param.Add('Type',$Type)
        }
    }

    process
    {
        while ($haveMore)
        {
            $tempResult = $null
            $tempResult = Get-QuarantineMessage @param
            $collection += $tempResult

            Write-Verbose "TotalCount:$($collection.Count) Page:$($param.Page) Runtime:$($timer.Elapsed.ToString()) ResultCount:$($tempResult.Count)"
            if ($tempResult.Count -eq $PageSize)
            {
                Write-Verbose "Increasing Page number"
                $param.Page++
            }
            else
            {
                $haveMore = $false
            }
        }
    }

    end
    {
        $collection | Sort-Object Received
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }

}

function global:Test-ExchangeAuditSetting
{
    <#
        .SYNOPSIS
            With this function you can compare Audit settings of mailboxes.
        .DESCRIPTION
            The function will retrieve the current mailbox audit and compare with desired settings.
        .PARAMETER Mailbox
            The Mailbox parameter specifies the mailbox you want to check. Could be a single or multiple ones. Piping is supported. You need pass the whole object from either Get-Mailbox or Get-EXOMailbox (here include the properties AuditOwnerDesired and AuditDelegateDesired!).
        .PARAMETER AuditOwnerDesired
            The AuditOwnerDesired parameter specifies an array of audited events for OwnerAccess.
        .PARAMETER AuditDelegateDesired
            The AuditDelegateDesired parameter specifies an array of audited events for DelegateAccess.
        .EXAMPLE
            Get-Mailbox -Identity ingo@bla.com | Test-ExchangeAuditSetting
            Get-EXOMailbox -Identity ingo@bla.com -Properties AuditOwner,AuditDelegate
            Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize unlimited -Properties AuditOwner,AuditDelegate | Test-ExchangeAuditSetting
        .LINK
            https://docs.microsoft.com/exchange/policy-and-compliance/mailbox-audit-logging/mailbox-audit-logging?view=exchserver-2019
    #>
    [CmdletBinding()]
    param(
        [Parameter(
            ValueFromPipeline=$true,
            Mandatory=$true,
            Position=0)]
        [System.Object[]]
        $Mailbox,
    
        [Parameter(
            Mandatory=$false,
            Position=1)]
        [System.String[]]
        $AuditOwnerDesired = @("Update","MoveToDeletedItems","SoftDelete","HardDelete","Create","UpdateFolderPermissions","UpdateInboxRules","UpdateCalendarDelegation","MailItemsAccessed","MailboxLogin"),
    
        [Parameter(
            Mandatory=$false,
            Position=2)]
        [System.String[]]
        $AuditDelegateDesired = @("Update","MoveToDeletedItems","SoftDelete","HardDelete","SendAs","SendOnBehalf","Create","UpdateFolderPermissions","UpdateInboxRules","MailItemsAccessed","FolderBind")
    )

    begin
    {

        $collection = [System.Collections.ArrayList]@()
        $toBeProcessed = [System.Collections.ArrayList]@()
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        <#
            .SYNOPSIS
                Takes an array of strings and converts each element in the array to
                all lowercase characters.

            .PARAMETER Array
                The array of System.String objects to convert into lowercase strings.
        #>
        function Convert-StringArrayToLowerCase
        {
            [CmdletBinding()]
            [OutputType([System.String[]])]
            param
            (
                [Parameter()]
                [System.String[]]
                $Array
            )

            [System.String[]] $arrayOut = New-Object -TypeName 'System.String[]' -ArgumentList $Array.Count

            for ($i = 0; $i -lt $Array.Count; $i++)
            {
                $arrayOut[$i] = $Array[$i].ToLower()
            }

            return $arrayOut
        }

        <#
            .SYNOPSIS
                Returns whether two string arrays have the same contents, where element
                order doesn't matter.
        
            .PARAMETER Array1
                The first System.String[] object to compare.
        
            .PARAMETER Array2
                The second System.String[] object to compare.
        
            .PARAMETER IgnoreCase
                Specifies that case should be ignored when comparing array contents.
        #>
        function Compare-ArrayContent
        {
            [CmdletBinding()]
            [OutputType([System.Boolean])]
            param
            (
                [Parameter()]
                [System.String[]]
                $Array1,

                [Parameter()]
                [System.String[]]
                $Array2,

                [Parameter()]
                [System.Management.Automation.SwitchParameter]
                $IgnoreCase
            )

            $hasSameContents = $true

            if ($Array1.Length -ne $Array2.Length)
            {
                $hasSameContents = $false
            }
            elseif ($Array1.Count -gt 0 -and $Array2.Count -gt 0)
            {
                if ($IgnoreCase -eq $true)
                {
                    $Array1 = Convert-StringArrayToLowerCase -Array $Array1
                    $Array2 = Convert-StringArrayToLowerCase -Array $Array2
                }

                foreach ($str in $Array1)
                {
                    if (!($Array2.Contains($str)))
                    {
                        $hasSameContents = $false
                        break
                    }
                }
            }

            return $hasSameContents
        }

        [System.Int32]$i='1'

    }

    process
    {

        foreach($ID in $Mailbox)
        {
            $toBeProcessed.Add($ID) | Out-Null
        }

    }

    end{

        foreach($ID in $toBeProcessed)
        {
            Write-Progress -id 1 -Activity "Processing mailbox - $($ID.PrimarySmtpAddress)" -PercentComplete ( $i / $toBeProcessed.count * 100) -Status "Remaining objects: $($toBeProcessed.count - $i)"

            $data = New-Object -TypeName PSObject
            $data | add-member -type NoteProperty -Name PrimarySmtpAddress -Value $($ID.PrimarySmtpAddress)

            if(-not [System.String]::IsNullOrEmpty($ID.AuditOwner))
            {
                $data | add-member -type NoteProperty -Name AuditOwner -Value $(Compare-ArrayContent -Array1 $AuditOwnerDesired -Array2 $ID.AuditOwner -IgnoreCase)
            }
            else
            {
                $data | add-member -type NoteProperty -Name AuditOwner -Value 'N/A'
            }

            if(-not [System.String]::IsNullOrEmpty($ID.AuditDelegate))
            {
                $data | add-member -type NoteProperty -Name AuditDelegate -Value $(Compare-ArrayContent -Array1 $AuditOwnerDesired -Array2 $ID.AuditOwner -IgnoreCase)
            }
            else
            {
                $data | add-member -type NoteProperty -Name AuditDelegate -Value 'N/A'
            }

            $collection.Add($data) | Out-Null
            $i++
        }

        Write-Progress -Activity "Processing mailbox - $($ID.PrimarySmtpAddress)" -Status "Ready" -Completed
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
        $collection
    }

}

function global:Get-EASDetails {
    <#
        .SYNOPSIS
            The function pulls the properties basic properties of mobile device statistics
        .DESCRIPTION
            The function is using Cmdlet Get-EXOMobileDeviceStatistics for mailboxes and Get-MobileDevice for deviceIDs and provides the following attributes: DeviceModel, DeviceOS, DeviceID, DeviceUserAgent, LastSyncAttemptTime, LastSuccessSync, DeviceAccessState, 
        .PARAMETER Mailbox
            This input parameter acts as Identity filter. The Mailbox parameter filters the results by the user mailbox that's associated with the mobile device. You can use any value that uniquely identifies the mailbox.
        .PARAMETER DeviceID
            The DeviceID parameter is used for filtering by DeviceID.
        .EXAMPLE
            Get-EASDetails -Mailbox ingo@bla.com -Verbose | Format-Table -AutoSize
    #>
    param(
    [parameter( Mandatory=$false, ParameterSetName="Mailbox")]
    [parameter( Position=0)]
    [System.String]$Mailbox,

    [parameter( Mandatory=$false, ParameterSetName="DeviceID")]
    [parameter( Position=1)]
    [System.String]$DeviceID
    )

    begin
    {
        if ($Mailbox)
        {
            $command = 'Get-EXOMobileDeviceStatistics -Mailbox ' + $Mailbox
            $processingObject = $Mailbox
        }
        else
        {
            $command = 'Get-MobileDevice -Filter {DeviceID -eq "' + $DeviceID + '"} | Sort-Object | ForEach{Get-MobileDeviceStatistics $_.identity }'
            $processingObject = $DeviceID
        }
    }

    process {
        try {
            Write-Warning "Working on $($processingObject)..."
            Invoke-Expression $command  | Sort-Object LastSuccessSync | Select-Object DeviceModel,DeviceOS,DeviceID,DeviceUserAgent,LastSyncAttemptTime,LastSuccessSync,DeviceAccessState
        }
        catch{
            $_.Exception
        }
    }
}

function global:Enable-PIMRole
{
    <#
        .SYNOPSIS
            This function simplifies the process for elevating your account.
        .DESCRIPTION
            The function simplifies the process of account elevation as you can specify the human readable role. It will lookup the role and request elevation for up-to 10 hours. Default reason will be used.
        .PARAMETER UserPrincipalName
            The account's UPN, for which role elevation is requested.
        .PARAMETER Role
            The role, which is requested.
        .PARAMETER Hours
            Rather specifying start and end time, a schedule of hours is used. The maximum is 10.
        .PARAMETER Reason
            The reason for elevation
        .PARAMETER ListActiveRoles
            Lists all active roles
        .PARAMETER ListEligibleRoles
            Lists roles, which are eligible for user to activate
        .PARAMETER ListRoleAssignmentScheduleRequest
            Lists all schedule resquests
        .PARAMETER UseAzureADPreview
            Using deprecated AzureADPreview module
        .EXAMPLE
            Enable-PIMRole -UserPrincipalName ingo@bla.com -Role 'Global Administrator'
            Enable-PIMRole -UserPrincipalName ingo@bla.com -Role 'Exchange Administrator'
        .NOTES
            The function and the new PIM module requires the latest AzureADPreview module as AzureAD module doesn't support the new PIM requests.
        .LINK
            https://docs.microsoft.com/azure/active-directory/privileged-identity-management/powershell-for-azure-ad-roles
    #>
    [CmdletBinding(DefaultParameterSetName = "EnableRole")]
    Param
    (
        [parameter(
            ParameterSetName="EnableRole",
            Mandatory=$true,
            Position=0
        )]
        [parameter(
            ParameterSetName="UseAzureADPreview",
            Mandatory=$true,
            Position=0
        )]
        [parameter(
            ParameterSetName="ListEligibleRoles",
            Mandatory=$true,
            Position=0
        )]
        [parameter(
            ParameterSetName="ListRoleAssignmentScheduleRequest",
            Mandatory=$false,
            Position=0
        )]
        [parameter(
            ParameterSetName="ListActiveRoles",
            Mandatory=$false,
            Position=0
        )]
        [System.String]
        $UserPrincipalName,

        [parameter(
            ParameterSetName="EnableRole",
            Position=2
        )]
        [parameter(
            ParameterSetName="UseAzureADPreview",
            Position=2
        )]
        [System.Int16]
        [ValidateRange(1,10)]
        $Hours = '10',

        [parameter(
            ParameterSetName="EnableRole",
            Position=3
        )]
        [parameter(
            ParameterSetName="UseAzureADPreview",
            Position=3
        )]
        [System.String]
        [ValidateNotNullOrEmpty()]
        $Reason = 'Daily work',

        [parameter(
            ParameterSetName="ListActiveRoles",
            Mandatory=$false,
            Position=1
        )]
        [System.Management.Automation.SwitchParameter]
        $ListActiveRoles,

        [parameter(
            ParameterSetName="ListEligibleRoles",
            Mandatory=$false,
            Position=1
        )]
        [System.Management.Automation.SwitchParameter]
        $ListEligibleRoles,

        [parameter(
            ParameterSetName="ListRoleAssignmentScheduleRequest",
            Mandatory=$false,
            Position=1
        )]
        [System.Management.Automation.SwitchParameter]
        $ListRoleAssignmentScheduleRequest,

        [parameter(
            ParameterSetName="UseAzureADPreview",
            Mandatory=$false,
            Position=4
        )]
        [System.Management.Automation.SwitchParameter]
        $UseAzureADPreview
    )
    
    DynamicParam
    {
        $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
        $attributeEnableRole = New-Object System.Management.Automation.ParameterAttribute
        $attributeEnableRole.ParameterSetName = 'EnableRole'
        $attributeEnableRole.Position = '1'
        $attributeEnableRole.Mandatory = $true

        $attributeUseAzureADPreview = New-Object System.Management.Automation.ParameterAttribute
        $attributeUseAzureADPreview.ParameterSetName = 'UseAzureADPreview'
        $attributeUseAzureADPreview.Position = '1'
        $attributeUseAzureADPreview.Mandatory = $true

        $attributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
        $attributeCollection.Add($attributeEnableRole)
        $attributeCollection.Add($attributeUseAzureADPreview)

        if ($UseAzureADPreview)
        {
            Write-Verbose 'Remove existing "old" AzureAD module and load AzureADPreview'
            Remove-Module Azuread -Force -ErrorAction silentlycontinue
            Import-Module AzureADPreview -Verbose:$false
            $AAD = Connect-AzureAD -AccountId $UserPrincipalname
            $resource = Get-AzureADMSPrivilegedResource -ProviderId AadRoles
            $script:allRoles = Get-AzureADMSPrivilegedRoleDefinition -ProviderId AadRoles -ResourceId $resource.Id
            $roleSet = $allRoles.DisplayName
        }
        else
        {
            # check for required permissions
            $requiredMGPermissions = @('User.ReadBasic.All','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
            # get current context
            $currentMGContext = Get-MgContext
            if (-not [System.String]::IsNullOrEmpty($currentMGContext))
            {
                foreach ($permission in $requiredMGPermissions)
                {
                    if (-not $currentMGContext.Scopes.Contains($permission))
                    {
                        Write-Error "Required permission missing:$($permission)"
                        [System.Boolean]$insufficientPerms = $true
                    }
                }
                if ($insufficientPerms)
                {
                    Write-Warning 'No existing connection. Please connect to MS Graph first! e.g.:Connect-MgGraph -Scopes User.ReadBasic.All,RoleEligibilitySchedule.Read.Directory,RoleAssignmentSchedule.ReadWrite.Directory'
                    break
                }
            }
            else
            {
                Import-Module Microsoft.Graph.DeviceManagement.Enrolment -Verbose:$false
                $null = Connect-MgGraph -Scopes User.ReadBasic.All,RoleEligibilitySchedule.Read.Directory,RoleAssignmentSchedule.ReadWrite.Directory
            }

            # Generate and set the ValidateSet
            #$roleSet = (Get-MgRoleManagementDirectoryRoleDefinition -All).DisplayName
            $script:allRoles = Get-MgRoleManagementDirectoryRoleDefinition -All
            $roleSet = $allRoles.DisplayName
        }
        $validateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($roleSet)
        $attributeCollection.Add($validateSetAttribute)
        
        $runtimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter('Role', [string], $attributeCollection)
        $paramDictionary.Add('Role', $runtimeParameter)
        return $paramDictionary

    }

    begin
    {
        $Error.Clear()
    }

    process
    {

        try {
            if ($UseAzureADPreview)
            {
                $AAD=Connect-AzureAD -AccountId $UserPrincipalname
                $resource = Get-AzureADMSPrivilegedResource -ProviderId AadRoles
                $roleDefinition = $script:allRoles | Where-Object -FilterScript {$_.DisplayName -eq $PSBoundParameters.Role}
                $subject = Get-AzureADUser -Filter "userPrincipalName eq '$($UserPrincipalname)'"
                $schedule = New-Object Microsoft.Open.MSGraph.Model.AzureADMSPrivilegedSchedule
                $schedule.Type = "Once"
                $schedule.Duration = "PT$($Hours)H"
    
                $MyRole = @{
                    ProviderId = 'aadRoles'
                    ResourceId = $resource.Id
                    SubjectID = $subject.ObjectId
                    AssignmentState = 'Active'
                    Type = 'UserAdd'
                    Reason =$Reason
                    RoleDefinitionId = $roleDefinition.Id
                    Schedule = $schedule
                    ErrorAction = 'Stop'
                }

                Open-AzureADMSPrivilegedRoleAssignmentRequest @Myrole

            }
            elseif ($ListEligibleRoles)
            {
                # initiate collection
                $collection = [System.Collections.ArrayList]@()
                $rawEligibleRoles = Invoke-MgGraphRequest -Method GET -Uri "v1.0/roleManagement/directory/roleEligibilitySchedules/microsoft.graph.filterByCurrentUser(on='principal')"
                $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All

                foreach ($eligibleRole in $rawEligibleRoles.value)
                {
                    Write-Verbose "Processing role with roleDefinitionId:$($eligibleRole.roleDefinitionId)"
                    # create role object
                    $roleDetails = New-Object -TypeName psobject
                    # get roleDefinition
                    $roleDefinition = $roleDefinitions | Where-Object -FilterScript {$_.Id -EQ $eligibleRole.roleDefinitionId }
                    $roleDetails | Add-Member -MemberType NoteProperty -Name createdUsing -Value $eligibleRole.createdUsing
                    $roleDetails | Add-Member -MemberType NoteProperty -Name createdDateTime -Value $eligibleRole.createdDateTime
                    $roleDetails | Add-Member -MemberType NoteProperty -Name directoryScopeId -Value $eligibleRole.directoryScopeId
                    $roleDetails | Add-Member -MemberType NoteProperty -Name modifiedDateTime -Value $eligibleRole.modifiedDateTime
                    $roleDetails | Add-Member -MemberType NoteProperty -Name scheduleInfo -Value $eligibleRole.scheduleInfo
                    $roleDetails | Add-Member -MemberType NoteProperty -Name id -Value $eligibleRole.id
                    $roleDetails | Add-Member -MemberType NoteProperty -Name memberType -Value $eligibleRole.memberType
                    $roleDetails | Add-Member -MemberType NoteProperty -Name status -Value $eligibleRole.status
                    $roleDetails | Add-Member -MemberType NoteProperty -Name roleDefinitionId -Value $eligibleRole.roleDefinitionId
                    $roleDetails | Add-Member -MemberType NoteProperty -Name Description -Value $roleDefinition.Description
                    $roleDetails | Add-Member -MemberType NoteProperty -Name DisplayName -Value $roleDefinition.DisplayName
                    $roleDetails | Add-Member -MemberType NoteProperty -Name RolePermissions -Value $roleDefinition.RolePermissions

                    # add to collection
                    $collection += $roleDetails
                }
                $collection
            }
            elseif ($ListRoleAssignmentScheduleRequest)
            {
                $subject = Get-MgUser -Filter "userPrincipalName eq '$($UserPrincipalname)'"
                $paramsoleAssignmentScheduleRequest = @{
                    Filter = "principalId eq '$($subject.Id)'"
                    All = $true
                }
                Get-MgRoleManagementDirectoryRoleAssignmentScheduleRequest @paramsoleAssignmentScheduleRequest
            }
            elseif ($ListActiveRoles)
            {
                # initiate collection
                $collection = [System.Collections.ArrayList]@()
                $subject = Get-MgUser -Filter "userPrincipalName eq '$($UserPrincipalname)'"
                $rawActiveRoles = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter "principalId eq '$($subject.Id)'" -ExpandProperty activatedUsing
                $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All

                foreach ($activeRole in $rawActiveRoles)
                {
                    # create role object
                    $activeDetails = New-Object -TypeName psobject
                    
                    # get roleDefinition
                    $roleDefinition = $roleDefinitions | Where-Object -FilterScript {$_.Id -EQ $activeRole.roleDefinitionId }
                    $activeDetails | Add-Member -MemberType NoteProperty -Name ActivatedUsing -Value $activeRole.ActivatedUsing
                    $activeDetails | Add-Member -MemberType NoteProperty -Name AppScope -Value $activeRole.AppScope
                    $activeDetails | Add-Member -MemberType NoteProperty -Name AppScopeId -Value $activeRole.AppScopeId
                    $activeDetails | Add-Member -MemberType NoteProperty -Name AssignmentType -Value $activeRole.AssignmentType
                    $activeDetails | Add-Member -MemberType NoteProperty -Name DirectoryScope -Value $activeRole.DirectoryScope
                    $activeDetails | Add-Member -MemberType NoteProperty -Name DirectoryScopeId -Value $activeRole.DirectoryScopeId
                    $activeDetails | Add-Member -MemberType NoteProperty -Name StartDateTime -Value $activeRole.StartDateTime
                    $activeDetails | Add-Member -MemberType NoteProperty -Name EndDateTime -Value $activeRole.EndDateTime
                    $activeDetails | Add-Member -MemberType NoteProperty -Name PrincipalId -Value $activeRole.PrincipalId
                    $activeDetails | Add-Member -MemberType NoteProperty -Name RoleDefinitionId -Value $activeRole.RoleDefinitionId
                    $activeDetails | Add-Member -MemberType NoteProperty -Name Description -Value $roleDefinition.Description
                    $activeDetails | Add-Member -MemberType NoteProperty -Name DisplayName -Value $roleDefinition.DisplayName
                    $activeDetails | Add-Member -MemberType NoteProperty -Name RolePermissions -Value $roleDefinition.RolePermissions

                    # add to collection
                    $collection += $activeDetails
                }
                $collection
            }
            else
            {
                # get role
                $roleDefinition = $script:allRoles | Where-Object -FilterScript {$_.DisplayName -eq $PSBoundParameters.Role}
                $subject = Get-MgUser -Filter "userPrincipalName eq '$($UserPrincipalname)'"

                $body = @{
                    action = "selfActivate"
                    justification = "$Reason"
                    roleDefinitionId = $roleDefinition.Id
                    directoryScopeId = "/"
                    principalId = $subject.Id
                    #startDateTime = "2022-08-06 19:47:46Z"
                    scheduleInfo = @{
                        startDateTime = Get-Date ([System.DateTime]::UtcNow) -Format o
                        #startDateTime = "2022-08-06T20:58:00Z"
                        expiration = @{
                            type = "AfterDuration"
                            duration = "PT$($Hours)H"
                        }
                    }
                }
                New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
            }
        }
        catch {
            $Error[0].Exception
        }
    }
}

function global:Get-MSGraphGroup
{
    <#
        .SYNOPSIS
            This function uses Microsoft Office application for retrieving access token and queries Microsoft Graph for group properties.
        .DESCRIPTION
            The Microsoft Office with ClientID d3590ed6-52b3-4102-aeff-aad2292ab01c can be used to retrieve an access token with the scopes AuditLog.Read.All, Calendar.ReadWrite, Calendars.Read.Shared, Calendars.ReadWrite, Contacts.ReadWrite, DeviceManagementConfiguration.Read.All, DeviceManagementConfiguration.ReadWrite.All, Directory.AccessAsUser.All, Directory.Read.All, email, Files.Read, Files.Read.All, Group.Read.All, Group.ReadWrite.All, Mail.ReadWrite, openid, People.Read, People.Read.All, profile, User.Read.All, User.ReadWrite, Users.Read
        .PARAMETER Group
            The parameter Group defines the id of the group. Unless you use the parameter ByMail. If this parameter is used in addition, the function tries to get the id of the group by searching for a group with the specified e-mail address.
        .PARAMETER AccessToken
            This optional parameter AccessToken can be used if you want to use your own application with delegated or application permission. The parameter takes a previously acquired access token.
        .PARAMETER ByMail
            The parameter ByMail is a switch, which can be used in combination with Group, when an e-mail address instead of an id is used.
        .PARAMETER Filter
            The parameter Filter can be used, when you want to use a complex filter.
        .PARAMETER ShowProgress
            The parameter ShowProgress will show the progress of the script.
        .PARAMETER ReturnMembers
            Switch to return members of group.
        .PARAMETER ReturnMembersTransitive
            Switch to return transitive members of group.
        .PARAMETER Threads
            The parameter Threads defines how many Threads will be created. Only used in combination with MultiThread.
        .PARAMETER MultiThread
            The parameters MultiThread defines whether the script is running using multithreading.
        .PARAMETER Authority
            The authority from where you get the token.
        .PARAMETER ClientId
            Application ID of the registered app.
        .PARAMETER ClientSecret
            The secret, which is used for Client Credentials flow.
        .PARAMETER Certificate
            The certificate, which is used for Client Credentials flow.
        .PARAMETER MaxRetry
            How many retries for each user in case of error.
        .PARAMETER TimeoutSec
            TimeoutSec for Cmdlet Invoke-RestMethod.
        .PARAMETER MaxFilterResult
            MaxFilterResult when Filter is used.
        .EXAMPLE
            Get-MSGraphGroup -Group ServicesSales@bla.com -ByMail
            Get-MSGraphGroup -Group 6288514a-9840-4426-as05-d2955a03ea27
            Get-MSGraphGroup -Filter Get-MSGraphGroup -Filter "startswith(mail,'ServicesSale')"
        .NOTES
            If you want to use your own application make sure you have all the necessary minimum permission assigned: Group.Read.All (this might change in the future. Consult the full permission reference for Microsoft Graph)
        .LINK
            https://docs.microsoft.com/graph/api/resources/group?view=graph-rest-beta
            https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-auth-code-flow
            https://docs.microsoft.com/graph/paging
            https://docs.microsoft.com/graph/json-batching
            https://docs.microsoft.com/graph/query-parameters
            https://officeclient.microsoft.com/config16
            https://docs.microsoft.com/graph/permissions-reference
            https://docs.oasis-open.org/odata/odata/v4.0/errata03/os/complete/part2-url-conventions/odata-v4.0-errata03-os-part2-url-conventions-complete.html#_Toc453752358
    #>
    [CmdletBinding()]
    param(
        [parameter( Position=0)]
        [System.String[]]
        $Group,

        [parameter( Position=1)]
        [System.String]
        $AccessToken,

        [parameter( Position=2)]
        [System.Management.Automation.SwitchParameter]
        $ByMail,

        [parameter( Position=3)]
        [System.String]
        $Filter,

        [parameter( Position=4)]
        [System.Management.Automation.SwitchParameter]
        $ShowProgress,

        [parameter( Position=5)]
        [System.Management.Automation.SwitchParameter]
        $ReturnMembers,

        [parameter( Position=6)]
        [System.Management.Automation.SwitchParameter]
        $ReturnMembersTransitive,

        [parameter( Position=7)]
        [System.Int16]
        $Threads = '20',

        [parameter( Position=8)]
        [System.Management.Automation.SwitchParameter]
        $MultiThread,

        [parameter( Position=9)]
        [System.String]
        $Authority,

        [parameter( Position=10)]
        [System.String]
        $ClientId,

        [parameter( Position=11)]
        [System.String]
        $ClientSecret,

        [parameter( Position=12)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $Certificate,

        [parameter( Position=13)]
        [System.Int16]
        $MaxRetry = '3',

        [parameter( Position=14)]
        [System.Int32]
        $TimeoutSec = '15',

        [parameter( Position=15)]
        [System.Int32]
        $MaxFilterResult

    )

    begin
    {

        $timer = [System.Diagnostics.Stopwatch]::StartNew()

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $Error.Clear()

        $collection = [System.Collections.ArrayList]@()

        function Get-AccessTokenNoLibraries
        {
            <#
                .SYNOPSIS
                    This function acquires an access token using either shared secret or certificate.
                .DESCRIPTION
                    This function will acquire an access token from Azure AD using either shared secret or certificate and OAuth2.0 client credentials flow without the need of having any DLLs installed.
                .PARAMETER Authority
                    The parameter AZKeyVaultBaseUri is required and is the base uri of the Azure Key Vault.
                .PARAMETER ClientID
                    The parameter ClientID is required and defines the registered application.
                .PARAMETER ClientSecret
                    The parameter ClientSecret is optional and used for ClientCredential flow.
                .PARAMETER Certificate
                    The parameter Certificate is required, when a certificate is used and for ClientCredential flow. You need to provide a X509Certificate2 object.
                .EXAMPLE
                    Get-AccessTokenNoLibraries -Authority = 'https://login.microsoftonline.com/53g7676c-f466-423c-82f6-vb2f55791wl7/oauth2/v2.0/token' -ClientID '1521a4b6-v487-4078-be14-c9fvb17f8ab0' -Certificate $(dir Cert:\CurrentUser\My\ | ? thumbprint -eq 'F5B5C6135719644F5582BC184B4C2239D5112C73')
                    Get-AccessTokenNoLibraries -Authority = 'https://login.microsoftonline.com/53g7676c-f466-423c-82f6-vb2f55791wl7/oauth2/v2.0/token' -ClientID '1521a4b6-v487-4078-be14-c9fvb17f8ab0' -ClientSecret 'gavuCG5Pm__cG4F1~SN_03KQDAN2N~7o.L'
                .NOTES
                    
                .LINK
                    https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-certificate-credentials
                    https://adamtheautomator.com/microsoft-graph-api-powershell/#Acquire_an_Access_Token_(Using_a_Certificate)
                    https://adamtheautomator.com/microsoft-graph-api-powershell/#Acquire_an_Access_Token_(Application_Id_and_Secret)
            #>

            [CmdletBinding()]
            Param
            (
                [System.String]
                $Authority,

                [System.String]
                $ClientId,

                [System.String]
                $ClientSecret,

                [System.Security.Cryptography.X509Certificates.X509Certificate2]
                $Certificate
            )

            begin
            {
                $Scope = "https://graph.microsoft.com/.default"

                if ($Certificate)
                {

                    # Create base64 hash of certificate
                    $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())

                    # Create JWT timestamp for expiration
                    $StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
                    $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
                    $JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0)

                    # Create JWT validity start timestamp
                    $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
                    $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan,0)

                    # Create JWT header
                    $JWTHeader = @{
                        alg = 'RS256'
                        typ = 'JWT'
                        # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
                        x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='
                    }

                    # Create JWT payload
                    $JWTPayLoad = @{
                        # What endpoint is allowed to use this JWT
                        aud = $Authority # "https://login.microsoftonline.com/$TenantName/oauth2/token"

                        # Expiration timestamp
                        exp = $JWTExpiration

                        # Issuer = your application
                        iss = $ClientId

                        # JWT ID: random guid
                        jti = [System.Guid]::NewGuid()

                        # Not to be used before
                        nbf = $NotBefore

                        # JWT Subject
                        sub = $ClientId
                    }

                    # Convert header and payload to base64
                    $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
                    $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

                    $JWTPayLoadToByte =  [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
                    $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

                    # Join header and Payload with "." to create a valid (unsigned) JWT
                    $JWT = $EncodedHeader + "." + $EncodedPayload

                    # Get the private key object of your certificate
                    $PrivateKey = $Certificate.PrivateKey

                    # Define RSA signature and hashing algorithm
                    $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
                    $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

                    # Create a signature of the JWT
                    $Signature = [Convert]::ToBase64String(
                        $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT),$HashAlgorithm,$RSAPadding)
                    ) -replace '\+','-' -replace '/','_' -replace '='

                    # Join the signature to the JWT with "."
                    $JWT = $JWT + "." + $Signature

                    # Create a hash with body parameters
                    $Body = @{
                        client_id = $ClientId
                        client_assertion = $JWT
                        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                        scope = $Scope
                        grant_type = 'client_credentials'
                    
                    }

                    # Use the self-generated JWT as Authorization
                    $Header = @{
                        Authorization = "Bearer $JWT"
                    }

                }
                else
                {
                    # Create body
                    $Body = @{
                        client_id = $ClientId
                        client_secret = $ClientSecret
                        scope = $Scope
                        grant_type = 'client_credentials'
                    }

                }

                # Splat the parameters for Invoke-Restmethod for cleaner code
                $PostSplat = @{
                    ContentType = 'application/x-www-form-urlencoded'
                    Method = 'POST'
                    Body = $Body
                    Uri = $Authority
                }

                if ($Certificate)
                {
                    Write-Verbose 'Adding headers for certificate based request...'
                    $PostSplat.Add('Headers',$Header)
                }

            }

            process
            {
                try {
                    $accessToken = Invoke-RestMethod @PostSplat
                }
                catch{
                    $_
                }
            }

            end
            {
                $accessToken
            }
        }

        if ([System.String]::IsNullOrWhiteSpace($AccessToken))
        {
            # check for necessary parameters
            if ( (($Authority -and $ClientId) -and $ClientSecret) -or (($Authority -and $ClientId) -and $Certificate) )
            {
                Write-Verbose 'No AccessToken provided. Will acquire silently...'
                $acquireToken = @{
                    Authority = $Authority
                    ClientId = $ClientId
                    ClientSecret = $ClientSecret
                    Certificate = $Certificate
                }

                $AccessToken = (Get-AccessTokenNoLibraries @acquireToken).access_token

                $global:token = $AccessToken

                if ([System.String]::IsNullOrWhiteSpace($AccessToken))
                {
                    Write-Verbose "Couldn't acquire a token. Will stop now..."
                    break
                }

                Write-Verbose 'Setting accesstoken...'
                $psBoundParameters['AccessToken'] = $AccessToken
            }
            else
            {
                Write-Host 'No accesstoken provided and not enough data for acquiring one. Will exit...'
                break
            }
        }
        else
        {
            $global:token = $AccessToken
        }

        [System.Collections.ArrayList]$script:selectProperties = @(
            "allowExternalSenders",
            "assignedLicenses",
            "assignedLabels",
            "assignedLicenses",
            "autoSubscribeNewMembers",
            "classification",
            "createdByAppId",
            "createdDateTime",
            "deletedDateTime",
            "description",
            "displayName",
            "expirationDateTime",
            "groupTypes",
            "hideFromAddressLists",
            "hideFromOutlookClients",
            "id",
            "isSubscribedByMail",
            "licenseProcessingState",
            "mail",
            "mailEnabled",
            "mailNickname",
            "membershipRule",
            "membershipRuleProcessingState",
            "onPremisesDomainName",
            "onPremisesLastSyncDateTime",
            "onPremisesNetBiosName",
            "onPremisesProvisioningErrors",
            "onPremisesSamAccountName",
            "onPremisesSecurityIdentifier",
            "onPremisesSyncEnabled",
            "preferredDataLocation",
            "preferredLanguage",
            "proxyAddresses",
            "renewedDateTime",
            "resourceBehaviorOptions",
            "resourceProvisioningOptions",
            "securityEnabled",
            "securityIdentifier",
            "theme",
            "unseenConversationsCount",
            "unseenCount",
            "unseenMessagesCount",
            "visibility")

        if ($Filter)
        {

            [System.Boolean]$retryRequest = $true

            do {
                try {

                    Write-Verbose 'Found custom Filter. Will try to find group based on...'
                    $filterParams = @{
                        ContentType = 'application/json'
                        Method = 'GET'
                        Headers = @{ Authorization = "Bearer $($AccessToken)"}
                        Uri = 'https://graph.microsoft.com/beta/groups?$filter=' + $Filter
                        TimeoutSec = $TimeoutSec
                        ErrorAction = 'Stop'
                    }

                    $groupResponse = Invoke-RestMethod @filterParams

                    if ($groupResponse.'@odata.nextLink')
                    {
                        Write-Verbose 'Need to fetch more data...'
                        [System.Int16]$counter = '2'
                        # create collection
                        $groupCollection = [System.Collections.ArrayList]@()

                        # add first batch of groups to collection
                        if ($groupResponse.Value)
                        {
                            $groupCollection += $groupResponse.Value
                        }
                        else
                        {
                            $groupCollection += $groupResponse
                        }

                        do
                        {
                            $filterParams = @{
                                ContentType = 'application/json'
                                Method = 'GET'
                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                Uri = $($groupResponse.'@odata.nextLink')
                                TimeoutSec = $TimeoutSec
                                ErrorAction = 'Stop'
                                Verbose = $false
                            }

                            $groupResponse = Invoke-RestMethod @filterParams

                            $groupCollection += $groupResponse.Value

                            Write-Verbose "Pagecount:$($counter)..."
                            $counter++

                            if ( $MaxFilterResult -and ($groupCollection.Count -ge $MaxFilterResult))
                            {
                                Write-Verbose "MaxFilterResult reached. Will stop searching now..."
                                $groupResponse.'@odata.nextLink' = $null
                                $groupCollection = $groupCollection  | Select-Object -First $MaxFilterResult
                            }

                        } while ($groupResponse.'@odata.nextLink')

                        $Group = $groupCollection.id
                    }
                    else
                    {
                        $Group = $groupResponse.value.id
                    }

                    $retryRequest = $false

                    Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"

                }
                catch
                {
                    $_
                    break
                }
            } while ($retryRequest)

            if ($Group.count -eq 0)
            {
                Write-Verbose $('No group found for filter "' + $($Filter) + '"! Terminate now...')
                break
            }
            else
            {
                Write-Host "Found $($Group.count) groups for provided filter..."
            }
        }

        if ($MultiThread)
        {
            #initiate runspace and make sure we are using single-threaded apartment STA
            $Jobs = @()
            $Sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $Threads, $Sessionstate, $Host)
            $RunspacePool.ApartmentState = 'STA'
            $RunspacePool.Open()
        }

        # initiate counter
        [System.Int32]$j = '1'

    }

    process
    {
        if ($MultiThread)
        {
            Write-Verbose "Will run multithreaded...GroupCount:$($Group.count)"
            #create scriptblock from function
            $scriptBlock = [System.Management.Automation.ScriptBlock]::Create((Get-ChildItem Function:\Get-MSGraphGroup).Definition)
            
            ForEach ($account in $Group)
            {
                 $progressParams = @{
                    Activity = "Adding group to jobs and starting job - $($account)"
                    PercentComplete = $j / $Group.count * 100
                    Status = "Remaining: $($Group.count - $j) out of $($Group.count)"
                 }
                 Write-Progress @progressParams

                 $j++

                 try {
                    $PowershellThread = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock).AddParameter('Group',$account)
                    [void]$PowershellThread.AddParameter('AccessToken',$AccessToken)
                    [void]$PowershellThread.AddParameter('ByMail',$ByMail)
                    [void]$PowershellThread.AddParameter('Authority',$Authority)
                    [void]$PowershellThread.AddParameter('ClientId',$ClientId)
                    [void]$PowershellThread.AddParameter('ClientSecret',$ClientSecret)
                    [void]$PowershellThread.AddParameter('Certificate',$Certificate)
                    [void]$PowershellThread.AddParameter('MaxRetry',$MaxRetry)
                    [void]$PowershellThread.AddParameter('TimeoutSec',$TimeoutSec)
                    if ($MaxFilterResult)
                    {
                        [void]$PowershellThread.AddParameter('MaxFilterResult',$MaxFilterResult)
                    }
                    [void]$PowershellThread.AddParameter('ShowProgress',$false)

                    $PowershellThread.RunspacePool = $RunspacePool
                    $Handle = $PowershellThread.BeginInvoke()
                    $Job = "" | Select-Object Handle, Thread, Object
                    $Job.Handle = $Handle
                    $Job.Thread = $PowershellThread
                    $Job.Object = $acount
                    $Jobs += $Job
                 }
                 catch {
                    $_
                 }
            }

            $progressParams.Status = "Ready"
            $progressParams.Add('Completed',$true)
            Write-Progress @progressParams
        }
        else
        {
            foreach($account in $Group)
            {
                [System.Int16]$retryCount = '0'
                [System.Boolean]$keepTrying = $true

                do {

                    try {
                        $processingTime = [System.Diagnostics.Stopwatch]::StartNew()

                        if ($ShowProgress)
                        {
                            $progressParams = @{
                                Activity = "Processing group - $($account)"
                                PercentComplete = $j / $Group.count * 100
                                Status = "Remaining: $($Group.count - $j) out of $($Group.count)"
                            }
                            Write-Progress @progressParams
                        }

                        $j++

                        # get group id
                        if ($ByMail)
                        {
                            Write-Verbose 'Get group by email...'
                            $byMailParams = @{
                                Uri = "https://graph.microsoft.com/beta/groups?filter=mail eq '$($account)'"
                                Method = 'GET'
                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                ErrorAction = 'Stop'
                            }

                            $id = (Invoke-RestMethod @byMailParams).value.id

                            if ([System.String]::IsNullOrWhiteSpace($id))
                            {
                                Write-Verbose "No group found for e-mail $($account) ..."
                                break
                            }
                        }
                        else
                        {
                            Write-Verbose 'Get group by id...'
                            $id = $account
                        }

                        $body = @{
                            requests = @(
                                @{
                                    url = "/groups/$id" + '?$select=' + $($selectProperties -join ',')
                                    method = 'GET'
                                    id = '1'
                                },
                                @{
                                    url = "/groups/$id/owners"
                                    method = 'GET'
                                    id = '2'
                                },
                                @{
                                    url = "/groups/$id/sites/root"
                                    method = 'GET'
                                    id = '3'
                                }
                            )
                        }

                        if ($ReturnMembers)
                        {
                            $members = @{
                                    url = "/groups/$id/members"
                                    method = 'GET'
                                    id = '4'
                            }

                        }
                        elseif ($ReturnMembersTransitive)
                        {
                            $members = @{
                                    url = "/groups/$id/transitiveMembers"
                                    method = 'GET'
                                    id = '4'
                                    #headers = @{"ConsistencyLevel" = "eventual"}
                            }

                        }
                        else
                        {
                            $members = @{
                                    #url = "/groups/$id/members/" + '$count'
                                    url = "/groups/$id/transitiveMembers/" + '$count'
                                    method = 'GET'
                                    id = '4'
                                    headers = @{"ConsistencyLevel"="eventual"}
                            }
                        }

                        $body.requests += $members

                        $restParams = @{
                            ContentType = 'application/json'
                            Method = 'POST'
                            Headers = @{ Authorization = "Bearer $($AccessToken)"}
                            Body = $body | ConvertTo-Json -Depth 4
                            Uri = 'https://graph.microsoft.com/beta/$batch'
                            TimeoutSec = $TimeoutSec
                            #ErrorAction = 'Stop'
                        }

                        $global:data = Invoke-RestMethod @restParams

                        # create custom object
                        $groupInfo = $null
                        # check for error
                        $responseGroup = $data.responses | Where-Object -FilterScript { $_.id -eq 1}
                        $groupInfo = $responseGroup.Body | Select-Object * -ExcludeProperty "@odata.context"

                        if (($responseGroup.status -ne 200) -and ('MailboxNotEnabledForRESTAPI|UnsupportedQueryOption|AppOnlyAccessNotEnabledForTarget|ErrorAccessDenied' -match $responseGroup.body.error.code))
                        {
                            Write-Verbose "Error MailboxNotEnabledForRESTAPI|UnsupportedQueryOption thrown. Will try again without certain properties..."

                            # create list with unsupported properties
                            [System.Collections.ArrayList]$unsupportedProperties = @(
                                "allowExternalSenders",
                                "autoSubscribeNewMembers",
                                "hideFromAddressLists",
                                "hideFromOutlookClients",
                                "isSubscribedByMail",
                                "unseenConversationsCount",
                                "unseenCount",
                                "unseenMessagesCount")

                            # remove unsupported properties
                            foreach ($prop in $unsupportedProperties)
                            {
                                $selectProperties.Remove($prop)
                            }

                            # set the new URL and body
                            $body.requests[0].url = "/groups/$id" + '?$select=' + $($selectProperties -join ',')
                            $restParams.Body = $body | ConvertTo-Json -Depth 4

                            $global:data = Invoke-RestMethod @restParams
                            $groupInfo = ($data.responses | Where-Object -FilterScript { $_.id -eq 1}).Body | Select-Object * -ExcludeProperty "@odata.context"
                        }

                        $groupProperties = $groupInfo | Get-Member -MemberType NoteProperty
                        $groupObject = New-Object -TypeName psobject

                        foreach ($property in $groupProperties)
                        {
                            $groupObject | Add-Member -MemberType NoteProperty -Name $( $property.Name ) -Value $( $groupInfo.$( $property.Name ) )
                        }

                        # add owners to object
                        $responseOwner = $data.responses | Where-Object -FilterScript {$_.id -eq 2}

                        if ('200' -eq $responseOwner.status)
                        {
                            $groupObject | Add-Member -MemberType NoteProperty -Name Owners -Value @( $($responseOwner.body.value | Select-Object * -ExcludeProperty "@odata.type"))
                        }
                        else
                        {
                            Write-Verbose "Error found in response for owners..."
                            $groupObject | Add-Member -MemberType NoteProperty -Name Owners -Value @( $($responseOwner.body.error))
                        }

                        # add members to object
                        $responseMember = $data.responses | Where-Object -FilterScript {$_.id -eq 4}

                        if ('200' -eq $responseMember.status)
                        {
                            if ((-not $ReturnMembers) -and (-not $ReturnMembersTransitive))
                            {
                                $groupObject | Add-Member -MemberType NoteProperty -Name MemberCount -Value $($responseMember.body)
                            }
                            else
                            {
                                if ($responseMember.body.'@odata.nextLink')
                                {

                                    Write-Verbose 'Need to fetch more data for members...'
                                    [System.Int16]$counter = '2'
                                    # create collection
                                    $memberCollection = [System.Collections.ArrayList]@()

                                    # add first batch of members to collection
                                    $memberCollection += $responseMember.body.value | Select-Object * -ExcludeProperty "@odata.type"

                                    do
                                    {
                                        if ($responseMember.body.'@odata.nextLink')
                                        {
                                            $URI = $responseMember.body.'@odata.nextLink'
                                        }
                                        else
                                        {
                                            $URI = $responseMember.'@odata.nextLink'
                                        }

                                        $groupParams = @{
                                                ContentType = 'application/json'
                                                Method = 'GET'
                                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                                Uri = $URI
                                                TimeoutSec = $TimeoutSec
                                                Verbose = $false
                                            }

                                        $responseMember = Invoke-RestMethod @groupParams

                                        if ($responseMember.body.value)
                                        {
                                            $memberCollection += $responseMember.body.value | Select-Object * -ExcludeProperty "@odata.type"
                                        }
                                        else
                                        {
                                            $memberCollection += $responseMember.value | Select-Object * -ExcludeProperty "@odata.type"
                                        }

                                        Write-Verbose "Pagecount:$($counter)..."
                                        $counter++

                                    } while ($responseMember.'@odata.nextLink')
                                    
                                    $groupObject | Add-Member -MemberType NoteProperty -Name Members -Value @( $memberCollection )
                                }
                                else
                                {
                                    $groupObject | Add-Member -MemberType NoteProperty -Name Members -Value @( $($responseMember.body.value | Select-Object * -ExcludeProperty "@odata.type"))
                                }
                            }
                        }
                        else
                        {
                            Write-Verbose "Error found in response for members..."
                            $groupObject | Add-Member -MemberType NoteProperty -Name Members -Value @( $($responseMember.body.error))
                        }

                        # add root site to object
                        $responseSite = $data.responses | Where-Object -FilterScript {$_.id -eq 3}

                        if ('200' -eq $responseSite.status)
                        {
                            $groupObject | Add-Member -MemberType NoteProperty -Name Sites -Value @( $($responseSite.body | Select-Object * -ExcludeProperty "@odata.context"))
                        }
                        else
                        {
                            Write-Verbose "Error found in response for sites..."
                            $groupObject | Add-Member -MemberType NoteProperty -Name Sites -Value @( $($responseSite.body.error))
                        }

                        $collection += $groupObject

                        if ($ShowProgress)
                        {
                            $progressParams.Status = "Ready"
                            $progressParams.Add('Completed',$true)
                            Write-Progress @progressParams
                        }

                        $processingTime.Stop()
                        $keepTrying = $false
                        Write-Verbose "Group processing time:$($processingTime.Elapsed.ToString())"

                    }
                    catch {

                        $statusCode = $_.Exception.Response.StatusCode.value__

                        if ('401' -eq $statusCode)
                        {

                            Write-Verbose "HTTP error $($statusCode) thrown. Will try to acquire new access token..."
                            $acquireToken = @{
                                Authority = $Authority
                                ClientId = $ClientId
                                ClientSecret = $ClientSecret
                                Certificate = $Certificate
                            }

                            $AccessToken = (Get-AccessTokenNoLibraries @acquireToken).access_token

                            if ([System.String]::IsNullOrWhiteSpace($AccessToken))
                            {
                                Write-Verbose "Couldn't acquire a token. Will stop now..."

                                $groupObject = New-Object -TypeName psobject
                                $groupObject | Add-Member -MemberType NoteProperty -Name GroupID -Value $account
                                $groupObject | Add-Member -MemberType NoteProperty -Name Error -Value 'Could not acquire access token'
                                $collection += $groupObject
                                break
                            }

                            Write-Verbose 'Setting accesstoken...'
                            $psBoundParameters['AccessToken'] = $AccessToken

                            # increase retrycount
                            $retryCount++
                            Write-Verbose "Retrycount:$($retryCount)"

                        }
                        elseif ('403' -eq $statusCode)
                        {
                            Write-Verbose "Error for resource $($_.Exception.Response.ResponseUri.ToString()) with StatusCode $($_.Exception.Response.StatusCode.ToString())"
                            # increase retrycount
                            $retryCount++
                        }
                        elseif ('429' -eq $statusCode)
                        {

                            Write-Verbose "HTTP error $($statusCode) thrown. Will backoff..."
                            Start-Sleep -Seconds 2

                        }
                        elseif ('503' -eq $statusCode)
                        {

                            Write-Verbose "HTTP error $($statusCode) thrown. Will retry..."
                            # increase retrycount
                            $retryCount++
                            Write-Verbose "Retrycount:$($retryCount)"

                        }
                        else
                        {
                            $groupObject = New-Object -TypeName psobject
                            $groupObject | Add-Member -MemberType NoteProperty -Name GroupID -Value $account
                            $groupObject | Add-Member -MemberType NoteProperty -Name Error -Value $_.Exception
                            $collection += $groupObject
                            Write-Verbose "Error occured for group $($account)..."
                            break
                        }

                        if ($retryCount -eq $MaxRetry)
                        {
                            Write-Verbose 'MaxRetries done. Skip this group now...'
                            $keepTrying = $false

                            $groupObject = New-Object -TypeName psobject
                            $groupObject | Add-Member -MemberType NoteProperty -Name GroupID -Value $account
                            $groupObject | Add-Member -MemberType NoteProperty -Name Error -Value $_.Exception
                            $collection += $groupObject
                        }
                    }

                } while ($keepTrying)
            }
        }

    }

    end
    {
        if ($MultiThread)
        {
            $SleepTimer = 1000
            While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0) {

                Write-Progress `
                    -id 1 `
                    -Activity "Waiting for Jobs - $($Threads - $($RunspacePool.GetAvailableRunspaces())) of $Threads threads running" `
                    -PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
                    -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining out of $($Jobs.Count)"

                ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $true})) {
                    $Job.Thread.EndInvoke($Job.Handle)
                    $Job.Thread.Dispose()
                    $Job.Thread = $Null
                    $Job.Handle = $Null
                }

                # kill all incomplete threads when hit "CTRL+q"
                if ($Host.UI.RawUI.KeyAvailable) {
                    $KeyInput = $Host.UI.RawUI.ReadKey("IncludeKeyUp,NoEcho")
                    if (($KeyInput.ControlKeyState -cmatch '(Right|Left)CtrlPressed') -and ($KeyInput.VirtualKeyCode -eq '81')) {
                        Write-Host -fore red "Kill all incomplete threads..."
                            ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})) {
                                Write-Host -fore yellow "Stopping job $($Job.Object)..."
                                $Job.Thread.Stop()
                                $Job.Thread.Dispose()
                             }
                        Write-Host -fore red "All jobs terminated..."
                        break
                    }
                }

                Start-Sleep -Milliseconds $SleepTimer

            }

            # clean-up
            Write-Verbose 'Perform cleanup of runspaces...'
            $RunspacePool.Close() | Out-Null
            $RunspacePool.Dispose() | Out-Null
            [System.GC]::Collect()
        }
        else
        {
            $collection
        }

        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Get-MSGraphUser
{
    <#
        .SYNOPSIS
            This function retrieves properties of a user.
        .DESCRIPTION
            This function retrieves properties and additional information e.g.: authentication methods, serviceprincipal of a user.
        .PARAMETER User
            The parameter User defines the user to be queried.
        .PARAMETER AccessToken
            This required parameter AccessToken takes the Bearer access token for authentication the requests. The parameter takes a previously acquired access token.
        .PARAMETER GetMailboxSettings
            The parameter GetMailboxSettings can be used to retrieve MailboxSettings for a mailbox and is an optional switch.
        .PARAMETER GetDeltaToken
            The parameter GetDeltaToken returns a DeltaToken for monitoring changes.
        .PARAMETER Filter
            The parameter Filter can be used, when you want to use a complex filter.
        .PARAMETER GetAuthMethods
            The parameter GetAuthMethods returns current authentication methods for a user.
        .PARAMETER ReturnDeletedUsers
            The parameter ReturnDeletedUsers forces the script to query deleted users.
        .PARAMETER ShowProgress
            The parameter ShowProgress will show the progress of the script.
        .PARAMETER Threads
            The parameter Threads defines how many Threads will be created. Only used in combination with MultiThread.
        .PARAMETER MultiThread
            The parameters MultiThread defines whether the script is running using multithreading.
        .PARAMETER Authority
            The authority from where you get the token.
        .PARAMETER ClientId
            Application ID of the registered app.
        .PARAMETER ClientSecret
            The secret, which is used for Client Credentials flow.
        .PARAMETER Certificate
            The certificate, which is used for Client Credentials flow.
        .PARAMETER MaxRetry
            How many retries for each user in case of error.
        .PARAMETER TimeoutSec
            TimeoutSec for Cmdlet Invoke-RestMethod.
        .PARAMETER MinAttributeSet
            Request will be done only for default returned attributes
        .PARAMETER SetAdvancedQueryProps
            Set ConsistencyLevel header and append &$count=true to filter for advanced query
        .EXAMPLE
            # retrieve data for a specific user by providing a previoues retreieved Bearer access token
            Get-MSGraphUser -User ingo.gegenwarth@contoso.com -GetMailboxSettings -GetDeltaToken -GetAuthMethods -AccessToken eyJ0eXAiOiJKV1QiLC...AwNzY5OTE4LCJuYmYiO -Verbose
            # retrieve data for a specific user and acquire an access token with ClientCredentials flow, where a certificate is used for client credential
            Get-MSGraphUser -User ingo.gegenwarth@contoso.com -Authority "https://login.microsoftonline.com/42f...791af7/oauth2/v2.0/token" -ClientId 1961a...2f8ab0 -Certificate $(dir Cert:\CurrentUser\My\ | ? thumbprint -eq 'F5A9C6...C2239D5112C73')
            # retrive data for multiple users using multithreading and filter
            Get-MSGraphUser -Filter "StartsWith(mail,'ingo')" -MultiThread -Authority "https://login.microsoftonline.com/42f...791af7/oauth2/v2.0/token" -ClientId 1961a...2f8ab0 -Certificate $(dir Cert:\CurrentUser\My\ | ? thumbprint -eq 'F5A9C6...C2239D5112C73')
        .NOTES
            If you want to leverage all functionality you will need to provide an access token with the following claims:
                Directory.Read.All
                Group.Read.All
                MailboxSettings.Read
                User.Read.All
                UserAuthenticationMethod.Read.All
        .LINK
            https://docs.microsoft.com/graph/api/resources/user?view=graph-rest-beta
            https://docs.microsoft.com/graph/paging
            https://docs.microsoft.com/graph/json-batching
            https://docs.microsoft.com/graph/query-parameters
            https://docs.microsoft.com/graph/permissions-reference
            https://docs.oasis-open.org/odata/odata/v4.0/errata03/os/complete/part2-url-conventions/odata-v4.0-errata03-os-part2-url-conventions-complete.html#_Toc453752358
            https://docs.microsoft.com/en-us/graph/throttling
            https://docs.microsoft.com/graph/errors
            https://docs.microsoft.com/graph/best-practices-concept#handling-expected-errors
    #>
    [CmdletBinding()]
    param(
        [parameter( Position=0)]
        [System.String[]]
        $User,

        [parameter( Position=1)]
        [System.String]
        $AccessToken,

        [parameter( Position=2)]
        [System.Management.Automation.SwitchParameter]
        $GetMailboxSettings,

        [parameter( Position=3)]
        [System.Management.Automation.SwitchParameter]
        $GetDeltaToken,

        [parameter( Position=4)]
        [System.String]
        $Filter,

        [parameter( Position=5)]
        [System.Management.Automation.SwitchParameter]
        $ListAll,

        [parameter( Position=6)]
        [System.Management.Automation.SwitchParameter]
        $GetAuthMethods,

        [parameter( Position=7)]
        [System.Management.Automation.SwitchParameter]
        $ReturnDeletedUsers,

        [parameter( Position=8)]
        [System.Management.Automation.SwitchParameter]
        $ShowProgress,

        [parameter( Position=9)]
        [System.Int16]
        $Threads = '20',

        [parameter( Position=10)]
        [System.Management.Automation.SwitchParameter]
        $MultiThread,

        [parameter( Position=11)]
        [System.String]
        $Authority,

        [parameter( Position=12)]
        [System.String]
        $ClientId,

        [parameter( Position=13)]
        [System.String]
        $ClientSecret,

        [parameter( Position=14)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $Certificate,

        [parameter( Position=15)]
        [System.Int16]
        $MaxRetry = '3',

        [parameter( Position=16)]
        [System.Int32]
        $TimeoutSec = '15',

        [parameter( Position=17)]
        [System.Int32]
        $MaxFilterResult,

        [parameter( Position=18)]
        [System.Management.Automation.SwitchParameter]
        $MinAttributeSet,

        [parameter( Position=19)]
        [System.Management.Automation.SwitchParameter]
        $SetAdvancedQueryProps

    )

    begin
    {

        $timer = [System.Diagnostics.Stopwatch]::StartNew()

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $Error.Clear()

        $collection = [System.Collections.ArrayList]@()

        function Get-AccessTokenNoLibraries
        {
            <#
                .SYNOPSIS
                    This function acquires an access token using either shared secret or certificate.
                .DESCRIPTION
                    This function will acquire an access token from Azure AD using either shared secret or certificate and OAuth2.0 client credentials flow without the need of having any DLLs installed.
                .PARAMETER Authority
                    The parameter AZKeyVaultBaseUri is required and is the base uri of the Azure Key Vault.
                .PARAMETER ClientID
                    The parameter ClientID is required and defines the registered application.
                .PARAMETER ClientSecret
                    The parameter ClientSecret is optional and used for ClientCredential flow.
                .PARAMETER Certificate
                    The parameter Certificate is required, when a certificate is used and for ClientCredential flow. You need to provide a X509Certificate2 object.
                .EXAMPLE
                    Get-AccessTokenNoLibraries -Authority = 'https://login.microsoftonline.com/53g7676c-f466-423c-82f6-vb2f55791wl7/oauth2/v2.0/token' -ClientID '1521a4b6-v487-4078-be14-c9fvb17f8ab0' -Certificate $(dir Cert:\CurrentUser\My\ | ? thumbprint -eq 'F5B5C6135719644F5582BC184B4C2239D5112C73')
                    Get-AccessTokenNoLibraries -Authority = 'https://login.microsoftonline.com/53g7676c-f466-423c-82f6-vb2f55791wl7/oauth2/v2.0/token' -ClientID '1521a4b6-v487-4078-be14-c9fvb17f8ab0' -ClientSecret 'gavuCG5Pm__cG4F1~SN_03KQDAN2N~7o.L'
                .NOTES
                    
                .LINK
                    https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-certificate-credentials
                    https://adamtheautomator.com/microsoft-graph-api-powershell/#Acquire_an_Access_Token_(Using_a_Certificate)
                    https://adamtheautomator.com/microsoft-graph-api-powershell/#Acquire_an_Access_Token_(Application_Id_and_Secret)
            #>

            [CmdletBinding()]
            Param
            (
                [System.String]
                $Authority,

                [System.String]
                $ClientId,

                [System.String]
                $ClientSecret,

                [System.Security.Cryptography.X509Certificates.X509Certificate2]
                $Certificate
            )

            begin
            {
                $Scope = "https://graph.microsoft.com/.default"

                if ($Certificate)
                {

                    # Create base64 hash of certificate
                    $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())

                    # Create JWT timestamp for expiration
                    $StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
                    $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
                    $JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0)

                    # Create JWT validity start timestamp
                    $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
                    $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan,0)

                    # Create JWT header
                    $JWTHeader = @{
                        alg = 'RS256'
                        typ = 'JWT'
                        # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
                        x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='
                    }

                    # Create JWT payload
                    $JWTPayLoad = @{
                        # What endpoint is allowed to use this JWT
                        aud = $Authority # "https://login.microsoftonline.com/$TenantName/oauth2/token"

                        # Expiration timestamp
                        exp = $JWTExpiration

                        # Issuer = your application
                        iss = $ClientId

                        # JWT ID: random guid
                        jti = [System.Guid]::NewGuid()

                        # Not to be used before
                        nbf = $NotBefore

                        # JWT Subject
                        sub = $ClientId
                    }

                    # Convert header and payload to base64
                    $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
                    $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

                    $JWTPayLoadToByte =  [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
                    $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

                    # Join header and Payload with "." to create a valid (unsigned) JWT
                    $JWT = $EncodedHeader + "." + $EncodedPayload

                    # Get the private key object of your certificate
                    $PrivateKey = $Certificate.PrivateKey

                    # Define RSA signature and hashing algorithm
                    $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
                    $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

                    # Create a signature of the JWT
                    $Signature = [Convert]::ToBase64String(
                        $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT),$HashAlgorithm,$RSAPadding)
                    ) -replace '\+','-' -replace '/','_' -replace '='

                    # Join the signature to the JWT with "."
                    $JWT = $JWT + "." + $Signature

                    # Create a hash with body parameters
                    $Body = @{
                        client_id = $ClientId
                        client_assertion = $JWT
                        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                        scope = $Scope
                        grant_type = 'client_credentials'
                    
                    }

                    # Use the self-generated JWT as Authorization
                    $Header = @{
                        Authorization = "Bearer $JWT"
                    }

                }
                else
                {
                    # Create body
                    $Body = @{
                        client_id = $ClientId
                        client_secret = $ClientSecret
                        scope = $Scope
                        grant_type = 'client_credentials'
                    }

                }

                # Splat the parameters for Invoke-Restmethod for cleaner code
                $PostSplat = @{
                    ContentType = 'application/x-www-form-urlencoded'
                    Method = 'POST'
                    Body = $Body
                    Uri = $Authority
                }

                if ($Certificate)
                {
                    Write-Verbose 'Adding headers for certificate based request...'
                    $PostSplat.Add('Headers',$Header)
                }

            }

            process
            {
                try {
                    $accessToken = Invoke-RestMethod @PostSplat
                }
                catch{
                    $_
                }
            }

            end
            {
                $accessToken
            }
        }

        if ([System.String]::IsNullOrWhiteSpace($AccessToken))
        {
            # check for necessary parameters
            if ( (($Authority -and $ClientId) -and $ClientSecret) -or (($Authority -and $ClientId) -and $Certificate) )
            {
                Write-Verbose 'No AccessToken provided. Will acquire silently...'
                $acquireToken = @{
                    Authority = $Authority
                    ClientId = $ClientId
                    ClientSecret = $ClientSecret
                    Certificate = $Certificate
                }

                $AccessToken = (Get-AccessTokenNoLibraries @acquireToken).access_token

                $global:token = $AccessToken

                if ([System.String]::IsNullOrWhiteSpace($AccessToken))
                {
                    Write-Verbose "Couldn't acquire a token. Will stop now..."
                    break
                }

                Write-Verbose 'Setting accesstoken...'
                $psBoundParameters['AccessToken'] = $AccessToken
            }
            else
            {
                Write-Host 'No accesstoken provided and not enough data for acquiring one. Will exit...'
                break
            }
        }

        [System.Collections.ArrayList]$script:selectProperties = @(
            "aboutMe",
            "accountEnabled",
            "ageGroup",
            "assignedLicenses",
            "assignedPlans",
            "birthday",
            "businessPhones",
            "city",
            "companyName",
            "consentProvidedForMinor",
            "country",
            "createdDateTime",
            "creationType",
            "deletedDateTime",
            "department",
            "displayName",
            "employeeId",
            "externalUserState",
            "externalUserStateChangeDateTime",
            "faxNumber",
            "givenName",
            "hireDate",
            "id",
            "identities",
            "imAddresses",
            "infoCatalogs",
            "interests",
            #"isResourceAccount",
            "jobTitle",
            "lastPasswordChangeDateTime",
            "legalAgeGroupClassification",
            "licenseAssignmentStates",
            "mail",
            "mailNickname",
            "mobilePhone",
            "mySite",
            "officeLocation",
            "onPremisesDistinguishedName",
            "onPremisesDomainName",
            "onPremisesExtensionAttributes",
            "onPremisesImmutableId",
            "onPremisesLastSyncDateTime",
            "onPremisesProvisioningErrors",
            "onPremisesSamAccountName",
            "onPremisesSecurityIdentifier",
            "onPremisesSyncEnabled",
            "onPremisesUserPrincipalName",
            "otherMails",
            "passwordPolicies",
            "passwordProfile",
            "pastProjects",
            "postalCode",
            "preferredDataLocation",
            "preferredLanguage",
            "preferredName",
            "provisionedPlans",
            "proxyAddresses",
            "refreshTokensValidFromDateTime",
            "responsibilities",
            "schools",
            "showInAddressList",
            "signInSessionsValidFromDateTime",
            "skills",
            #"signInActivity",
            "state",
            "streetAddress",
            "surname",
            "usageLocation",
            "userPrincipalName",
            "userType")

        if ($Filter -or $ListAll)
        {

            [System.Boolean]$retryRequest = $true

            do {
                try {

                    Write-Verbose 'Found custom Filter. Will try to find user based on...'
                    if ($ReturnDeletedUsers)
                    {
                        if ($ListAll)
                        {
                            $URI = 'https://graph.microsoft.com/beta/directory/deletedItems/microsoft.graph.user'
                        }
                        else
                        {
                            $URI = 'https://graph.microsoft.com/beta/directory/deletedItems/microsoft.graph.user?$filter='
                        }
                    }
                    else
                    {
                        if ($ListAll)
                        {
                            $URI = 'https://graph.microsoft.com/beta/users'
                        }
                        else
                        {
                            $URI = 'https://graph.microsoft.com/beta/users?$filter='
                            $URI = $URI + $Filter
                        }
                    }

                    $filterParams = @{
                        ContentType = 'application/json'
                        Method = 'GET'
                        Headers = @{ Authorization = "Bearer $($AccessToken)"}
                        Uri = $URI
                        TimeoutSec = $TimeoutSec
                        ErrorAction = 'Stop'
                    }

                    if ($SetAdvancedQueryProps)
                    {
                        $filterParams.Headers.Add('ConsistencyLevel','eventual')
                        $filterParams.Uri = $filterParams.Uri + '&$count=true'
                    }

                    $userResponse = Invoke-RestMethod @filterParams

                    if ($userResponse.'@odata.nextLink')
                    {
                        Write-Verbose 'Need to fetch more data...'
                        [System.Int16]$counter = '2'
                        # create collection
                        $userCollection = [System.Collections.ArrayList]@()

                        # add first batch of groups to collection
                        if ($userResponse.Value)
                        {
                            $userCollection += $userResponse.Value
                        }
                        else
                        {
                            $userCollection += $userResponse
                        }

                        do
                        {
                            $filterParams = @{
                                ContentType = 'application/json'
                                Method = 'GET'
                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                Uri = $($userResponse.'@odata.nextLink')
                                TimeoutSec = $TimeoutSec
                                ErrorAction = 'Stop'
                                Verbose = $false
                            }

                            $userResponse = Invoke-RestMethod @filterParams

                            $userCollection += $userResponse.Value

                            Write-Verbose "Pagecount:$($counter)..."
                            $counter++

                            if ( $MaxFilterResult -and ($userCollection.Count -ge $MaxFilterResult))
                            {
                                Write-Verbose "MaxFilterResult reached. Will stop searching now..."
                                $userResponse.'@odata.nextLink' = $null
                                $userCollection = $userCollection  | Select-Object -First $MaxFilterResult
                            }

                        } while ($userResponse.'@odata.nextLink')

                        $User = $userCollection.id
                    }
                    else
                    {
                        $User = $userResponse.value.id
                    }

                    $retryRequest = $false

                    Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"

                }
                catch
                {
                    $_
                    break
                }
            } while ($retryRequest)

            if ($User.count -eq 0)
            {
                Write-Verbose $('No user found for filter "' + $($Filter) + '"! Terminate now...')
            }
            else
            {
                Write-Host "Found $($user.count) users for provided filter..."
            }
        }

        if ($MultiThread)
        {
            #initiate runspace and make sure we are using single-threaded apartment STA
            $Jobs = @()
            $Sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $Threads, $Sessionstate, $Host)
            $RunspacePool.ApartmentState = 'STA'
            $RunspacePool.Open()
        }

        # initiate counter
        [System.Int32]$j = '1'

    }

    process
    {
        if ($MultiThread)
        {
            Write-Verbose 'Will run multithreaded...'
            #create scriptblock from function
            $scriptBlock = [System.Management.Automation.ScriptBlock]::Create((Get-ChildItem Function:\Get-MSGraphUser).Definition)
            
            ForEach ($account in $User)
            {
                 $progressParams = @{
                    Activity = "Adding user to jobs and starting job - $($account)"
                    PercentComplete = $j / $User.count * 100
                    Status = "Remaining: $($User.count - $j) out of $($User.count)"
                 }
                 Write-Progress @progressParams

                 $j++

                 try {
                    $PowershellThread = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock).AddParameter('User',$account)
                    [void]$PowershellThread.AddParameter('AccessToken',$AccessToken)
                    [void]$PowershellThread.AddParameter('GetMailboxSettings',$GetMailboxSettings)
                    [void]$PowershellThread.AddParameter('GetDeltaToken',$GetDeltaToken)
                    [void]$PowershellThread.AddParameter('GetAuthMethods',$GetAuthMethods)
                    [void]$PowershellThread.AddParameter('ReturnDeletedUsers',$ReturnDeletedUsers)
                    [void]$PowershellThread.AddParameter('Authority',$Authority)
                    [void]$PowershellThread.AddParameter('ClientId',$ClientId)
                    [void]$PowershellThread.AddParameter('ClientSecret',$ClientSecret)
                    [void]$PowershellThread.AddParameter('Certificate',$Certificate)
                    [void]$PowershellThread.AddParameter('MaxRetry',$MaxRetry)
                    [void]$PowershellThread.AddParameter('TimeoutSec',$TimeoutSec)
                    [void]$PowershellThread.AddParameter('ShowProgress',$false)

                    $PowershellThread.RunspacePool = $RunspacePool
                    $Handle = $PowershellThread.BeginInvoke()
                    $Job = "" | Select-Object Handle, Thread, Object
                    $Job.Handle = $Handle
                    $Job.Thread = $PowershellThread
                    $Job.Object = $account
                    $Jobs += $Job
                 }
                 catch {
                    $_
                 }
            }

            $progressParams.Status = "Ready"
            $progressParams.Add('Completed',$true)
            Write-Progress @progressParams

        }
        else
        {

            foreach ($account in $User)
            {

                [System.Int16]$retryCount = '0'
                [System.Boolean]$keepTrying = $true

                do {

                    try {

                        $processingTime = [System.Diagnostics.Stopwatch]::StartNew()

                        if ($ShowProgress)
                        {
                            $progressParams = @{
                                Activity = "Processing user - $($account)"
                                PercentComplete = $j / $User.count * 100
                                Status = "Remaining: $($User.count - $j) out of $($User.count)"
                            }
                            Write-Progress @progressParams
                        }

                        $j++

                        if ($ReturnDeletedUsers)
                        {
                            # create list with unsupported properties for deleted users
                            [System.Collections.ArrayList]$unsupportedProperties = @(
                                "aboutMe",
                                "birthday",
                                "hireDate",
                                "interests",
                                "mySite",
                                "preferredName",
                                "pastProjects",
                                "responsibilities",
                                "schools",
                                "skills")

                            # remove unsupported properties
                            foreach ($prop in $unsupportedProperties)
                            {
                                $selectProperties.Remove($prop)
                            }

                            # sanity checks
                            # check for GUID
                            if ([System.Guid]::TryParse($account,$([ref][System.Guid]::Empty)))
                            {
                                Write-Verbose 'Found valid GUID...'
                                $userLink = "/$($account)" + '?$select=' + $($selectProperties -join ',')
                            }
                            if ([System.Boolean]($account -as [System.Net.Mail.MailAddress]))
                            {
                                Write-Verbose 'Seems to be an mail/UPN...'
                                $userLink = '/?$filter=mail eq ' + "'" + $account + "'" + '?$select=' + $($selectProperties -join ',')
                            }

                            $restParams = @{
                                ContentType = 'application/json'
                                Method = 'GET'
                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                Uri = 'https://graph.microsoft.com/beta/directory/deletedItems/microsoft.graph.user' + $userLink
                                TimeoutSec = $TimeoutSec
                            }

                            $global:data = Invoke-RestMethod @restParams

                            # create custom object
                            if ($data.value)
                            {
                                $userObject = $data.value | Select-Object * -ExcludeProperty "@odata.context"
                            }
                            else
                            {
                                $userObject = $data | Select-Object * -ExcludeProperty "@odata.context"
                            }

                        }
                        else
                        {
                            if ($MinAttributeSet)
                            {
                                $body = @{
                                requests = @(
                                    @{
                                        url = "/users/$($account)"
                                        method = 'GET'
                                        id = '1'
                                    })
                                }
                            }
                            else
                            {
                                $body = @{
                                    requests = @(
                                        @{
                                            url = "/users/$($account)" + '?$select=' + $($selectProperties -join ',')
                                            method = 'GET'
                                            id = '1'
                                        },
                                        @{
                                            url = "/users/$($account)/manager"
                                            method = 'GET'
                                            id = '2'
                                        },
                                        @{
                                            url = "/users/$($account)/memberof"
                                            method = 'GET'
                                            id = '3'
                                        },
                                        @{
                                            url = "/users/$($account)/licenseDetails"
                                            method = 'GET'
                                            id = '4'
                                        },
                                        @{
                                            url = "/users/$($account)/registeredDevices"
                                            method = 'GET'
                                            id = '9'
                                        },
                                        @{
                                            url = "/users/$($account)/ownedDevices"
                                            method = 'GET'
                                            id = '10'
                                        },
                                        @{
                                            url = "/users/$($account)/ownedObjects"
                                            method = 'GET'
                                            id = '11'
                                        },
                                        @{
                                            url = "/users/$($account)/createdObjects"
                                            method = 'GET'
                                            id = '12'
                                        },
                                        @{
                                            url = "/users/$($account)/drive"
                                            method = 'GET'
                                            id = '13'
                                        },
                                        @{
                                            url = "/users/$($account)/oauth2PermissionGrants"
                                            method = 'GET'
                                            id = '14'
                                        }
                                    )
                                }
                            }

                            if ($GetMailboxSettings)
                            {
                                $mailboxsettings = @{
                                        url = "/users/$($account)" + '/mailboxSettings'
                                        method = 'GET'
                                        id = '5'
                                    }

                                $body.requests += $mailboxsettings
                            }

                            if ($GetAuthMethods)
                            {
                                $methods = @{
                                        url = "/users/$($account)/authentication/methods"
                                        method = 'GET'
                                        id = '6'
                                    }

                                $body.requests += $methods

                                $passwordMethods = @{
                                        url = "/users/$($account)/authentication/passwordMethods"
                                        method = 'GET'
                                        id = '7'
                                    }

                                $body.requests += $passwordMethods

                                $phoneMethods = @{
                                        url = "/users/$($account)/authentication/phoneMethods"
                                        method = 'GET'
                                        id = '8'
                                    }

                                $body.requests += $phoneMethods

                                $fido2Methods = @{
                                        url = "/users/$($account)/authentication/fido2Methods"
                                        method = 'GET'
                                        id = '15'
                                    }

                                $body.requests += $fido2Methods

                                $passwordlessMicrosoftAuthenticatorMethods = @{
                                        url = "/users/$($account)/authentication/passwordlessMicrosoftAuthenticatorMethods"
                                        method = 'GET'
                                        id = '16'
                                    }

                                $body.requests += $passwordlessMicrosoftAuthenticatorMethods

                            }

                            $restParams = @{
                                ContentType = 'application/json'
                                Method = 'POST'
                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                Body = $body | ConvertTo-Json -Depth 4
                                Uri = 'https://graph.microsoft.com/beta/$batch'
                                TimeoutSec = $TimeoutSec
                            }

                            $global:data = Invoke-RestMethod @restParams

                            # create custom object
                            $userObject = New-Object -TypeName psobject
                            $userInfo = $null
                            $userInfo = ($data.responses | Where-Object -FilterScript { $_.id -eq 1}).Body | Select-Object * -ExcludeProperty "@odata.context"
                            if ('ResourceNotFound' -eq $userInfo.error.code)
                            {
                                Write-Verbose "ResourceNotFound thrown for:$($account)..."
                                $userObject = New-Object -TypeName psobject
                                $userObject | Add-Member -MemberType NoteProperty -Name User -Value $account
                                $userObject | Add-Member -MemberType NoteProperty -Name Error -Value 'ResourceNotFound'
                                $collection += $userObject
                                break
                            }

                            $userProperties = $userInfo | Get-Member -MemberType NoteProperty

                            foreach ($property in $userProperties)
                            {
                                $userObject | Add-Member -MemberType NoteProperty -Name $( $property.Name ) -Value $( $userInfo.$( $property.Name ) )
                            }

                            if (-not $MinAttributeSet)
                            {

                                # add manager to object
                                $managerResponse = $data.responses | Where-Object -FilterScript {$_.id -eq 2}
    
                                if ('200' -eq $managerResponse.status)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name Manager -Value @( $($managerResponse.body | Select-Object * -ExcludeProperty "@odata.Context","@odata.type") )
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name Manager -Value @( $($managerResponse.body.error) )
                                }
    
                                # extract memberOf response
                                $responseMemberOf = $data.responses | Where-Object -FilterScript {$_.id -eq 3}
    
                                if ('200' -eq $responseMemberOf.status)
                                {
                                    if ($responseMemberOf.body.'@odata.nextLink')
                                    {
    
                                        Write-Verbose 'Need to fetch more data for memberOf...'
                                        [System.Int16]$counter = '2'
                                        # create collection
                                        $groupCollection = [System.Collections.ArrayList]@()
    
                                        # add first batch of groups to collection
                                        $groupCollection += $responseMemberOf.body.value | Select-Object * -ExcludeProperty "@odata.type"
    
                                        do
                                        {
                                            $groupParams = @{
                                                ContentType = 'application/json'
                                                Method = 'GET'
                                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                                Uri = $responseMemberOf.body.'@odata.nextLink'
                                                TimeoutSec = $TimeoutSec
                                            }
    
                                            $responseMemberOf = Invoke-RestMethod @groupParams
    
                                            if ($responseMemberOf.body.value)
                                            {
                                                $groupCollection += $responseMemberOf.body.value | Select-Object * -ExcludeProperty "@odata.type"
                                            }
                                            else
                                            {
                                                $groupCollection += $responseMemberOf.value | Select-Object * -ExcludeProperty "@odata.type"
                                            }
    
                                            Write-Verbose "Pagecount:$($counter)..."
                                            $counter++
    
                                        } while ($responseMemberOf.body.'@odata.nextLink')
    
                                        $userObject | Add-Member -MemberType NoteProperty -Name MemberOf -Value @( $groupCollection )
    
                                    }
                                    else
                                    {
                                        $userObject | Add-Member -MemberType NoteProperty -Name MemberOf -Value @( $($responseMemberOf.body.value | Select-Object * -ExcludeProperty "@odata.type") )
                                    }
    
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name MemberOf -Value @( $($responseMemberOf.body.error) )
                                }
    
                                if ($userInfo.id)
                                {
                                    try {
                                        # retrieve joined teams
                                        $teamsParams = @{
                                                    ContentType = 'application/json'
                                                    Method = 'GET'
                                                    Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                                    Uri = "https://graph.microsoft.com/beta/users/$($userInfo.id)/joinedTeams"
                                                    TimeoutSec = $TimeoutSec
                                                    ErrorAction = 'SilentlyContinue'
                                        }
    
                                        $responseJoinedTeams = Invoke-RestMethod @teamsParams
    
                                        if ($responseJoinedTeams.'@odata.nextLink')
                                        {
    
                                            Write-Verbose 'Need to fetch more data for joinedTeams...'
                                            # create collection
                                            $teamsCollection = [System.Collections.ArrayList]@()
    
                                            # add first batch of groups to collection
                                            $teamsCollection += $responseJoinedTeams.Value
    
                                            do
                                            {
                                                $groupParams = @{
                                                    ContentType = 'application/json'
                                                    Method = 'GET'
                                                    Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                                    Uri = $($responseJoinedTeams.'@odata.nextLink')
                                                    TimeoutSec = $TimeoutSec
                                                }
    
                                                $responseJoinedTeams = Invoke-RestMethod @groupParams
    
                                                $teamsCollection += $responseJoinedTeams.Value
    
                                            } while ($responseJoinedTeams.'@odata.nextLink')
    
                                            $userObject | Add-Member -MemberType NoteProperty -Name JoinedTeams -Value @( $teamsCollection )
    
                                        }
                                        else
                                        {
                                            $userObject | Add-Member -MemberType NoteProperty -Name JoinedTeams -Value @($responseJoinedTeams.Value)
                                        }
                                    }
                                    catch {
                                        $userObject | Add-Member -MemberType NoteProperty -Name JoinedTeams -Value $("Error for resource:$($_.Exception.Response.ResponseUri.ToString()) with StatusCode:$($_.Exception.Response.StatusCode.ToString())")
                                    }
    
                                    if ($GetDeltaToken)
                                    {
                                        try {
                                            Write-Verbose "Get delta for $($userInfo.userPrincipalName)"
                                            $deltaParams = @{
                                                    ContentType = 'application/json'
                                                    Method = 'GET'
                                                    Headers = @{ Authorization = "Bearer $($AccessToken)"; prefer = "return=minimal"}
                                                    Uri = 'https://graph.microsoft.com/beta/users/delta?' + '$filter=id eq ' + "'$($userInfo.id)'" + '&$deltaToken=latest'
                                                    TimeoutSec = $TimeoutSec
                                            }
    
                                            $responseDelta = Invoke-RestMethod @deltaParams
    
                                            if ( -not [System.String]::IsNullOrEmpty($responseDelta.'@odata.deltaLink') )
                                            {
                                                # create custom object
                                                $deltaObject = New-Object -TypeName psobject
                                                # add properties to custom object
                                                $deltaObject | Add-Member -MemberType NoteProperty -Name createdDateTimeUTC -Value $(Get-Date (Get-Date).ToUniversalTime() -Format u)
                                                $deltaObject | Add-Member -MemberType NoteProperty -Name deltaLink -Value $($responseDelta.'@odata.deltaLink')
                                                # add custom object to user object
                                                $userObject | Add-Member -MemberType NoteProperty -Name DeltaLink -Value @( $deltaObject )
                                            }
                                        }
                                        catch {
                                            $userObject | Add-Member -MemberType NoteProperty -Name DeltaLink -Value $("Error for resource:$($_.Exception.Response.ResponseUri.ToString()) with StatusCode:$($_.Exception.Response.StatusCode.ToString())")
                                        }
                                    }
                                }
    
                                if ($GetMailboxSettings)
                                {
                                    $responseMailboxsettings = $data.responses | Where-Object -FilterScript {$_.id -eq 5}
    
                                    if ('200' -eq $responseMailboxsettings.status)
                                    {
                                        $userObject | Add-Member -MemberType NoteProperty -Name MailboxSettings -Value  @( $($responseMailboxsettings.body | Select-Object * -ExcludeProperty '@odata.context') )
                                    }
                                    else
                                    {
                                        $userObject | Add-Member -MemberType NoteProperty -Name MailboxSettings -Value  @( $($responseMailboxsettings.body.error) )
                                    }
                                }
    
                                if ($GetAuthMethods)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name AuthenticationMethods -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 6) -and ($_.status -eq 200)}).body.value ))
                                    $userObject | Add-Member -MemberType NoteProperty -Name PasswordMethods -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 7) -and ($_.status -eq 200)}).body.value ))
                                    $userObject | Add-Member -MemberType NoteProperty -Name PhoneMethods -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 8) -and ($_.status -eq 200)}).body.value ))
                                    $userObject | Add-Member -MemberType NoteProperty -Name Fido2Methods -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 15) -and ($_.status -eq 200)}).body.value ))
                                    $userObject | Add-Member -MemberType NoteProperty -Name PasswordlessMicrosoftAuthenticatorMethods -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 16) -and ($_.status -eq 200)}).body.value ))
                                }
    
                                $licenseResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 4}
    
                                if ('200' -eq $licenseResponse.status)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name LicenseDetails -Value  @( $($licenseResponse.body.value) )
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name LicenseDetails -Value  @( $($licenseResponse.body.error) )
                                }
    
                                $registeredDevicesResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 9}
    
                                if ('200' -eq $registeredDevicesResponse.status)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name RegisteredDevices -Value  @( $($registeredDevicesResponse.body.value | Select-Object * -ExcludeProperty "@odata.type") )
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name RegisteredDevices -Value  @( $($registeredDevicesResponse.body.error) )
                                }
    
                                $ownedDevicesResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 10}
    
                                if ('200' -eq $ownedDevicesResponse.status)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name OwnedDevices -Value  @( $($ownedDevicesResponse.body.value | Select-Object * -ExcludeProperty "@odata.type") )
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name OwnedDevices -Value  @( $($ownedDevicesResponse.body.error) )
                                }
    
                                $ownedObjectsResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 11}
    
                                if ('200' -eq $ownedObjectsResponse.status)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name OwnedObjects -Value  @( $($ownedObjectsResponse.body.value) )
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name OwnedObjects -Value  @( $($ownedObjectsResponse.body.error) )
                                }
    
                                $createdObjectsResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 12}
    
                                if ('200' -eq $createdObjectsResponse.status)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name CreatedObjects -Value  @( $($createdObjectsResponse.body.value) )
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name CreatedObjects -Value  @( $($createdObjectsResponse.body.error) )
                                }
    
                                if ($createdObjectsResponse.body.value.'@odata.type' -contains '#microsoft.graph.servicePrincipal')
                                {
                                    Write-Verbose "ServicePrincipals found as ownedObjects. Gathering details..."
                                    # get id for servicePrincipal
                                    $servicePrincipalIDs = ($createdObjectsResponse.body.value | Where-Object '@odata.type' -eq '#microsoft.graph.servicePrincipal').id
    
                                    # create collection
                                    [System.Collections.ArrayList]$servicePrincipalCollection = @()
    
                                    foreach ($id in $servicePrincipalIDs)
                                    {
                                        Write-Verbose "Requesting details for SPN $($id)..."
                                        $spnParams = @{
                                                ContentType = 'application/json'
                                                Method = 'GET'
                                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                                Uri = "https://graph.microsoft.com/beta/servicePrincipals/$($id)"
                                                TimeoutSec = $TimeoutSec
                                        }
    
                                        $responseSPN = Invoke-RestMethod @spnParams
    
                                        $servicePrincipalCollection += $responseSPN
                                    }
    
                                    $userObject | Add-Member -MemberType NoteProperty -Name ServicePrincipalDetails -Value @( $servicePrincipalCollection )
                                }
    
                                if ($ownedObjectsResponse.body.value.'@odata.type' -contains '#microsoft.graph.application')
                                {
                                    Write-Verbose "ServicePrincipals found as ownedObjects. Gathering details..."
                                    # get id for application
                                    $appIDs = ($ownedObjectsResponse.body.value | Where-Object '@odata.type' -eq '#microsoft.graph.application').id
    
                                    # create collection
                                    [System.Collections.ArrayList]$applicationCollection = @()
    
                                    foreach ($id in $appIDs)
                                    {
                                        Write-Verbose "Requesting details for Application $($id)..."
                                        $appParams = @{
                                                ContentType = 'application/json'
                                                Method = 'GET'
                                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                                Uri = "https://graph.microsoft.com/beta/applications/$($id)"
                                                TimeoutSec = $TimeoutSec
                                        }
    
                                        $responseApp = Invoke-RestMethod @appParams
    
                                        $applicationCollection += $responseApp | Select-Object * -ExcludeProperty "@odata.context"
                                    }
    
                                    $userObject | Add-Member -MemberType NoteProperty -Name ApplicationDetails -Value @( $applicationCollection )
                                
                                }
    
                                $driveResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 13}
                                if ('200' -eq $driveResponse.status)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name OneDrive -Value  @( $($driveResponse.body | Select-Object * -ExcludeProperty "@odata.context") )
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name OneDrive -Value  @( $($driveResponse.body.error) )
                                }
    
                                $oauth2PermissionGrantsResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 14}
                                if ('200' -eq $oauth2PermissionGrantsResponse.status)
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name OAuth2PermissionGrants -Value  @( $($oauth2PermissionGrantsResponse.body.value | Select-Object * -ExcludeProperty "@odata.context") )
                                }
                                else
                                {
                                    $userObject | Add-Member -MemberType NoteProperty -Name OAuth2PermissionGrants -Value  @( $($oauth2PermissionGrantsResponse.body.error) )
                                }

                            }

                        }

                        $collection += $userObject

                        if ($ShowProgress)
                        {
                            $progressParams.Status = "Ready"
                            $progressParams.Add('Completed',$true)
                            Write-Progress @progressParams
                        }

                        $processingTime.Stop()
                        $keepTrying = $false
                        Write-Verbose "User processing time:$($processingTime.Elapsed.ToString())"

                    }
                    catch {

                    $statusCode = $_.Exception.Response.StatusCode.value__

                    if ('401' -eq $statusCode)
                    {

                        Write-Verbose "HTTP error $($statusCode) thrown. Will try to acquire new access token..."
                        $acquireToken = @{
                            Authority = $Authority
                            ClientId = $ClientId
                            ClientSecret = $ClientSecret
                            Certificate = $Certificate
                        }

                        $AccessToken = (Get-AccessTokenNoLibraries @acquireToken).access_token

                        if ([System.String]::IsNullOrWhiteSpace($AccessToken))
                        {
                            Write-Verbose "Couldn't acquire a token. Will stop now..."

                            $userObject = New-Object -TypeName psobject
                            $userObject | Add-Member -MemberType NoteProperty -Name User -Value $account
                            $userObject | Add-Member -MemberType NoteProperty -Name Error -Value 'Could not acquire access token'
                            $collection += $userObject
                            break
                        }

                        Write-Verbose 'Setting accesstoken...'
                        $psBoundParameters['AccessToken'] = $AccessToken

                        # increase retrycount
                        $retryCount++
                        Write-Verbose "Retrycount:$($retryCount)"

                    }
                    elseif ('403' -eq $statusCode)
                    {
                        Write-Verbose "Error for resource $($_.Exception.Response.ResponseUri.ToString()) with StatusCode $($_.Exception.Response.StatusCode.ToString())"
                        # increase retrycount
                        $retryCount++
                    }
                    elseif ('429' -eq $statusCode)
                    {

                        Write-Verbose "HTTP error $($statusCode) thrown. Will backoff..."
                        Start-Sleep -Seconds 2

                    }
                    elseif ('503' -eq $statusCode)
                    {

                        Write-Verbose "HTTP error $($statusCode) thrown. Will retry..."
                        # increase retrycount
                        $retryCount++
                        Write-Verbose "Retrycount:$($retryCount)"

                    }
                    else
                    {
                        $userObject = New-Object -TypeName psobject
                        $userObject | Add-Member -MemberType NoteProperty -Name User -Value $account
                        $userObject | Add-Member -MemberType NoteProperty -Name Error -Value $_.Exception
                        $collection += $userObject
                        Write-Verbose "Error occured for account $($account)..."
                        break
                    }

                    if ($retryCount -eq $MaxRetry)
                    {
                        Write-Verbose 'MaxRetries done. Skip this user now...'
                        $keepTrying = $false

                        $userObject = New-Object -TypeName psobject
                        $userObject | Add-Member -MemberType NoteProperty -Name User -Value $account
                        $userObject | Add-Member -MemberType NoteProperty -Name Error -Value $_.Exception
                        $collection += $userObject
                    }
                }

                } while ($keepTrying)

            }

        }
    }

    end
    {
        #monitor and retrieve the created jobs
        if ($MultiThread)
        {
            $SleepTimer = 1000
            While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0) {

                Write-Progress `
                    -id 1 `
                    -Activity "Waiting for Jobs - $($Threads - $($RunspacePool.GetAvailableRunspaces())) of $Threads threads running" `
                    -PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
                    -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining out of $($Jobs.Count)"

                ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $true})) {
                    $Job.Thread.EndInvoke($Job.Handle)
                    $Job.Thread.Dispose()
                    $Job.Thread = $Null
                    $Job.Handle = $Null
                }

                # kill all incomplete threads when hit "CTRL+q"
                if ($Host.UI.RawUI.KeyAvailable) {
                    $KeyInput = $Host.UI.RawUI.ReadKey("IncludeKeyUp,NoEcho")
                    if (($KeyInput.ControlKeyState -cmatch '(Right|Left)CtrlPressed') -and ($KeyInput.VirtualKeyCode -eq '81')) {
                        Write-Host -fore red "Kill all incomplete threads..."
                            ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})) {
                                Write-Host -fore yellow "Stopping job $($Job.Object)..."
                                $Job.Thread.Stop()
                                $Job.Thread.Dispose()
                             }
                        Write-Host -fore red "All jobs terminated..."
                        break
                    }
                }

                Start-Sleep -Milliseconds $SleepTimer

            }

            # clean-up
            Write-Verbose 'Perform cleanup of runspaces...'
            $RunspacePool.Close() | Out-Null
            $RunspacePool.Dispose() | Out-Null
            [System.GC]::Collect()
        }
        else
        {
            $collection
        }

        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Get-RESTAzKeyVaultSecret
{
    <#
        .SYNOPSIS
            This function uses Azure Key Vault REST API for accessing secrets.
        .DESCRIPTION
            This function will retrieve secrets from Azure Key Vault using REST API. You can either use AuthCode or ClientCredential flow for acquiring an access token and accessing Azure Key Vault.
        .PARAMETER AZKeyVaultBaseUri
            The parameter AZKeyVaultBaseUri is required and is the base uri of the Azure Key Vault.
        .PARAMETER ClientID
            The parameter ClientID is required and defines the registered application.
        .PARAMETER ClientSecret
            The parameter ClientSecret is optional and used for ClientCredential flow.
        .PARAMETER TenantID
            The parameter TenantID is required, when ClientSecret is used and for ClientCredential flow.
        .PARAMETER RedirectUri
            The parameter RedirectUri is required for AuthCode flow.
        .PARAMETER SecretName
            The parameter SecretName defines the name of the secret you want to access in Azure Key Vault.
        .PARAMETER CertificateName
            The parameter CertificateName defines the name of the certificate stored in Azure Key Vault.
        .PARAMETER ListSecrets
            The parameter ListSecrets will list all secrest stored in Azure Key Vault.
        .EXAMPLE
            Get-RESTAzKeyVaultSecret -AZKeyVaultBaseUri https://exov2.vault.azure.net/ -ClientID f7f6eg58-14e0-4v12-8675-eb7980a05c7e -RedirectUri https://login.microsoftonline.com/common/oauth2/nativeclient -ListSecrets
            Get-RESTAzKeyVaultSecret -AZKeyVaultBaseUri https://exov2.vault.azure.net/ -ClientID f7f6eg58-14e0-4v12-8675-eb7980a05c7e -RedirectUri https://login.microsoftonline.com/common/oauth2/nativeclient -SecretName MySecretString
        .NOTES
            
        .LINK
            https://docs.microsoft.com/rest/api/keyvault/getsecrets/getsecrets
            https://docs.microsoft.com/rest/api/keyvault/getsecret/getsecret
            https://docs.microsoft.com/rest/api/keyvault/getcertificate/getcertificate
    #>

    [CmdletBinding(DefaultParameterSetName='AuthCodeFlow')]
    Param (
        [parameter( Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [System.Uri]
        $AZKeyVaultBaseUri,

        [parameter( Mandatory=$true, Position=1)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ClientID,

        [parameter( Mandatory=$false, Position=2, ParameterSetName='ClientSecretFlow')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ClientSecret,

        [parameter( Mandatory=$true, Position=3, ParameterSetName='ClientSecretFlow')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $TenantID,

        [parameter( Mandatory=$true, Position=4, ParameterSetName='AuthCodeFlow')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $RedirectUri,

        [parameter( Mandatory=$false, Position=5, ParameterSetName="Secret")]
        [Parameter( ParameterSetName="AuthCodeFlow")]
        [Parameter( ParameterSetName="ClientSecretFlow")]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $SecretName,

        [parameter( Mandatory=$false, Position=6, ParameterSetName="Certificate")]
        [Parameter( ParameterSetName="AuthCodeFlow")]
        [Parameter( ParameterSetName="ClientSecretFlow")]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $CertificateName,

        [parameter( Mandatory=$false, Position=7, ParameterSetName="ListSecret")]
        [Parameter( ParameterSetName="AuthCodeFlow")]
        [Parameter( ParameterSetName="ClientSecretFlow")]
        [System.Management.Automation.SwitchParameter]
        $ListSecrets

    )

    begin
    {

        Write-Verbose "ParameterSet:$($PSCmdlet.ParameterSetName)"

        if ( ([System.String]::IsNullOrWhiteSpace($SecretName)) -and ([System.String]::IsNullOrWhiteSpace($CertificateName)) )
        {
            Write-Verbose 'Secret- and CertificateName are empty. Will set ListSecrets to $true...'
            $ListSecrets = $true
        }

        function Get-AADAuth
        {
            [CmdletBinding()]
            Param
            (
                [System.Uri]
                $Authority,

                [System.String]
                $Tenant,

                [System.String]
                $Client_ID,

                [ValidateSet("code","token")]
                [System.String]
                $Response_Type = 'code',

                [System.Uri]
                $Redirect_Uri,

                [ValidateSet("query","fragment")]
                [System.String]
                $Response_Mode,

                [System.String]
                $State,

                [System.String]
                $Resource,

                [System.String]
                $Scope,

                [ValidateSet("login","select_account","consent","admin_consent","none")]
                [System.String]
                $Prompt,

                [System.String]
                $Login_Hint,

                [System.String]
                $Domain_Hint,

                [ValidateSet("plain","S256")]
                [System.String]
                $Code_Challenge_Method,

                [System.String]
                $Code_Challenge,

                [System.Management.Automation.SwitchParameter]
                $V2
            )

            begin
            {
                Add-Type -AssemblyName System.Web

                if ($V2)
                {
                    $OAuthSub = '/oauth2/v2.0/authorize?'
                }
                else
                {
                    $OAuthSub = '/oauth2/authorize?'
                }

                #create autorithy Url
                $AuthUrl = $Authority.AbsoluteUri + $Tenant + $OAuthSub
                Write-Verbose -Message "AuthUrl:$($AuthUrl)"

                #create empty body variable
                $Body = @{}
                $Url_String = ''

                function Show-OAuthWindow
                {
                    [CmdletBinding()]
                    param(
                        [System.Uri]
                        $Url,

                        [ValidateSet("query","fragment")]
                        [System.String]
                        $Response_Mode
                    )

                    Write-Verbose "Show-OAuthWindow Url:$($Url)"
                    Add-Type -AssemblyName System.Windows.Forms

                    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
                    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url ) }
                    $DocComp  = {
                        $uri = $web.Url.AbsoluteUri
                        if ($Uri -match "error=[^&]*|code=[^&]*|code=[^#]*|#access_token=*")
                        {
                            $form.Close()
                        }
                    }

                    if (-not $Redirect_Uri.AbsoluteUri -eq 'urn:ietf:wg:oauth:2.0:oob' )
                    {
                        $web.ScriptErrorsSuppressed = $true
                    }
                    $web.Add_DocumentCompleted($DocComp)
                    $form.Controls.Add($web)
                    $form.Add_Shown({$form.Activate()})
                    $form.ShowDialog() | Out-Null

                    switch ($Response_Mode)
                    {
                        "query"     {$UrlToBeParsed = $web.Url.Query}
                        "fragment"  {$UrlToBeParsed = $web.Url.Fragment}
                        "form_post" {$UrlToBeParsed = $web.Url.Fragment}
                    }

                    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($UrlToBeParsed)
                    $result = $web
                    $output = @{}
                    foreach($key in $queryOutput.Keys){
                        $output["$key"] = $queryOutput[$key]
                    }

                    $output
                }
            }

            process
            {
                $Params = $PSBoundParameters.GetEnumerator() | Where-Object -FilterScript {$_.key -inotmatch 'Verbose|v2|authority|tenant|Redirect_Uri'}
                foreach ($Param in $Params)
                {
                    Write-Verbose -Message "$($Param.Key)=$($Param.Value)"
                    $Url_String += "&" + $Param.Key + '=' + [System.Web.HttpUtility]::UrlEncode($Param.Value)
                }

                if ($Redirect_Uri)
                {
                    $Url_String += "&Redirect_Uri=$Redirect_Uri"
                }
                $Url_String = $Url_String.TrimStart("&")
                Write-Verbose "RedirectURI:$($Redirect_Uri)"
                Write-Verbose "URL:$($Url_String)"
                $Response = Show-OAuthWindow -Url $($AuthUrl + $Url_String) -Response_Mode $Response_Mode
            }

            end
            {
                if ($Response.Count -gt 0)
                {
                    $Response
                }
                else
                {
                    Write-Verbose "Error occured"
                    Add-Type -AssemblyName System.Web
                    [System.Web.HttpUtility]::UrlDecode($result.Url.OriginalString)
                }
            }
        }

        if ($ClientSecret)
        {
            Write-Verbose 'Request token using ClientSecret...'
            $bodyGetToken = @{
                client_id = $ClientID
                client_secret = $ClientSecret
                grant_type = 'client_credentials'
                scope = 'https://vault.azure.net/.default'
            }

            $paramsGetToken = @{
                ContentType = 'application/x-www-form-urlencoded'
                Uri = 'https://login.microsoftonline.com/' + $TenantID + '/oauth2/v2.0/token'
                Body = $bodyGetToken
                Method = 'POST'
            }

            $global:token = Invoke-RestMethod @paramsGetToken

        }
        else
        {
            Write-Verbose 'Request token using AuthCode flow...'

            $authParams = @{
                Authority = 'https://login.microsoftonline.com/'
                Tenant = 'common'
                Client_ID = $ClientID
                Redirect_Uri = $RedirectUri
                Resource = 'https://vault.azure.net'
                Prompt = 'select_account'
                Response_Mode = 'query'
                Response_Type = 'code'
            }

            $global:authCode = Get-AADAuth @authParams

            $body = @{
                client_id = $authParams.Client_ID
                code = $($authCode['code'])
                redirect_uri = $authParams.Redirect_URI
                grant_type = "authorization_code"
            }

            $params = @{
                ContentType = 'application/x-www-form-urlencoded'
                Method = 'POST'
                Uri = "https://login.microsoftonline.com/common/oauth2/token"
                Body = $body
            }

            $global:token = Invoke-RestMethod @params
        }
        $collection = [System.Collections.ArrayList]@()

        $secretObject = New-Object -TypeName psobject

    }

    process
    {

        
        if ($ListSecrets)
        {
            $paramsGetSecret = @{
                Method = 'GET'
                URI = $AZKeyVaultBaseUri.AbsoluteUri + 'secrets?api-version=7.0'
                Headers = @{ Authorization = "Bearer $($token.access_token)"; Accept = "*/*"; "Accept-Encoding" = 'gzip, deflate, br'}
                ContentType = 'application/json'
            }

            $secrets = Invoke-RestMethod @paramsGetSecret

            $collection += $secrets.value
        }

        if ($SecretName)
        {
            $paramsSecretName = @{
                Method = 'GET'
                URI = $AZKeyVaultBaseUri.AbsoluteUri + "secrets/$($SecretName)?api-version=7.0"
                Headers = @{ Authorization = "Bearer $($token.access_token)"; Accept = "*/*"; "Accept-Encoding" = 'gzip, deflate, br'}
                ContentType = 'application/json'
            }

            $secret = Invoke-RestMethod @paramsSecretName

            $collection += $secret
        }

        if ($CertificateName)
        {
            $paramsCertificateName = @{
                Method = 'GET'
                URI = $AZKeyVaultBaseUri.AbsoluteUri + "certificates/$($CertificateName)?api-version=7.0"
                Headers = @{ Authorization = "Bearer $($token.access_token)"; Accept = "*/*"; "Accept-Encoding" = 'gzip, deflate, br'}
                ContentType = 'application/json'
            }

            $cert = Invoke-RestMethod @paramsCertificateName

            $collection += $cert
        }

    }

    end{

        $collection

    }

}

function global:ConvertFrom-AzKeVaultString
{
    <#
        .SYNOPSIS
            This function converts a Base64 secured certificate string to a X509Certificate2 object.
        .DESCRIPTION
            When you retrieve a certificate secret from Azure Key Vault, it is a Base64 string and needs to be converted in case of the need of a X509Certificate2 PowerShell object. With this function you can provide the string value and the output will be the object.
        .PARAMETER value
            
        .EXAMPLE
            $secretCert = Get-RESTAzKeyVaultSecret -AZKeyVaultBaseUri https://exov2.vault.azure.net/ -ClientID f7f6eg58-14e0-4v12-8675-eb7980a05c7e -RedirectUri https://login.microsoftonline.com/common/oauth2/nativeclient -SecretName EXOv2Cert
            ConvertFrom-AzKeVaultString -value $secretCert.value
        .LINK
            https://docs.microsoft.com/dotnet/api/system.security.cryptography.x509certificates.x509certificate2.-ctor?view=netcore-3.1
            https://stackoverflow.com/questions/30237307/converting-base64-string-to-x509-certifcate
    #>
    [CmdletBinding()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    Param
    (
        [Parameter(
            Mandatory=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [System.String]
        $value
    )

    try
    {
        [System.Security.Cryptography.X509Certificates.X509Certificate2][System.Convert]::FromBase64String($value)
    }
    catch
    {
        $_
    }
}

function global:Get-MSGraphTeam
{
    <#
        .SYNOPSIS
            This function retrieves a Microsoft Teams properties.
        .DESCRIPTION
            This function retrieves a Microsoft Teams properties including channels,tabs and installedApps.
        .PARAMETER ID
            The parameter ID specifies the id of the Microsoft Team to be queried.
        .PARAMETER AccessToken
            This required parameter AccessToken takes the Bearer access token for authentication the requests. The parameter takes a previously acquired access token.
        .EXAMPLE
            Get-MSGraphTeam -ID 6288514a-8950-4426-be05-d2955a03ea27
        .NOTES
            If you want to leverage all functionality you will need to provide an access token with the following claims:
                Directory.Read.All
                Group.Read.All
        .LINK
            https://docs.microsoft.com/graph/api/resources/user?view=graph-rest-beta
            https://docs.microsoft.com/graph/paging
            https://docs.microsoft.com/graph/json-batching
            https://docs.microsoft.com/graph/query-parameters
            https://docs.microsoft.com/graph/permissions-reference
    #>
    [CmdletBinding()]
    param(
        [parameter( Position=0)]
        [System.String[]]
        $ID,

        [parameter( Position=1)]
        [System.String]
        $AccessToken = $MSGraphToken[0].AccessToken
    )

    begin
    {

        $timer = [System.Diagnostics.Stopwatch]::StartNew()

        $collection = [System.Collections.ArrayList]@()

    }

    process
    {

        foreach($team in $ID)
        {

            $body = @{
                requests = @(
                    @{
                        url = "/teams/$team"
                        method = 'GET'
                        id = '1'
                    },
                    @{
                        url = "/teams/$team/channels"
                        method = 'GET'
                        id = '2'
                    },
                    @{
                        url = "/teams/$team/channels/$team/tabs"
                        method = 'GET'
                        id = '3'
                    },
                    @{
                        url = "/teams/$team/installedApps" + '?$expand=teamsAppDefinition'
                        method = 'GET'
                        id = '4'
                    }
                )
            }

            $restParams = @{
                ContentType = 'application/json'
                Method = 'POST'
                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                Body = $body | ConvertTo-Json -Depth 4
                Uri = 'https://graph.microsoft.com/beta/$batch'
            }

            $global:data = Invoke-RestMethod @restParams

            # create custom object
            $teamObject = New-Object -TypeName psobject
            $teamInfo = $null
            $teamInfo = ($data.responses | Where-Object -FilterScript { $_.id -eq 1}).Body | Select-Object * -ExcludeProperty "@odata.context"
            $teamProperties = $teamInfo | Get-Member -MemberType NoteProperty

            foreach ($property in $teamProperties)
            {
                $teamObject | Add-Member -MemberType NoteProperty -Name $( $property.Name ) -Value $( $teamInfo.$( $property.Name ) )
            }

            # add channels
            $teamObject | Add-Member -MemberType NoteProperty -Name Channels -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 2) -and ($_.status -eq 200)}).Body.value ))
            # add tabs
            $teamObject | Add-Member -MemberType NoteProperty -Name Tabs -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 3) -and ($_.status -eq 200)}).Body | Select-Object * -ExcludeProperty "@odata.Context","@odata.type" ))
            # add apps
            $teamObject | Add-Member -MemberType NoteProperty -Name InstalledApps -Value @( ($data.responses | Where-Object -FilterScript { ($_.id -eq 4) -and ($_.status -eq 200)}).Body.value.teamsAppDefinition )

            $collection += $teamObject
        }
    }

    end
    {
        $collection
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Get-AccessTokenNoLibraries
{
    <#
        .SYNOPSIS
            This function acquires an access token using either shared secret or certificate.
        .DESCRIPTION
            This function will acquire an access token from Azure AD using either shared secret or certificate and OAuth2.0 client credentials flow without the need of having any DLLs installed.
        .PARAMETER Authority
            The parameter AZKeyVaultBaseUri is required and is the base uri of the Azure Key Vault.
        .PARAMETER ClientID
            The parameter ClientID is required and defines the registered application.
        .PARAMETER ClientSecret
            The parameter ClientSecret is optional and used for ClientCredential flow.
        .PARAMETER Certificate
            The parameter Certificate is required, when a certificate is used and for ClientCredential flow. You need to provide a X509Certificate2 object.
        .EXAMPLE
            Get-AccessTokenNoLibraries -Authority = 'https://login.microsoftonline.com/53g7676c-f466-423c-82f6-vb2f55791wl7/oauth2/v2.0/token' -ClientID '1521a4b6-v487-4078-be14-c9fvb17f8ab0' -Certificate $(dir Cert:\CurrentUser\My\ | ? thumbprint -eq 'F5B5C6135719644F5582BC184B4C2239D5112C73')
            Get-AccessTokenNoLibraries -Authority = 'https://login.microsoftonline.com/53g7676c-f466-423c-82f6-vb2f55791wl7/oauth2/v2.0/token' -ClientID '1521a4b6-v487-4078-be14-c9fvb17f8ab0' -ClientSecret 'gavuCG5Pm__cG4F1~SN_03KQDAN2N~7o.L'
        .NOTES
            
        .LINK
            https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-certificate-credentials
            https://adamtheautomator.com/microsoft-graph-api-powershell/#Acquire_an_Access_Token_(Using_a_Certificate)
            https://adamtheautomator.com/microsoft-graph-api-powershell/#Acquire_an_Access_Token_(Application_Id_and_Secret)
    #>

    [CmdletBinding()]
    Param
    (
        [System.String]
        $Authority,

        [System.String]
        [ValidateNotNullOrEmpty()]
        $ClientId,

        [System.String]
        $ClientSecret,

        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $Certificate
    )

    begin
    {
        $Scope = "https://graph.microsoft.com/.default"

        if ($Certificate)
        {

            # Create base64 hash of certificate
            $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())

            # Create JWT timestamp for expiration
            $StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
            $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
            $JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0)

            # Create JWT validity start timestamp
            $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
            $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan,0)

            # Create JWT header
            $JWTHeader = @{
                alg = 'RS256'
                typ = 'JWT'
                # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
                x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='
            }

            # Create JWT payload
            $JWTPayLoad = @{
                # What endpoint is allowed to use this JWT
                aud = $Authority # "https://login.microsoftonline.com/$TenantName/oauth2/token"

                # Expiration timestamp
                exp = $JWTExpiration

                # Issuer = your application
                iss = $ClientId

                # JWT ID: random guid
                jti = [System.Guid]::NewGuid()

                # Not to be used before
                nbf = $NotBefore

                # JWT Subject
                sub = $ClientId
            }

            # Convert header and payload to base64
            $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
            $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

            $JWTPayLoadToByte =  [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
            $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

            # Join header and Payload with "." to create a valid (unsigned) JWT
            $JWT = $EncodedHeader + "." + $EncodedPayload

            # Get the private key object of your certificate
            $PrivateKey = $Certificate.PrivateKey

            # Define RSA signature and hashing algorithm
            $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
            $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

            # Create a signature of the JWT
            $Signature = [Convert]::ToBase64String(
                $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT),$HashAlgorithm,$RSAPadding)
            ) -replace '\+','-' -replace '/','_' -replace '='

            # Join the signature to the JWT with "."
            $JWT = $JWT + "." + $Signature

            # Create a hash with body parameters
            $Body = @{
                client_id = $ClientId
                client_assertion = $JWT
                client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                scope = $Scope
                grant_type = 'client_credentials'
            
            }

            # Use the self-generated JWT as Authorization
            $Header = @{
                Authorization = "Bearer $JWT"
            }

        }
        else
        {
            # Create body
            $Body = @{
                client_id = $ClientId
                client_secret = $ClientSecret
                scope = $Scope
                grant_type = 'client_credentials'
            }

        }

        # Splat the parameters for Invoke-Restmethod for cleaner code
        $PostSplat = @{
            ContentType = 'application/x-www-form-urlencoded'
            Method = 'POST'
            Body = $Body
            Uri = $Authority
        }

        if ($Certificate)
        {
            Write-Verbose 'Adding headers for certificate based request...'
            $PostSplat.Add('Headers',$Header)
        }

    }

    process
    {
        try {

            

            $accessToken = Invoke-RestMethod @PostSplat

        }
        catch{
            $_
        }
    }

    end
    {
        $accessToken
    }

}

function global:Get-UserRealm
{
    [CmdletBinding()]
    Param
    (
        [System.String]
        $UserPrincipalName,

        [System.Int32]
        $TimeoutSec = '15'
    )
    
    $params = @{
        Uri = "https://login.microsoftonline.com/GetUserRealm.srf?login=$UserPrincipalName"
        TimeoutSec = $TimeoutSec
    }

    Invoke-RestMethod @params
}

function global:Get-ProtocolLogs
{
    [CmdletBinding()]
    param(
        [parameter( Mandatory=$false, Position=0)]
        [System.Management.Automation.SwitchParameter]
        $ByRemoteIP,

        [parameter( Mandatory=$false, Position=1)]
        [System.String[]]
        $DetailsForRemoteIP,

        [parameter( Mandatory=$false, Position=2)]
        [System.String]
        $DataContains,

        [parameter( Mandatory=$false, Position=3)]
        [System.String]
        $ContextContains,

        [parameter( Mandatory=$false, Position=4)]
        [System.String[]]
        $SessionID,

        [parameter( Mandatory=$false, Position=5)]
        [ValidateScript({if (Test-Path $_ -PathType container) {$true} else {Throw "$_ is not a valid path!"}})]
        [System.String]
        $Outpath = $env:temp,

        [parameter( Mandatory=$false, Position=6)]
        [System.Int32]
        $StartDate = "$((Get-Date).ToString("yyMMdd"))",

        [parameter( Mandatory=$false, Position=7)]
        [System.Int32]
        $EndDate = "$((Get-Date).ToString("yyMMdd"))",

        [parameter( Mandatory=$false, Position=8)]
        [System.String]
        [ValidateSet('Incoming','Outgoing')]
        $Direction = 'Incoming',

        [parameter( Mandatory=$false, Position=9)]
        [ValidateScript({if (Test-Path $_ -PathType leaf) {$true} else {throw "Logparser could not be found!"}})]
        [System.String]
        $Logparser="C:\Program Files (x86)\Log Parser 2.2\LogParser.exe"
    )

    begin
    {
        # get logs
        switch ($Direction)
        {
            'Incoming' {$smptLogPath = 'C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Edge\ProtocolLog\SmtpReceive'}
            'Outgoing' {$smptLogPath = 'C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Edge\ProtocolLog\SmtpSend'}
        }

        $allLogs = Get-ChildItem -Path $smptLogPath -Include *.LOG -Recurse
        # filter logs based on date
        $allLogs = $allLogs | Where-Object { $_.Name.SubString(6,6) -ge $StartDate -and $_.Name.SubString(6,6) -le $EndDate }
        [System.String[]]$logsFrom = @()
        $logsFrom = $allLogs.FullName -join "','"
        $logsFrom = "'" + $logsFrom + "'"
        Write-Verbose "Found $($allLogs.Count) log files..."
    }

    process
    {
        if ($ByRemoteIP)
        {
            Write-Verbose "Get statistics by Remote IP address..."
            $stamp = 'ByRemoteIP_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')
            $query = @"
            SELECT DISTINCT RemoteIP,RemotePort,RemoteDNS,Count(*) as Hits
            USING
            EXTRACT_PREFIX(remote-endpoint,0,':') as RemoteIP,
            EXTRACT_SUFFIX(remote-endpoint,0,':') as RemotePort,
            REVERSEDNS(EXTRACT_PREFIX(remote-endpoint,0,':')) as RemoteDNS
            INTO $Outpath\$stamp.csv 
            FROM $logsFrom
            WHERE data LIKE 'RCPT%'
            GROUP BY RemoteIP,RemotePort,RemoteDNS
            ORDER BY Hits DESC
"@
        }

        if ($DetailsForRemoteIP)
        {
            Write-Verbose "Get details for Remote IP address..."
            $stamp = 'DetailsForRemoteIP_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')
            $DetailsForRemoteIPString = "'" + $($DetailsForRemoteIP -join "';'") + "'"
            $query = @"
            SELECT day,time,[connector-id],[session-id],[sequence-number],[local-endpoint],[remote-endpoint],event,data,context
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time,
            EXTRACT_PREFIX(remote-endpoint,0,':') as RemoteIP
            INTO $Outpath\$stamp.csv 
            FROM $logsFrom
            WHERE RemoteIP IN ($DetailsForRemoteIPString) AND data LIKE 'RCPT%'
            GROUP BY day,time,[connector-id],[session-id],[sequence-number],[local-endpoint],[remote-endpoint],event,data,context
            ORDER BY day,time ASC
"@
        }

        if ($DataContains)
        {
            Write-Verbose "Search data field for $($DataContains)..."
            $stamp = 'DataContains_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,[connector-id],[session-id],[sequence-number],[local-endpoint],[remote-endpoint],event,data,context,Filename
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time,
            EXTRACT_PREFIX(remote-endpoint,0,':') as RemoteIP
            INTO $Outpath\$stamp.csv 
            FROM $logsFrom
            WHERE data LIKE '%$($DataContains)%'
            GROUP BY day,time,[connector-id],[session-id],[sequence-number],[local-endpoint],[remote-endpoint],event,data,context,Filename
            ORDER BY day,time ASC
"@
        }

        if ($ContextContains)
        {
            Write-Verbose "Search context field for $($ContextContains)..."
            $stamp = 'ContextContains_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,[connector-id],[session-id],[sequence-number],[local-endpoint],[remote-endpoint],event,data,context,Filename
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time,
            EXTRACT_PREFIX(remote-endpoint,0,':') as RemoteIP
            INTO $Outpath\$stamp.csv 
            FROM $logsFrom
            WHERE context LIKE '%$($ContextContains)%'
            GROUP BY day,time,[connector-id],[session-id],[sequence-number],[local-endpoint],[remote-endpoint],event,data,context,Filename
            ORDER BY day,time ASC
"@
        }

        if ($SessionID)
        {
            Write-Verbose "Search session-id for $($SessionID)..."
            $stamp = 'SessionID' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,[connector-id],[session-id],[sequence-number],[local-endpoint],[remote-endpoint],event,data,context,Filename
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time,
            EXTRACT_PREFIX(remote-endpoint,0,':') as RemoteIP
            INTO $Outpath\$stamp.csv 
            FROM $logsFrom
            WHERE TO_STRING([session-id]) IN ($("'" + $($SessionID -join "';'") + "'"))
            GROUP BY day,time,[connector-id],[session-id],[sequence-number],[local-endpoint],[remote-endpoint],event,data,context,Filename
            ORDER BY day,time ASC
"@
        }
        # workaround for limitation of path length, therefore we put the query into a file
        Set-Content -Value $query -Path $outpath\query.txt -Force
        Write-Output -InputObject "Starting query!"
        & $Logparser file:$outpath\query.txt -i:csv -o:csv -nSkipLines:4
        Write-Output -InputObject "Query done!"
    }

    end
    {
        Write-Verbose "Done!"
    }

}

function global:Get-MessageTrackingStats
{
    [CmdletBinding()]
    param(
        [parameter( Mandatory=$false, Position=0)]
        [System.String]
        $SourceContextContains,

        [parameter( Mandatory=$false, Position=1)]
        [System.String]
        $SourceContextNotContains,

        [parameter( Mandatory=$false, Position=3)]
        [System.Management.Automation.SwitchParameter]
        $ByEventID,

        [parameter( Mandatory=$false, Position=4)]
        [System.Int16]
        $ByEventIDIntervall = '3600',

        [parameter( Mandatory=$false, Position=4)]
        [System.String[]]
        $InEventID,

        [parameter( Mandatory=$false, Position=4)]
        [ValidateScript({if (Test-Path $_ -PathType container) {$true} else {Throw "$_ is not a valid path!"}})]
        [System.String]
        $Outpath = $env:temp,

        [parameter( Mandatory=$false, Position=5)]
        [System.Int32]
        $StartDate = "$((Get-Date).ToString("yyMMdd"))",

        [parameter( Mandatory=$false, Position=6)]
        [System.Int32]
        $EndDate = "$((Get-Date).ToString("yyMMdd"))",

        [parameter( Mandatory=$false, Position=8)]
        [ValidateScript({if (Test-Path $_ -PathType leaf) {$true} else {throw "Logparser could not be found!"}})]
        [System.String]
        $Logparser="C:\Program Files (x86)\Log Parser 2.2\LogParser.exe",

        [parameter( Mandatory=$false, Position=8)]
        [System.String]
        $MSGTrackingLogPath = 'C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\MessageTracking'
    )

    begin
    {
        # get logs
        $allLogs = Get-ChildItem -Path $MSGTrackingLogPath -Include *.LOG -Recurse
        # filter logs based on date
        if (($StartDate.ToString().Length -gt 6) -or ($EndDate.ToString().Length -gt 6))
        {
            if (($StartDate.ToString().Length -gt 6) -and ($EndDate.ToString().Length -gt 6))
            {
                $allLogs = $allLogs | ?{$_.name.substring(8,8) -ge $StartDate -and $_.name.substring(8,8) -le $EndDate}
            }
            elseif (($StartDate.ToString().Length -gt 6) -and ($EndDate.ToString().Length -eq 6))
            {
                $allLogs = $allLogs | ?{$_.name.substring(8,8) -ge $StartDate -and $_.name.substring(8,6) -le $EndDate}
            }
            else
            {
                $allLogs = $allLogs | ?{$_.name.substring(8,6) -ge $StartDate -and $_.name.substring(8,6) -le $EndDate}
            }
        }
        else
        {
            $allLogs = $allLogs | ?{$_.name.substring(8,6) -ge $startdate -and $_.name.substring(8,6) -le $enddate}
        }
        #$allLogs = $allLogs | Where-Object { $_.Name.SubString(6,6) -ge $StartDate -and $_.Name.SubString(6,6) -le $EndDate }
        [System.String[]]$logsFrom = @()
        $logsFrom = $allLogs.FullName -join "','"
        $logsFrom = "'" + $logsFrom + "'"
        Write-Verbose "Found $($allLogs.Count) log files..."
    }

    process
    {

        if ($SourceContextContains)
        {
            Write-Verbose "Search SourceContext field for $($SourceContextContains)..."
            $stamp = 'SourceContextContains_' + $env:COMPUTERNAME + '_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,[client-ip],[client-hostname],[server-ip],[server-hostname],[source-context],[connector-id],source,[event-id],[internal-message-id],[message-id],[network-message-id],[recipient-address],[recipient-status],[total-bytes],[recipient-count],[related-recipient-address],reference,[message-subject],[sender-address],[return-path],[message-info],directionality,[tenant-id],[original-client-ip],[original-server-ip],[custom-data]
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time
            INTO $Outpath\$stamp.csv
            FROM $logsFrom
            WHERE [source-context] LIKE '%$($SourceContextContains)%'
            GROUP BY day,time,[client-ip],[client-hostname],[server-ip],[server-hostname],[source-context],[connector-id],source,[event-id],[internal-message-id],[message-id],[network-message-id],[recipient-address],[recipient-status],[total-bytes],[recipient-count],[related-recipient-address],reference,[message-subject],[sender-address],[return-path],[message-info],directionality,[tenant-id],[original-client-ip],[original-server-ip],[custom-data]
"@
        }

        if ($SourceContextNotContains)
        {
            Write-Verbose "Search SourceContext field for $($SourceContextNotContains)..."
            $stamp = 'SourceContextNotContains_' + $env:COMPUTERNAME + '_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,[client-ip],[client-hostname],[server-ip],[server-hostname],[source-context],[connector-id],source,[event-id],[internal-message-id],[message-id],[network-message-id],[recipient-address],[recipient-status],[total-bytes],[recipient-count],[related-recipient-address],reference,[message-subject],[sender-address],[return-path],[message-info],directionality,[tenant-id],[original-client-ip],[original-server-ip],[custom-data]
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time
            INTO $Outpath\$stamp.csv
            FROM $logsFrom
            WHERE [source-context] NOT LIKE '%$($SourceContextContains)%'
            GROUP BY day,time,[client-ip],[client-hostname],[server-ip],[server-hostname],[source-context],[connector-id],source,[event-id],[internal-message-id],[message-id],[network-message-id],[recipient-address],[recipient-status],[total-bytes],[recipient-count],[related-recipient-address],reference,[message-subject],[sender-address],[return-path],[message-info],directionality,[tenant-id],[original-client-ip],[original-server-ip],[custom-data]
"@
        }

        if ($ByEventID)
        {
            Write-Verbose "Search for EventIDs..."
            $stamp = 'ByEventID_' + $env:COMPUTERNAME + '_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,[event-id], Count(*) AS Count 
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            QUANTIZE(TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss'), $ByEventIDIntervall) AS time
            INTO $Outpath\$stamp.csv
            FROM $logsFrom
            WHERE day IS NOT NULL
            GROUP BY day,time,[event-id]
"@
        }

        if ($InEventID)
        {
            Write-Verbose "Search for EventID conatains $($EventIDContains)..."
            $stamp = 'EventIDContains_' + $env:COMPUTERNAME + '_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,[client-ip],[client-hostname],[server-ip],[server-hostname],[source-context],[connector-id],source,[event-id],[internal-message-id],[message-id],[network-message-id],[recipient-address],[recipient-status],[total-bytes],[recipient-count],[related-recipient-address],reference,[message-subject],[sender-address],[return-path],[message-info],directionality,[tenant-id],[original-client-ip],[original-server-ip],[custom-data]
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time
            INTO $Outpath\$stamp.csv
            FROM $logsFrom
            --WHERE [event-id] LIKE '%$($EventIDContains)%'
            WHERE [event-id] IN ( $(("'" +$($InEventID -join "';'") + "'").ToUpper()) )
            GROUP BY day,time,[client-ip],[client-hostname],[server-ip],[server-hostname],[source-context],[connector-id],source,[event-id],[internal-message-id],[message-id],[network-message-id],[recipient-address],[recipient-status],[total-bytes],[recipient-count],[related-recipient-address],reference,[message-subject],[sender-address],[return-path],[message-info],directionality,[tenant-id],[original-client-ip],[original-server-ip],[custom-data]
"@
        }

        # workaround for limitation of path length, therefore we put the query into a file
        Set-Content -Value $query -Path $outpath\query.txt -Force
        Write-Output -InputObject "Starting query!"
        & $Logparser file:$outpath\query.txt -i:csv -o:csv -nSkipLines:4
        Write-Output -InputObject "Query done!"
    }

    end
    {
        # clean query file
        Get-ChildItem -LiteralPath $Outpath -Filter query.txt | Remove-Item -Confirm:$false | Out-Null
        Write-Verbose "Done!"
    }

}

function global:ConvertTo-X500
{

    param(
    [parameter(Mandatory=$true)]
    [System.String]
    $IMCEAEX,

    [parameter(Mandatory=$false)]
    [System.Management.Automation.SwitchParameter]
    $URLEncoded
    )

    if ($URLEncoded)
    {
        Add-Type -AssemblyName System.Web
        $IMCEAEX = [System.Web.HttpUtility]::UrlDecode($IMCEAEX)
    }

    $IMCEAEX = $IMCEAEX.replace("_","/").replace("+20"," ").replace("+28","(").replace("+29",")").replace("+40","@").replace("+2E",".").replace("+2C",",").replace("+5F","_").replace("IMCEAEX-","X500:").split("@")[0]

    return $IMCEAEX
}

function global:Get-MSGraphServicePrincipal
{
    <#
        .SYNOPSIS
            This function retrieves properties of a ServicePrincipal.
        .DESCRIPTION
            This function retrieves properties and additional information e.g.: appRoleAssignments, oauth2PermissionGrants or owners.
        .PARAMETER ServicePrincipal
            The parameter ServicePrincipal defines the servicePrincipal to be queried.
        .PARAMETER AccessToken
            This required parameter AccessToken takes the Bearer access token for authentication the requests. The parameter takes a previously acquired access token.
        .PARAMETER Filter
            The parameter Filter can be used, when you want to use a complex filter.
        .PARAMETER ShowProgress
            The parameter ShowProgress will show the progress of the script.
        .PARAMETER Threads
            The parameter Threads defines how many Threads will be created. Only used in combination with MultiThread.
        .PARAMETER MultiThread
            The parameters MultiThread defines whether the script is running using multithreading.
        .PARAMETER Authority
            The authority from where you get the token.
        .PARAMETER ClientId
            Application ID of the registered app.
        .PARAMETER ClientSecret
            The secret, which is used for Client Credentials flow.
        .PARAMETER Certificate
            The certificate, which is used for Client Credentials flow.
        .PARAMETER MaxRetry
            How many retries for each SPN in case of error.
        .PARAMETER TimeoutSec
            TimeoutSec for Cmdlet Invoke-RestMethod.
        .EXAMPLE
            # retrieve data for a specific SPN by providing a previoues retreieved Bearer access token
            Get-MSGraphServicePrincipal -ServicePrincipal e1316a28-0ae7-4674-943b-17dcf20b4dd7 -AccessToken eyJ0eXAiOiJKV1QiLC...AwNzY5OTE4LCJuYmYiO -Verbose
            # retrieve data for a specific user and acquire an access token with ClientCredentials flow, where a certificate is used for client credential
            Get-MSGraphServicePrincipal -ServicePrincipal e1316a28-0ae7-4674-943b-17dcf20b4dd7 -Authority "https://login.microsoftonline.com/42f...791af7/oauth2/v2.0/token" -ClientId 1961a...2f8ab0 -Certificate $(dir Cert:\CurrentUser\My\ | ? thumbprint -eq 'F5A9C6...C2239D5112C73')
            # retrive data for multiple users using multithreading and filter
            Get-MSGraphServicePrincipal -Filter "StartsWith(displayname,'Microsoft')" -MultiThread -Authority "https://login.microsoftonline.com/42f...791af7/oauth2/v2.0/token" -ClientId 1961a...2f8ab0 -Certificate $(dir Cert:\CurrentUser\My\ | ? thumbprint -eq 'F5A9C6...C2239D5112C73')
        .NOTES
            If you want to leverage all functionality you will need to provide an access token with the following claims:
                Directory.Read.All
        .LINK
            https://docs.microsoft.com/graph/api/resources/user?view=graph-rest-beta
            https://docs.microsoft.com/graph/paging
            https://docs.microsoft.com/graph/json-batching
            https://docs.microsoft.com/graph/query-parameters
            https://docs.microsoft.com/graph/permissions-reference
            https://docs.oasis-open.org/odata/odata/v4.0/errata03/os/complete/part2-url-conventions/odata-v4.0-errata03-os-part2-url-conventions-complete.html#_Toc453752358
            https://docs.microsoft.com/en-us/graph/throttling
            https://docs.microsoft.com/graph/errors
            https://docs.microsoft.com/graph/best-practices-concept#handling-expected-errors
    #>
    [CmdletBinding()]
    param(
        [parameter( Position=0)]
        [System.String[]]
        $ServicePrincipal,

        [parameter( Position=1)]
        [System.String]
        $AccessToken,

        [parameter( Position=2)]
        [System.String]
        $Filter,

        [parameter( Position=3)]
        [System.Management.Automation.SwitchParameter]
        $ListAll,

        [parameter( Position=4)]
        [System.Management.Automation.SwitchParameter]
        $ReturnAppRoleAssignedTo,

        [parameter( Position=5)]
        [System.Management.Automation.SwitchParameter]
        $ReturnOauth2PermissionGrants,

        [parameter( Position=6)]
        [System.Management.Automation.SwitchParameter]
        $ShowProgress,

        [parameter( Position=7)]
        [System.Int16]
        $Threads = '20',

        [parameter( Position=8)]
        [System.Management.Automation.SwitchParameter]
        $MultiThread,

        [parameter( Position=9)]
        [System.String]
        $Authority,

        [parameter( Position=10)]
        [System.String]
        $ClientId,

        [parameter( Position=11)]
        [System.String]
        $ClientSecret,

        [parameter( Position=12)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $Certificate,

        [parameter( Position=13)]
        [System.Int16]
        $MaxRetry = '3',

        [parameter( Position=14)]
        [System.Int32]
        $TimeoutSec = '15',

        [parameter( Position=15)]
        [System.Management.Automation.SwitchParameter]
        $ApplicationPermissionsReport,

        [parameter( Position=16)]
        [System.Int32]
        $MaxFilterResult

    )

    begin
    {

        $timer = [System.Diagnostics.Stopwatch]::StartNew()

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $Error.Clear()

        $collection = [System.Collections.ArrayList]@()

        function Get-AccessTokenNoLibraries
        {
            <#
                .SYNOPSIS
                    This function acquires an access token using either shared secret or certificate.
                .DESCRIPTION
                    This function will acquire an access token from Azure AD using either shared secret or certificate and OAuth2.0 client credentials flow without the need of having any DLLs installed.
                .PARAMETER Authority
                    The parameter AZKeyVaultBaseUri is required and is the base uri of the Azure Key Vault.
                .PARAMETER ClientID
                    The parameter ClientID is required and defines the registered application.
                .PARAMETER ClientSecret
                    The parameter ClientSecret is optional and used for ClientCredential flow.
                .PARAMETER Certificate
                    The parameter Certificate is required, when a certificate is used and for ClientCredential flow. You need to provide a X509Certificate2 object.
                .EXAMPLE
                    Get-AccessTokenNoLibraries -Authority = 'https://login.microsoftonline.com/53g7676c-f466-423c-82f6-vb2f55791wl7/oauth2/v2.0/token' -ClientID '1521a4b6-v487-4078-be14-c9fvb17f8ab0' -Certificate $(dir Cert:\CurrentUser\My\ | ? thumbprint -eq 'F5B5C6135719644F5582BC184B4C2239D5112C73')
                    Get-AccessTokenNoLibraries -Authority = 'https://login.microsoftonline.com/53g7676c-f466-423c-82f6-vb2f55791wl7/oauth2/v2.0/token' -ClientID '1521a4b6-v487-4078-be14-c9fvb17f8ab0' -ClientSecret 'gavuCG5Pm__cG4F1~SN_03KQDAN2N~7o.L'
                .NOTES
                    
                .LINK
                    https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-certificate-credentials
                    https://adamtheautomator.com/microsoft-graph-api-powershell/#Acquire_an_Access_Token_(Using_a_Certificate)
                    https://adamtheautomator.com/microsoft-graph-api-powershell/#Acquire_an_Access_Token_(Application_Id_and_Secret)
            #>

            [CmdletBinding()]
            Param
            (
                [System.String]
                $Authority,

                [System.String]
                $ClientId,

                [System.String]
                $ClientSecret,

                [System.Security.Cryptography.X509Certificates.X509Certificate2]
                $Certificate
            )

            begin
            {
                $Scope = "https://graph.microsoft.com/.default"

                if ($Certificate)
                {

                    # Create base64 hash of certificate
                    $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())

                    # Create JWT timestamp for expiration
                    $StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
                    $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
                    $JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0)

                    # Create JWT validity start timestamp
                    $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
                    $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan,0)

                    # Create JWT header
                    $JWTHeader = @{
                        alg = 'RS256'
                        typ = 'JWT'
                        # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
                        x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='
                    }

                    # Create JWT payload
                    $JWTPayLoad = @{
                        # What endpoint is allowed to use this JWT
                        aud = $Authority # "https://login.microsoftonline.com/$TenantName/oauth2/token"

                        # Expiration timestamp
                        exp = $JWTExpiration

                        # Issuer = your application
                        iss = $ClientId

                        # JWT ID: random guid
                        jti = [System.Guid]::NewGuid()

                        # Not to be used before
                        nbf = $NotBefore

                        # JWT Subject
                        sub = $ClientId
                    }

                    # Convert header and payload to base64
                    $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
                    $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

                    $JWTPayLoadToByte =  [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
                    $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

                    # Join header and Payload with "." to create a valid (unsigned) JWT
                    $JWT = $EncodedHeader + "." + $EncodedPayload

                    # Get the private key object of your certificate
                    $PrivateKey = $Certificate.PrivateKey

                    # Define RSA signature and hashing algorithm
                    $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
                    $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

                    # Create a signature of the JWT
                    $Signature = [Convert]::ToBase64String(
                        $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT),$HashAlgorithm,$RSAPadding)
                    ) -replace '\+','-' -replace '/','_' -replace '='

                    # Join the signature to the JWT with "."
                    $JWT = $JWT + "." + $Signature

                    # Create a hash with body parameters
                    $Body = @{
                        client_id = $ClientId
                        client_assertion = $JWT
                        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                        scope = $Scope
                        grant_type = 'client_credentials'
                    
                    }

                    # Use the self-generated JWT as Authorization
                    $Header = @{
                        Authorization = "Bearer $JWT"
                    }

                }
                else
                {
                    # Create body
                    $Body = @{
                        client_id = $ClientId
                        client_secret = $ClientSecret
                        scope = $Scope
                        grant_type = 'client_credentials'
                    }

                }

                # Splat the parameters for Invoke-Restmethod for cleaner code
                $PostSplat = @{
                    ContentType = 'application/x-www-form-urlencoded'
                    Method = 'POST'
                    Body = $Body
                    Uri = $Authority
                }

                if ($Certificate)
                {
                    Write-Verbose 'Adding headers for certificate based request...'
                    $PostSplat.Add('Headers',$Header)
                }

            }

            process
            {
                try {
                    $accessToken = Invoke-RestMethod @PostSplat
                }
                catch{
                    $_
                }
            }

            end
            {
                $accessToken
            }
        }

        if ([System.String]::IsNullOrWhiteSpace($AccessToken))
        {
            # check for necessary parameters
            if ( (($Authority -and $ClientId) -and $ClientSecret) -or (($Authority -and $ClientId) -and $Certificate) )
            {
                Write-Verbose 'No AccessToken provided. Will acquire silently...'
                $acquireToken = @{
                    Authority = $Authority
                    ClientId = $ClientId
                    ClientSecret = $ClientSecret
                    Certificate = $Certificate
                }

                $AccessToken = (Get-AccessTokenNoLibraries @acquireToken).access_token

                $global:token = $AccessToken

                if ([System.String]::IsNullOrWhiteSpace($AccessToken))
                {
                    Write-Verbose "Couldn't acquire a token. Will stop now..."
                    break
                }

                Write-Verbose 'Setting accesstoken...'
                $psBoundParameters['AccessToken'] = $AccessToken
            }
            else
            {
                Write-Host 'No accesstoken provided and not enough data for acquiring one. Will exit...'
                break
            }
        }

        if ($Filter -or $ListAll)
        {

            [System.Boolean]$retryRequest = $true

            do {
                try {

                    if ($ListAll)
                    {
                        Write-Verbose 'Will list all ServicePrincipal...'
                        $URI = 'https://graph.microsoft.com/beta/servicePrincipals?$select=id&$count=true'
                    }
                    else
                    {
                        Write-Verbose 'Found custom Filter. Will try to find ServicePrincipal based on...'
                        $URI = 'https://graph.microsoft.com/beta/servicePrincipals?$filter='
                        $URI = $URI + $Filter
                    }

                    $filterParams = @{
                        ContentType = 'application/json'
                        Method = 'GET'
                        Headers = @{ Authorization = "Bearer $($AccessToken)"; ConsistencyLevel = "eventual"}
                        Uri = $URI
                        TimeoutSec = $TimeoutSec
                        ErrorAction = 'Stop'
                    }

                    $sPNResponse = Invoke-RestMethod @filterParams

                    if ($sPNResponse.'@odata.nextLink')
                    {
                        Write-Verbose 'Need to fetch more data...'
                        Write-Verbose "Totalcount:$($sPNResponse.'@odata.count')"
                        [System.Int16]$counter = '2'
                        # create collection
                        $sPNCollection = [System.Collections.ArrayList]@()

                        # add first batch of groups to collection
                        if ($sPNResponse.Value)
                        {
                            $sPNCollection += $sPNResponse.Value
                        }
                        else
                        {
                            $sPNCollection += $sPNResponse
                        }

                        do
                        {
                            $filterParams = @{
                                ContentType = 'application/json'
                                Method = 'GET'
                                Headers = @{ Authorization = "Bearer $($AccessToken)"; ConsistencyLevel = "eventual"}
                                Uri = $($sPNResponse.'@odata.nextLink')
                                TimeoutSec = $TimeoutSec
                                ErrorAction = 'Stop'
                                Verbose = $false
                            }

                            $sPNResponse = Invoke-RestMethod @filterParams

                            $sPNCollection += $sPNResponse.Value

                            Write-Verbose "Pagecount:$($counter)..."
                            $counter++

                            if ( $MaxFilterResult -and ($sPNCollection.Count -ge $MaxFilterResult))
                            {
                                Write-Verbose "MaxFilterResult reached. Will stop searching now..."
                                $sPNResponse.'@odata.nextLink' = $null
                                $sPNCollection = $sPNCollection  | Select-Object -First $MaxFilterResult
                            }

                        } while ($sPNResponse.'@odata.nextLink')

                        $ServicePrincipal = $sPNCollection.id

                    }
                    else
                    {
                        $ServicePrincipal = $sPNResponse.value.id

                    }

                    $retryRequest = $false

                    Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"

                }
                catch
                {
                    $_
                    break
                }
            } while ($retryRequest)

            if ($ServicePrincipal.count -eq 0)
            {
                Write-Verbose $('No ServicePrincipal found for filter "' + $($Filter) + '"! Terminate now...')
                break
            }
            else
            {
                Write-Host "Found $($ServicePrincipal.count) ServicePrincipals for provided filter..."
            }
        }

        if ($MultiThread)
        {
            #initiate runspace and make sure we are using single-threaded apartment STA
            $Jobs = @()
            $Sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $Threads, $Sessionstate, $Host)
            $RunspacePool.ApartmentState = 'STA'
            $RunspacePool.Open()
        }

        # initiate counter
        [System.Int32]$j = '1'

    }

    process
    {
        if ($MultiThread)
        {
            Write-Verbose 'Will run multithreaded...'
            #create scriptblock from function
            $scriptBlock = [System.Management.Automation.ScriptBlock]::Create((Get-ChildItem Function:\Get-MSGraphServicePrincipal).Definition)
            
            ForEach ($account in $ServicePrincipal)
            {
                 $progressParams = @{
                    Activity = "Adding ServicePrincipal to jobs and starting job - $($account)"
                    PercentComplete = $j / $ServicePrincipal.count * 100
                    Status = "Remaining: $($ServicePrincipal.count - $j) out of $($ServicePrincipal.count)"
                 }
                 Write-Progress @progressParams

                 $j++

                 try {
                    $PowershellThread = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock).AddParameter('ServicePrincipal',$account)
                    [void]$PowershellThread.AddParameter('AccessToken',$AccessToken)
                    [void]$PowershellThread.AddParameter('Authority',$Authority)
                    [void]$PowershellThread.AddParameter('ClientId',$ClientId)
                    [void]$PowershellThread.AddParameter('ClientSecret',$ClientSecret)
                    [void]$PowershellThread.AddParameter('Certificate',$Certificate)
                    [void]$PowershellThread.AddParameter('MaxRetry',$MaxRetry)
                    [void]$PowershellThread.AddParameter('TimeoutSec',$TimeoutSec)
                    [void]$PowershellThread.AddParameter('ShowProgress',$false)
                    [void]$PowershellThread.AddParameter('ReturnAppRoleAssignedTo',$ReturnAppRoleAssignedTo)
                    [void]$PowershellThread.AddParameter('ReturnOauth2PermissionGrants',$ReturnOauth2PermissionGrants)
                    [void]$PowershellThread.AddParameter('ApplicationPermissionsReport',$ApplicationPermissionsReport)

                    $PowershellThread.RunspacePool = $RunspacePool
                    $Handle = $PowershellThread.BeginInvoke()
                    $Job = "" | Select-Object Handle, Thread, Object
                    $Job.Handle = $Handle
                    $Job.Thread = $PowershellThread
                    $Job.Object = $account
                    $Jobs += $Job
                 }
                 catch {
                    $_
                 }
            }

            $progressParams.Status = "Ready"
            $progressParams.Add('Completed',$true)
            Write-Progress @progressParams

        }
        else
        {

            foreach ($account in $ServicePrincipal)
            {

                [System.Int16]$retryCount = '0'
                [System.Boolean]$keepTrying = $true

                do {

                    try {

                        $processingTime = [System.Diagnostics.Stopwatch]::StartNew()

                        if ($ShowProgress)
                        {
                            $progressParams = @{
                                Activity = "Processing ServicePrincipal - $($account)"
                                PercentComplete = $j / $ServicePrincipal.count * 100
                                Status = "Remaining: $($ServicePrincipal.count - $j) out of $($ServicePrincipal.count)"
                            }
                            Write-Progress @progressParams
                        }

                        $j++

                        if ($ApplicationPermissionsReport)
                        {
                            $body = @{
                                requests = @(
                                    @{
                                        url = "/servicePrincipals/$($account)"
                                        method = 'GET'
                                        id = '1'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/appRoleAssignments"
                                        method = 'GET'
                                        id = '3'
                                    }
                                )
                            }
                        }
                        else
                        {

                            $body = @{
                                requests = @(
                                    @{
                                        url = "/servicePrincipals/$($account)"
                                        method = 'GET'
                                        id = '1'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/appRoleAssignments"
                                        method = 'GET'
                                        id = '3'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/claimsMappingPolicies"
                                        method = 'GET'
                                        id = '4'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/createdObjects"
                                        method = 'GET'
                                        id = '5'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/delegatedPermissionClassifications"
                                        method = 'GET'
                                        id = '6'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/endpoints"
                                        method = 'GET'
                                        id = '7'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/homeRealmDiscoveryPolicies"
                                        method = 'GET'
                                        id = '8'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/memberOf"
                                        method = 'GET'
                                        id = '9'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/ownedObjects"
                                        method = 'GET'
                                        id = '11'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/owners"
                                        method = 'GET'
                                        id = '12'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/tokenIssuancePolicies"
                                        method = 'GET'
                                        id = '13'
                                    },
                                    @{
                                        url = "/servicePrincipals/$($account)/tokenLifetimePolicies"
                                        method = 'GET'
                                        id = '14'
                                    }
                                )
                            }

                        }

                        if ($ReturnAppRoleAssignedTo)
                        {
                            $appRoleAssignedTo = @{
                                    url = "/servicePrincipals/$($account)/appRoleAssignedTo"
                                    method = 'GET'
                                    id = '2'
                            }

                            $body.requests += $appRoleAssignedTo
                        }

                        if ($ReturnOauth2PermissionGrants)
                        {
                            $oauth2PermissionGrants = @{
                                    url = "/servicePrincipals/$($account)/oauth2PermissionGrants"
                                    method = 'GET'
                                    id = '10'
                            }

                            $body.requests += $oauth2PermissionGrants
                        }

                        $restParams = @{
                            ContentType = 'application/json'
                            Method = 'POST'
                            Headers = @{ Authorization = "Bearer $($AccessToken)"}
                            Body = $body | ConvertTo-Json -Depth 4
                            Uri = 'https://graph.microsoft.com/beta/$batch'
                            TimeoutSec = $TimeoutSec
                        }

                        $global:data = Invoke-RestMethod @restParams

                        # create custom object
                        $sPNObject = New-Object -TypeName psobject
                        $sPNInfo = $null
                        $sPNInfo = ($data.responses | Where-Object -FilterScript { $_.id -eq 1}).Body | Select-Object * -ExcludeProperty "@odata.context"
                        if ('ResourceNotFound' -eq $sPNInfo.error.code)
                        {
                            Write-Verbose "ResourceNotFound thrown for:$($account)..."
                            $sPNObject = New-Object -TypeName psobject
                            $sPNObject | Add-Member -MemberType NoteProperty -Name ServicePrincipal -Value $account
                            $sPNObject | Add-Member -MemberType NoteProperty -Name Error -Value 'ResourceNotFound'
                            $collection += $sPNObject
                            break
                        }

                        $sPNProperties = $sPNInfo | Get-Member -MemberType NoteProperty

                        foreach ($property in $sPNProperties)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name $( $property.Name ) -Value $( $sPNInfo.$( $property.Name ) )
                        }

                        # add appRoleAssigned to object
                        $appRoleAssignedToResponse = $data.responses | Where-Object -FilterScript {$_.id -eq 2}

                        if ('200' -eq $appRoleAssignedToResponse.status)
                        {
                            if ($appRoleAssignedToResponse.body.'@odata.nextLink')
                            {

                                Write-Verbose 'Need to fetch more data for appRoleAssignedTo...'
                                [System.Int16]$counter = '2'
                                # create collection
                                $AppRoleAssignedToCollection = [System.Collections.ArrayList]@()

                                # add first batch of groups to collection
                                $AppRoleAssignedToCollection += $appRoleAssignedToResponse.body.value | Select-Object * -ExcludeProperty "@odata.type"

                                do
                                {
                                    if ($appRoleAssignedToResponse.body.'@odata.nextLink')
                                    {
                                        $uriNext = $appRoleAssignedToResponse.body.'@odata.nextLink'
                                    }
                                    else
                                    {
                                        $uriNext = $appRoleAssignedToResponse.'@odata.nextLink'
                                    }
                                    
                                    $AppRoleAssignedToParams = @{
                                        ContentType = 'application/json'
                                        Method = 'GET'
                                        Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                        Uri = $uriNext
                                        TimeoutSec = $TimeoutSec
                                        Verbose = $false
                                    }

                                    $appRoleAssignedToResponse = Invoke-RestMethod @AppRoleAssignedToParams

                                    if ($appRoleAssignedToResponse.body.value)
                                    {
                                        $AppRoleAssignedToCollection += $appRoleAssignedToResponse.body.value | Select-Object * -ExcludeProperty "@odata.type"
                                    }
                                    else
                                    {
                                        $AppRoleAssignedToCollection += $appRoleAssignedToResponse.value | Select-Object * -ExcludeProperty "@odata.type"
                                    }

                                    Write-Verbose "Pagecount:$($counter)..."
                                    $counter++

                                } while ($appRoleAssignedToResponse.'@odata.nextLink')

                                $sPNObject | Add-Member -MemberType NoteProperty -Name AppRoleAssignedTo -Value @( $AppRoleAssignedToCollection )

                            }
                            else
                            {
                                $sPNObject | Add-Member -MemberType NoteProperty -Name AppRoleAssignedTo -Value @( $($appRoleAssignedToResponse.body.value | Select-Object * -ExcludeProperty "@odata.Context","@odata.type") )
                            }
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name AppRoleAssignedTo -Value @( $($appRoleAssignedToResponse.body.error) )
                        }

                        # extract appRoleAssignments response
                        $responseappRoleAssignments = $data.responses | Where-Object -FilterScript {$_.id -eq 3}

                        if ('200' -eq $responseappRoleAssignments.status)
                        {
                            if ($responseappRoleAssignments.body.'@odata.nextLink')
                            {

                                Write-Verbose 'Need to fetch more data for appRoleAssignments...'
                                [System.Int16]$counter = '2'
                                # create collection
                                $appRoleAssignmentsCollection = [System.Collections.ArrayList]@()

                                # add first batch of groups to collection
                                $appRoleAssignmentsCollection += $responseappRoleAssignments.body.value | Select-Object * -ExcludeProperty "@odata.type"

                                do
                                {
                                    $appRoleAssignmentsParams = @{
                                        ContentType = 'application/json'
                                        Method = 'GET'
                                        Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                        Uri = $responseappRoleAssignments.body.'@odata.nextLink'
                                        TimeoutSec = $TimeoutSec
                                    }

                                    $responseappRoleAssignments = Invoke-RestMethod @appRoleAssignmentsParams

                                    if ($responseappRoleAssignments.body.value)
                                    {
                                        $appRoleAssignmentsCollection += $responseappRoleAssignments.body.value | Select-Object * -ExcludeProperty "@odata.type"
                                    }
                                    else
                                    {
                                        $appRoleAssignmentsCollection += $responseappRoleAssignments.value | Select-Object * -ExcludeProperty "@odata.type"
                                    }

                                    Write-Verbose "Pagecount:$($counter)..."
                                    $counter++

                                } while ($responseappRoleAssignments.body.'@odata.nextLink')

                                $sPNObject | Add-Member -MemberType NoteProperty -Name AppRoleAssignments -Value @( $appRoleAssignmentsCollection )

                            }
                            else
                            {
                                $sPNObject | Add-Member -MemberType NoteProperty -Name AppRoleAssignments -Value @( $($responseappRoleAssignments.body.value | Select-Object * -ExcludeProperty "@odata.type") )
                            }

                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name AppRoleAssignments -Value @( $($responseappRoleAssignments.body.error) )
                        }

                        $claimsMappingPoliciesResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 4}

                        if ('200' -eq $claimsMappingPoliciesResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name ClaimsMappingPolicies -Value  @( $($claimsMappingPoliciesResponse.body.value) )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name ClaimsMappingPolicies -Value  @( $($claimsMappingPoliciesResponse.body.error) )
                        }

                        $createdObjectsResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 5}

                        if ('200' -eq $createdObjectsResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name CreatedObjects -Value  @( $($createdObjectsResponse.body.value | Select-Object * -ExcludeProperty "@odata.type") )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name CreatedObjects -Value  @( $($createdObjectsResponse.body.error) )
                        }

                        $delegatedPermissionClassificationsResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 6}

                        if ('200' -eq $delegatedPermissionClassificationsResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name DelegatedPermissionClassifications -Value  @( $($delegatedPermissionClassificationsResponse.body.value | Select-Object * -ExcludeProperty "@odata.type") )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name DelegatedPermissionClassifications -Value  @( $($delegatedPermissionClassificationsResponse.body.error) )
                        }

                        $endpointsResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 7}

                        if ('200' -eq $endpointsResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name Endpoints -Value  @( $($endpointsResponse.body.value) )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name Endpoints -Value  @( $($endpointsResponse.body.error) )
                        }

                        $homeRealmDiscoveryPoliciesResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 8}

                        if ('200' -eq $homeRealmDiscoveryPoliciesResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name HomeRealmDiscoveryPolicies -Value  @( $($homeRealmDiscoveryPoliciesResponse.body.value) )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name HomeRealmDiscoveryPolicies -Value  @( $($homeRealmDiscoveryPoliciesResponse.body.error) )
                        }

                        $memberOfResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 9}

                        if ('200' -eq $memberOfResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name MemberOf -Value  @( $($memberOfResponse.body.value) )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name MemberOf -Value  @( $($memberOfResponse.body.error) )
                        }


                        $oauth2PermissionGrantsResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 10}

                        if ('200' -eq $oauth2PermissionGrantsResponse.status)
                        {
                            if ($oauth2PermissionGrantsResponse.body.'@odata.nextLink')
                            {

                                Write-Verbose 'Need to fetch more data for oauth2PermissionGrants...'
                                [System.Int16]$counter = '2'
                                # create collection
                                $oauth2PermissionGrantsCollection = [System.Collections.ArrayList]@()

                                # add first batch of groups to collection
                                $oauth2PermissionGrantsCollection += $oauth2PermissionGrantsResponse.body.value | Select-Object * -ExcludeProperty "@odata.type"

                                do
                                {
                                    if ($oauth2PermissionGrantsResponse.body.'@odata.nextLink')
                                    {
                                        $uriNext = $oauth2PermissionGrantsResponse.body.'@odata.nextLink'
                                    }
                                    else
                                    {
                                        $uriNext = $oauth2PermissionGrantsResponse.'@odata.nextLink'
                                    }
                                    
                                    $oauth2PermissionGrantsParams = @{
                                        ContentType = 'application/json'
                                        Method = 'GET'
                                        Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                        Uri = $uriNext
                                        TimeoutSec = $TimeoutSec
                                        Verbose = $false
                                    }

                                    $oauth2PermissionGrantsResponse = Invoke-RestMethod @oauth2PermissionGrantsParams

                                    if ($oauth2PermissionGrantsResponse.body.value)
                                    {
                                        $oauth2PermissionGrantsCollection += $oauth2PermissionGrantsResponse.body.value | Select-Object * -ExcludeProperty "@odata.type"
                                    }
                                    else
                                    {
                                        $oauth2PermissionGrantsCollection += $oauth2PermissionGrantsResponse.value | Select-Object * -ExcludeProperty "@odata.type"
                                    }

                                    Write-Verbose "Pagecount:$($counter)..."
                                    $counter++

                                } while ($oauth2PermissionGrantsResponse.'@odata.nextLink')

                                $sPNObject | Add-Member -MemberType NoteProperty -Name OAuth2PermissionGrants -Value @( $oauth2PermissionGrantsCollection )

                            }
                            else
                            {
                                $sPNObject | Add-Member -MemberType NoteProperty -Name OAuth2PermissionGrants -Value  @( $($oauth2PermissionGrantsResponse.body.value | Select-Object * -ExcludeProperty "@odata.context") )
                            }
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name OAuth2PermissionGrants -Value  @( $($oauth2PermissionGrantsResponse.body.error) )
                        }

                        $ownedObjectsResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 11}

                        if ('200' -eq $ownedObjectsResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name OwnedObjects -Value  @( $($ownedObjectsResponse.body.value | Select-Object * -ExcludeProperty "@odata.context") )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name OwnedObjects -Value  @( $($ownedObjectsResponse.body.error) )
                        }

                        $ownersResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 12}

                        if ('200' -eq $ownersResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name Owners -Value  @( $($ownersResponse.body.value | Select-Object * -ExcludeProperty "@odata.context") )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name Owners -Value  @( $($ownersResponse.body.error) )
                        }

                        $tokenIssuancePoliciesResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 13}

                        if ('200' -eq $tokenIssuancePoliciesResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name TokenIssuancePolicies -Value  @( $($tokenIssuancePoliciesResponse.body.value | Select-Object * -ExcludeProperty "@odata.context") )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name TokenIssuancePolicies -Value  @( $($tokenIssuancePoliciesResponse.body.error) )
                        }

                        $tokenLifetimePoliciesResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 14}

                        if ('200' -eq $tokenLifetimePoliciesResponse.status)
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name TokenLifetimePolicies -Value  @( $($tokenLifetimePoliciesResponse.body.value | Select-Object * -ExcludeProperty "@odata.context") )
                        }
                        else
                        {
                            $sPNObject | Add-Member -MemberType NoteProperty -Name TokenLifetimePolicies -Value  @( $($tokenLifetimePoliciesResponse.body.error) )
                        }

                        $collection += $sPNObject

                        if ($ShowProgress)
                        {
                            $progressParams.Status = "Ready"
                            $progressParams.Add('Completed',$true)
                            Write-Progress @progressParams
                        }

                        $processingTime.Stop()
                        $keepTrying = $false
                        Write-Verbose "ServicePrincipal processing time:$($processingTime.Elapsed.ToString())"

                    }
                    catch {

                    $statusCode = $_.Exception.Response.StatusCode.value__

                    if ('401' -eq $statusCode)
                    {

                        Write-Verbose "HTTP error $($statusCode) thrown. Will try to acquire new access token..."
                        $acquireToken = @{
                            Authority = $Authority
                            ClientId = $ClientId
                            ClientSecret = $ClientSecret
                            Certificate = $Certificate
                        }

                        $AccessToken = (Get-AccessTokenNoLibraries @acquireToken).access_token

                        if ([System.String]::IsNullOrWhiteSpace($AccessToken))
                        {
                            Write-Verbose "Couldn't acquire a token. Will stop now..."

                            $sPNObject = New-Object -TypeName psobject
                            $sPNObject | Add-Member -MemberType NoteProperty -Name User -Value $account
                            $sPNObject | Add-Member -MemberType NoteProperty -Name Error -Value 'Could not acquire access token'
                            $collection += $sPNObject
                            break
                        }

                        Write-Verbose 'Setting accesstoken...'
                        $psBoundParameters['AccessToken'] = $AccessToken

                        # increase retrycount
                        $retryCount++
                        Write-Verbose "Retrycount:$($retryCount)"

                    }
                    elseif ('403' -eq $statusCode)
                    {
                        Write-Verbose "Error for resource $($_.Exception.Response.ResponseUri.ToString()) with StatusCode $($_.Exception.Response.StatusCode.ToString())"
                        # increase retrycount
                        $retryCount++
                    }
                    elseif ('429' -eq $statusCode)
                    {

                        Write-Verbose "HTTP error $($statusCode) thrown. Will backoff..."
                        Start-Sleep -Seconds 2

                    }
                    elseif ('503' -eq $statusCode)
                    {

                        Write-Verbose "HTTP error $($statusCode) thrown. Will retry..."
                        # increase retrycount
                        $retryCount++
                        Write-Verbose "Retrycount:$($retryCount)"

                    }
                    else
                    {
                        $sPNObject = New-Object -TypeName psobject
                        $sPNObject | Add-Member -MemberType NoteProperty -Name ServicePrincipal -Value $account
                        $sPNObject | Add-Member -MemberType NoteProperty -Name Error -Value $_.Exception
                        $collection += $sPNObject
                        Write-Verbose "Error occured for account $($account)..."
                        break
                    }

                    if ($retryCount -eq $MaxRetry)
                    {
                        Write-Verbose 'MaxRetries done. Skip this ServicePrincipal now...'
                        $keepTrying = $false

                        $sPNObject = New-Object -TypeName psobject
                        $sPNObject | Add-Member -MemberType NoteProperty -Name ServicePrincipal -Value $account
                        $sPNObject | Add-Member -MemberType NoteProperty -Name Error -Value $_.Exception
                        $collection += $sPNObject
                    }
                }

                } while ($keepTrying)

            }

        }
    }

    end
    {
        #monitor and retrieve the created jobs
        if ($MultiThread)
        {
            $SleepTimer = 1000
            While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0) {

                Write-Progress `
                    -id 1 `
                    -Activity "Waiting for Jobs - $($Threads - $($RunspacePool.GetAvailableRunspaces())) of $Threads threads running" `
                    -PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
                    -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining out of $($Jobs.Count)"

                ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $true})) {
                    $Job.Thread.EndInvoke($Job.Handle)
                    $Job.Thread.Dispose()
                    $Job.Thread = $Null
                    $Job.Handle = $Null
                }

                # kill all incomplete threads when hit "CTRL+q"
                if ($Host.UI.RawUI.KeyAvailable) {
                    $KeyInput = $Host.UI.RawUI.ReadKey("IncludeKeyUp,NoEcho")
                    if (($KeyInput.ControlKeyState -cmatch '(Right|Left)CtrlPressed') -and ($KeyInput.VirtualKeyCode -eq '81')) {
                        Write-Host -fore red "Kill all incomplete threads..."
                            ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})) {
                                Write-Host -fore yellow "Stopping job $($Job.Object)..."
                                $Job.Thread.Stop()
                                $Job.Thread.Dispose()
                             }
                        Write-Host -fore red "All jobs terminated..."
                        break
                    }
                }

                Start-Sleep -Milliseconds $SleepTimer

            }

            # clean-up
            Write-Verbose 'Perform cleanup of runspaces...'
            $RunspacePool.Close() | Out-Null
            $RunspacePool.Dispose() | Out-Null
            [System.GC]::Collect()
        }
        else
        {
            $collection
        }

        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Get-ConnectivityLogs
{
    [CmdletBinding()]
    param(
        [parameter( Mandatory=$false, Position=0)]
        [System.String]
        $DescriptionContains,

        [parameter( Mandatory=$false, Position=1)]
        [System.String[]]
        $SessionID,

        [parameter( Mandatory=$false, Position=2)]
        [System.String]
        $DestinationContains,

        [parameter( Mandatory=$false, Position=3)]
        [ValidateScript({if (Test-Path $_ -PathType container) {$true} else {Throw "$_ is not a valid path!"}})]
        [System.String]
        $Outpath = $env:temp,

        [parameter( Mandatory=$false, Position=4)]
        [System.Int32]
        $StartDate = "$((Get-Date).ToString("yyMMdd"))",

        [parameter( Mandatory=$false, Position=5)]
        [System.Int32]
        $EndDate = "$((Get-Date).ToString("yyMMdd"))",

        [parameter( Mandatory=$false, Position=6)]
        [ValidateScript({if (Test-Path $_ -PathType leaf) {$true} else {throw "Logparser could not be found!"}})]
        [System.String]
        $Logparser="C:\Program Files (x86)\Log Parser 2.2\LogParser.exe"
    )

    begin
    {
        # get logs
        $connectivityLogs = 'C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Edge\Connectivity'
        $allLogs = Get-ChildItem -Path $connectivityLogs -Include *.LOG -Recurse
        # filter logs based on date
        $allLogs = $allLogs | Where-Object { $_.Name.SubString(12,6) -ge $StartDate -and $_.Name.SubString(12,6) -le $EndDate }
        [System.String[]]$logsFrom = @()
        $logsFrom = $allLogs.FullName -join "','"
        $logsFrom = "'" + $logsFrom + "'"
        Write-Verbose "Found $($allLogs.Count) log files..."
    }

    process
    {

        if ($DescriptionContains)
        {
            Write-Verbose "Search data field for $($DescriptionContains)..."
            $stamp = 'DescriptionContains_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,session,source,Destination,direction,description,Filename
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time
            INTO $Outpath\$stamp.csv 
            FROM $logsFrom
            WHERE description LIKE '%$($DescriptionContains)%'
            GROUP BY day,time,session,source,Destination,direction,description,Filename
            ORDER BY day,time ASC
"@
        }

        if ($SessionID)
        {
            Write-Verbose "Search session-id for $($SessionID)..."
            $stamp = 'SessionID' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,session,source,Destination,direction,description,Filename
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time
            INTO $Outpath\$stamp.csv 
            FROM $logsFrom
            WHERE TO_STRING(session) IN ($("'" + $($SessionID -join "';'") + "'"))
            GROUP BY day,time,session,source,Destination,direction,description,Filename
            ORDER BY day,time ASC
"@
        }

        if ($DestinationContains)
        {
            Write-Verbose "Search data field for $($DestinationContains)..."
            $stamp = 'DestinationContains_' + $((Get-Date (Get-Date).ToUniversalTime() -Format u) -replace ' |:','_')

            $query = @"
            SELECT day,time,session,source,Destination,direction,description,Filename
            USING
            TO_STRING(TO_TIMESTAMP(EXTRACT_PREFIX(REPLACE_STR([#Fields: date-time],'T',' '),0,'.'), 'yyyy-MM-dd hh:mm:ss'),'yyMMdd') AS day,
            TO_TIMESTAMP(EXTRACT_PREFIX(TO_STRING(EXTRACT_SUFFIX([#Fields: date-time],0,'T')),0,'.'), 'hh:mm:ss') AS time
            INTO $Outpath\$stamp.csv 
            FROM $logsFrom
            WHERE Destination LIKE '%$($DestinationContains)%'
            GROUP BY day,time,session,source,Destination,direction,description,Filename
            ORDER BY day,time ASC
"@
        }

        # workaround for limitation of path length, therefore we put the query into a file
        Set-Content -Value $query -Path $outpath\query.txt -Force
        Write-Output -InputObject "Starting query!"
        & $Logparser file:$outpath\query.txt -i:csv -o:csv -nSkipLines:4
        Write-Output -InputObject "Query done!"
    }

    end
    {
        Write-Verbose "Done!"
    }

}

function global:Search-MSTeamsUser
{
    <#
    .SYNOPSIS
        This function is intended to query Microsoft Teams substrate for give address.
    .DESCRIPTION
        This function is intended to query Microsoft Teams substrate for give address and return found objects.
    .PARAMETER EmailAddress
        The parameter EmailAddress defines the e-mail address, which we are searching.
    .EXAMPLE
        Search-MSTeamsUser -EmailAddress MeganB
    .NOTES

    .LINK
        https://ingogegenwarth.wordpress.com/
    #>
    [CmdletBinding()]
    param(
        [parameter(
            mandatory = $true,
            Position = 0)]
        [System.String]
        $EmailAddress

    )

    begin
    {
        $timer = [System.Diagnostics.Stopwatch]::StartNew()

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $Error.Clear()

        $collection = [System.Collections.ArrayList]@()

        $Timeout = '20'

        # thanks to https://gsexdev.blogspot.com/
        function Show-OAuthWindow
        {
            [CmdletBinding()]
            param (
                [System.Uri]
                $Url

            )
            ## Start Code Attribution
            ## Show-AuthWindow function is the work of the following Authors and should remain with the function if copied into other scripts
            ## https://foxdeploy.com/2015/11/02/using-powershell-and-oauth/
            ## https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/
            ## End Code Attribution
            Add-Type -AssemblyName System.Web
            Add-Type -AssemblyName System.Windows.Forms

            $form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width = 440; Height = 640 }
            $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width = 420; Height = 600; Url = ($url) }
            $Navigated = {
                if ( ($web.DocumentText -match "document.location.replace") -or ($web.Url.AbsoluteUri -match "code=[^&]*") ) {
                    $Script:oAuthCode = [regex]::match($web.DocumentText, "code=(.*?)\\u0026").Groups[1].Value
                    if ([System.String]::IsNullOrEmpty($Script:oAuthCode))
                    {
                        if ($web.Url.AbsoluteUri -match "error=[^&]*")
                        {
                            $Script:oAuthCode = [System.Web.HttpUtility]::UrlDecode($web.Url.AbsoluteUri)
                        }
                        else
                        {
                            $Script:oAuthCode = [System.Web.HttpUtility]::ParseQueryString($web.Url.AbsoluteUri)[0]
                        }
                    }
                    $form.Close();
                }
            }
            $web.ScriptErrorsSuppressed = $true
            $web.Add_Navigated($Navigated)
            $form.Controls.Add($web)
            $form.Add_Shown( { $form.Activate() })
            $form.ShowDialog() | Out-Null
            return $Script:oAuthCode
        }

        try {

            Add-Type -AssemblyName System.Web
            $ClientID = '1fec8e78-bce4-4aaf-ab1b-5451cc387264'
            $audience = 'https://teams.microsoft.com'
            # parameters for auth code flow
            $RedirectURI = 'https://login.microsoftonline.com/common/oauth2/nativeclient'
            $state = Get-Random
            $authURI = "https://login.microsoftonline.com/Common"
            $authURI += "/oauth2/authorize?client_id=$ClientId"
            $authURI += "&response_type=code&redirect_uri= " + [System.Web.HttpUtility]::UrlEncode($RedirectURI)
            $authURI += "&response_mode=query&resource=" + [System.Web.HttpUtility]::UrlEncode($audience) + "&state=$state"
            $authURI += "&prompt=select_account"

            # acquire auth code from AAD
            $authCode = Show-OAuthWindow -Url $authURI

            $body = @{
                "grant_type" = "authorization_code";
                "client_id" = "$ClientId";
                "code" = $authCode;
                "redirect_uri" = $RedirectURI
            }

            $paramsTokenRequest = @{
                Body = $body
                ContentType = 'application/x-www-form-urlencoded'
                Method = 'Post'
                TimeoutSec = $Timeout
                Uri = "https://login.microsoftonline.com/common/oauth2/token"
            }

            # acquire access token
            $tokenRequest = Invoke-RestMethod @paramsTokenRequest
            $accessToken = $tokenRequest.access_token

            # retrieve authz config
            $headersMSTeamsAuthz = @{}
            $headersMSTeamsAuthz.Add('Accept','application/json, text/plain, */*')
            $headersMSTeamsAuthz.Add('Accept-Encoding','gzip, deflate')
            $headersMSTeamsAuthz.Add('Authorization',"Bearer $($accessToken)")
            $headersMSTeamsAuthz.Add('Charset','utf-8')

            $paramsTeamsAuthz = @{
                ContentType = 'application/json; charset=utf-8'
                Headers = $headersMSTeamsAuthz
                Method = 'POST'
                TimeoutSec = $Timeout
                Uri = 'https://teams.microsoft.com/api/authsvc/v1.0/authz'
            }

            $teamsAuthZ = Invoke-RestMethod @paramsTeamsAuthz

        }
        catch {
            $_
            break
        }

    }

    process
    {
        if (-not [System.String]::IsNullOrEmpty($teamsAuthZ.regionGtms.middleTier))
        {
            try {
                $headersMSTeams = @{}
                $headersMSTeams.Add('Authorization',"Bearer $($accessToken)")
                $headersMSTeams.Add('Accept','application/json, text/plain, */*')
                $headersMSTeams.Add('Accept-Encoding','gzip, deflate, br')

                $paramsTeamsRequest = @{
                    Body = $($EmailAddress | ConvertTo-Json -Depth 2)
                    ContentType = 'application/json'
                    Headers = $headersMSTeams
                    Method = 'POST'
                    TimeoutSec = $Timeout
                    Uri = $teamsAuthZ.regionGtms.middleTier + '/beta/users/searchV2?includeDLs=true&includeBots=true&enableGuest=true&source=addToTeam&skypeTeamsInfo=true'
                }

                $searchRequest = Invoke-RestMethod @paramsTeamsRequest

                $collection += $searchRequest.value
            }
            catch {
                $_
            }
        }
        else
        {
            Write-Verbose 'Could not detect endpoint!...'
            break
        }
    }

    end
    {
        $collection

        $timer.Stop()

        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:ConvertFrom-EncodedCommand
{
    <#
    .SYNOPSIS
        This function is intended to decode a string.
    .DESCRIPTION
        This function is intended to decode a string, which was used as value for the PowerShell parameter EncodedCommand.
    .PARAMETER EncodedString
        The parameter EncodedString takes the string to be decoded.
    .EXAMPLE
        ConvertFrom-EncodedCommand -EncodedString JABzACAAPQAgACc...QAgACQAZgAgAC0A
    .NOTES

    .LINK
        https://ingogegenwarth.wordpress.com/
    #>
    [CmdletBinding()]
    param(
        [parameter(
            mandatory = $true,
            Position = 0)]
        [System.String]
        $EncodedString
    )
    
    try {
        [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($EncodedString))
    }
    catch {
        $_
    }
}

function global:Format-CalDiag
{
    Param (
        [Parameter( Mandatory=$true, Position=0)]
        [System.Object[]]
        $CalendarDiagnosticObjects
    )

    $CalendarDiagnosticObjects | select OriginalLastModifiedTime,LastModifiedTime,OriginalCreationTime,CreationTime,CalendarLogTriggerAction,OriginalClientInfoString,ClientInfoString,MeetingRequestType,ItemClass,ItemVersion,ParentDisplayName,OriginalParentDisplayName,SenderEmailAddress,ResponsibleUserName,ClientIntent,SubjectProperty,NormalizedSubject,DisplayAttendeesAll,Location,ReceivedBy,ReceivedRepresenting,MapiPRStartDate,MapiPREndDate,ViewStartTime,ViewEndTime,CleanGlobalObjectId,Preview,ChangeHighlight,AppointmentState

}

function global:Get-MSOLUserError
{
    Param (
        [Parameter( Mandatory=$true, Position=0)]
        [System.String]
        $UserSearchString
    )

    (Get-MsolUser -SearchString $UserSearchString ).errors.errordetail.ObjectErrors.ErrorRecord.ErrorDescription
}

function global:Get-LAPFIfromStoreID
{
    <#

        .SYNOPSIS

        Created by: https://ingogegenwarth.wordpress.com/
        Version:    42 ("What do you get if you multiply six by nine?")
        Changed:    29.09.2021

        .LINK
        http://gsexdev.blogspot.com/

        .DESCRIPTION

        The purpose of the script is to convert a folder id in StoreID format to LastParentFolderID in HexEntryID format

        .PARAMETER SourceID

        The FolderID in StoreID format to be converted.

        .PARAMETER DestinationID

        Destination format for conversion.

        .PARAMETER WebServicesDLL

        Path to EWS managed API DLL.

        .PARAMETER AccessToken

        AccessToken to be used for access Exchange Online.

        .EXAMPLE

        # convert a folder ID and return the LAPFID
        Get-LAPFIfromStoreID -EmailAddress LynneR@M365x272199.OnMicrosoft.com -SourceId LgAAAABtXj...p3mzwuKAAUhS+Y5AAAB -AccessToken eyJ0eXAiOiJKV1QiLC...0Pf-4hM1Tu7WLw
    #>
    [CmdletBinding()]
    Param (
        [Parameter(
            Mandatory=$true,
            Position=0)]
        [System.String]
        $SourceID,

        [Parameter(
            Mandatory=$false,
            Position=1)]
        [ValidateSet('EwsLegacyId','EwsId','EntryId','HexEntryId','StoreId','OwaId')]
        [System.String]
        $DestinationID = 'HexEntryId',

        [Parameter(
            Mandatory=$false,
            Position=2)]
        [ValidateScript({If (Test-Path $_ -PathType leaf){$True} Else {Throw "WebServices DLL could not be found!"}})]
        [System.String]
        $WebServicesDLL,

        [Parameter(
            Mandatory=$false,
            Position=3)]
        $AccessToken
    )

    if ( [System.String]::IsNullOrEmpty($AccessToken) )
    {
        # thanks to https://gsexdev.blogspot.com/
        function Show-OAuthWindow
        {
            [CmdletBinding()]
            param (
                [System.Uri]
                $Url

            )
            ## Start Code Attribution
            ## Show-AuthWindow function is the work of the following Authors and should remain with the function if copied into other scripts
            ## https://foxdeploy.com/2015/11/02/using-powershell-and-oauth/
            ## https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/
            ## End Code Attribution
            Add-Type -AssemblyName System.Web
            Add-Type -AssemblyName System.Windows.Forms

            $form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width = 440; Height = 640 }
            $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width = 420; Height = 600; Url = ($url) }
            $Navigated = {
                if ( ($web.DocumentText -match "document.location.replace") -or ($web.Url.AbsoluteUri -match "code=[^&]*") ) {
                    $Script:oAuthCode = [regex]::match($web.DocumentText, "code=(.*?)\\u0026").Groups[1].Value
                    if ([System.String]::IsNullOrEmpty($Script:oAuthCode))
                    {
                        if ($web.Url.AbsoluteUri -match "error=[^&]*")
                        {
                            $Script:oAuthCode = [System.Web.HttpUtility]::UrlDecode($web.Url.AbsoluteUri)
                        }
                        else
                        {
                            $Script:oAuthCode = [System.Web.HttpUtility]::ParseQueryString($web.Url.AbsoluteUri)[0]
                        }
                    }
                    $form.Close();
                }
            }
            $web.ScriptErrorsSuppressed = $true
            $web.Add_Navigated($Navigated)
            $form.Controls.Add($web)
            $form.Add_Shown( { $form.Activate() })
            $form.ShowDialog() | Out-Null
            return $Script:oAuthCode
        }
        
    }

    try {

            Add-Type -AssemblyName System.Web
            $ClientID = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
            $audience = 'https://outlook.office365.com'
            # parameters for auth code flow
            $RedirectURI = 'urn:ietf:wg:oauth:2.0:oob'
            $state = Get-Random
            $authURI = "https://login.microsoftonline.com/Common"
            $authURI += "/oauth2/authorize?client_id=$ClientId"
            $authURI += "&response_type=code&redirect_uri= " + [System.Web.HttpUtility]::UrlEncode($RedirectURI)
            $authURI += "&response_mode=query&resource=" + [System.Web.HttpUtility]::UrlEncode($audience) + "&state=$state"
            $authURI += "&prompt=select_account"

            # acquire auth code from AAD
            $authCode = Show-OAuthWindow -Url $authURI

            $body = @{
                "grant_type" = "authorization_code";
                "client_id" = "$ClientId";
                "code" = $authCode;
                "redirect_uri" = $RedirectURI
            }

            $paramsTokenRequest = @{
                Body = $body
                ContentType = 'application/x-www-form-urlencoded'
                Method = 'Post'
                #TimeoutSec = $Timeout
                Uri = "https://login.microsoftonline.com/common/oauth2/token"
            }

            # acquire access token
            $tokenRequest = Invoke-RestMethod @paramsTokenRequest
            $AccessToken = $tokenRequest.access_token
    }
    catch {
        $_
        break
    }

    if ($WebServicesDLL)
    {
                try {
                    $EWSDLL = $WebServicesDLL
                    Import-Module -Name $EWSDLL
                }
                catch {
                    $Error[0].Exception
                    exit
                }
    }
    else
    {
        ## Load Managed API dll
        ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
        $EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
        if (Test-Path -Path $EWSDLL)
        {
            Import-Module -Name $EWSDLL
        }
        else
        {
            "$(get-date -format yyyyMMddHHmmss):"
            "This script requires the EWS Managed API 1.2 or later."
            "Please download and install the current version of the EWS Managed API from"
            "http://go.microsoft.com/fwlink/?LinkId=255472"
            ""
            "Exiting Script."
            break
        }
    }

    ## Set Exchange Version
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    ## Create Exchange Service Object
    $Service = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList ($ExchangeVersion)
    #$service.PreAuthenticate = $true
    #set DateTimePrecision to get milliseconds
    $Service.DateTimePrecision=[Microsoft.Exchange.WebServices.Data.DateTimePrecision]::Milliseconds
    #$service.TraceEnabled = $true
    $service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$AccessToken
    $service.Url = 'https://outlook.office365.com/EWS/exchange.asmx'

    $AltIdBase = New-Object -TypeName Microsoft.Exchange.WebServices.Data.AlternateId -ArgumentList ('StoreID',$SourceID,'foobar@bla.com')
    $Converted = $Service.ConvertId($AltIdBase,$DestinationID)
    $Converted.UniqueId.Substring(44,44)
}

function global:New-AppAccessPolicy
{
<#

        .SYNOPSIS

        Created by: https://ingogegenwarth.wordpress.com/
        Version:    42 ("What do you get if you multiply six by nine?")
        Changed:    13.08.2021

        .LINK
        https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access

        .DESCRIPTION

        The purpose of the function is to create an application access policy with short description

        .PARAMETER Recipient

        The Recipient parameter specifies the recipient to define in the policy.

        .PARAMETER ApplicationID

        The Identity parameter specifies the GUID of the apps to include in the policy.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(
            Mandatory=$true,
            Position=0)]
        [System.String]
        $Recipient,

        [Parameter(
            Mandatory=$true,
            Position=1)]
        [System.String]
        $ApplicationID
    )
    
    try {
        New-ApplicationAccessPolicy -Description "Restrict access to $($Recipient)" -AppId $ApplicationID -PolicyScopeGroupId $Recipient -AccessRight RestrictAccess
    }
    catch{
        $_
    }
}

function global:Format-FolderStats
{
    Param (
        [Parameter( Mandatory=$true, Position=0)]
        [System.Object[]]
        $FolderStatistics
    )

    [System.Collections.ArrayList]$script:selectProperties = @(
    "Date",
    "CreationTime",
    "LastModifiedTime",
    "Name",
    "FolderPath",
    "FolderType",
    "TargetQuota",
    "StorageQuota",
    "StorageWarningQuota",
    "VisibleItemsInFolder",
    "HiddenItemsInFolder",
    "ItemsInFolder",
    "DeletedItemsInFolder",
    #"FolderSize",
    '@{l="FolderSizeMB";e={ConvertFrom-SizeString -SizeString $_.FolderSize -ToMB}}',
    "DeletePolicy",
    "ArchivePolicy",
    "CompliancePolicy",
    "RetentionFlags",
    "ItemsInFolderAndSubfolders",
    #"FolderAndSubfolderSize",
    '@{l="FolderAndSubfolderSizeMB";e={ConvertFrom-SizeString -SizeString $_.FolderAndSubfolderSize -ToMB}}',
    "TopSubject",
    #"TopSubjectSize",
    '@{l="TopSubjectSizeMB";e={ConvertFrom-SizeString -SizeString $_.TopSubjectSize -ToMB}}',
    "TopSubjectCount",
    "TopSubjectClass",
    "TopSubjectPath",
    "TopSubjectReceivedTime",
    "TopSubjectFrom",
    "TopClientInfoForSubject",
    "TopClientInfoCountForSubject",
    "OldestItemReceivedDate",
    "NewestItemReceivedDate",
    "OldestDeletedItemReceivedDate",
    "NewestDeletedItemReceivedDate",
    "OldestItemLastModifiedDate",
    "NewestItemLastModifiedDate",
    "OldestDeletedItemLastModifiedDate",
    "NewestDeletedItemLastModifiedDate",
    "Identity",
    "SearchFolder",
    "Diagnostics",
    "DiagnosticInfo"
    )

    $cmd = '$FolderStatistics |' + "select $($selectProperties -join ',')"
    Invoke-Expression $cmd

}

function global:ConvertFrom-ImmutableId
{
    <#

        .SYNOPSIS

        Created by: https://ingogegenwarth.wordpress.com/
        Version:    42 ("What do you get if you multiply six by nine?")
        Changed:    01.10.2021

        .LINK
        https://ingogegenwarth.wordpress.com/

        .DESCRIPTION

        The purpose of the script is to convert a ImmutableId into a Guid

        .PARAMETER ImmutableId

        The ImmutableId to be converted.

        .EXAMPLE

        ConvertFrom-ImmutableId -ImmutableId "B/YXGYLqqU2xrbJKWyNDrB=="

    #>
    [CmdletBinding()]
    param(
        [System.String]
        $ImmutableId
    )

    try
    {
        ([System.Guid]([Convert]::FromBase64String($ImmutableId))).Guid
    }
    catch
    {
        $_
    }
}

function global:ConvertTo-ImmutableId
{
    <#

        .SYNOPSIS

        Created by: https://ingogegenwarth.wordpress.com/
        Version:    42 ("What do you get if you multiply six by nine?")
        Changed:    01.10.2021

        .LINK
        https://ingogegenwarth.wordpress.com/

        .DESCRIPTION

        The purpose of the script is to convert a Guid into a ImmutableId

        .PARAMETER Guid

        The Guid to be converted.

        .EXAMPLE

        ConvertTo-ImmutableId -Guid 127836d0-df84-4190-a6df-541df44ca0d9

    #>
    [CmdletBinding()]
    param(
        [System.Guid]
        $Guid
    )

    try
    {
        [System.Convert]::ToBase64String([System.Guid]::New($Guid).ToByteArray())
    }
    catch
    {
        $_
    }
}

function global:Get-SMTPDNSEntries
{
    [CmdletBinding()]
    param(
        [System.String]
        $Domain,

        [System.String]
        $Server = '8.8.8.8',

        [System.String]
        $Selector = 'selector1'

    )

    begin
    {
        # initiate objects
        $output = New-Object -TypeName psobject

        $paramsDNS = @{
            ErrorAction = 'SilentlyContinue'
        }
        if($Server) {
            $paramsDNS.Add('Server',$Server)
            $nameServer = $Server
        }

        # retrieve NS
        $nsDNS = [System.Collections.ArrayList]@()
        $nsDNS += Resolve-DnsName -Name $Domain -Type NS @paramsDNS
        if (-not [System.String]::IsNullOrWhiteSpace($nsDNS.NameHost))
        {
            $paramsDNS.Server = $nsDNS[0].NameHost
        }
        else
        {
            Write-Warning 'Could not find NS!'
        }

    }
    process
    {
        try
        {
            # get MX records
            $mxRecords = Resolve-DnsName -DnsOnly -Type MX -Name $Domain @paramsDNS
            $mx= $mxRecords | Where-Object 'Type' -eq 'MX' | Select-Object NameExchange,Preference,TTL,@{l='IPAddress';e={(Resolve-DnsName -Type A_AAAA -Name $_.NameExchange | Select-Object -ExpandProperty IPAddress) -join ','}}
            Write-Verbose "Done with MX records..."

            # get DMARC record
            $dmarcRecord = Resolve-DnsName -DnsOnly -Type TXT -Name "_dmarc.$Domain" @paramsDNS
            $dmarc = $dmarcRecord | Where-Object Type -eq 'TXT' | select @{l='Record';e={$_.Name +'|TTL='+ $_.TTL + '|' + $_.Strings }}
            Write-Verbose "Done with DMARC record..."

            # get SPF record
            $spfRecord = Resolve-DnsName -DnsOnly -Type TXT -Name "_spf.$Domain" @paramsDNS
            if (-not [System.String]::IsNullOrWhiteSpace($spfRecord))
            {
                $spf = $spfRecord  | Where-Object Type -eq 'TXT' | select @{l='Record';e={$_.Name +'|TTL='+ $_.TTL + '|' + $_.Strings }}
            }
            else
            {
                $spf = Resolve-DnsName -Name $Domain -Type TXT @paramsDNS | Where-Object Strings -match "v=spf" | select @{l='Record';e={$_.Name +'|TTL='+ $_.TTL + '|' + $_.Strings }}
            }
            Write-Verbose "Done with SPF record..."

            # get DKIM records
            $dkimResponse = Resolve-DnsName -Name "$selector._domainkey.$Domain" -Type txt @paramsDNS
            # check response

            if ($dkimResponse.Type -eq 'CNAME')
            {
                # retrieve NS
                if ($nameServer)
                {
                    $paramsDNS.Server = $nameServer
                }
                else
                {
                    $paramsDNS.Remove('Server')
                }

                if (-not $DoNotTrySOA)
                {
                    $nsDNSDKIM = [System.Collections.ArrayList]@()
                    $nsDNSDKIM = Resolve-DnsName -Name $dkimResponse.NameHost -Type NS @paramsDNS
                    if ( (-not [System.String]::IsNullOrWhiteSpace($nsDNSDKIM.NameHost)) -or (-not [System.String]::IsNullOrWhiteSpace($nsDNSDKIM.PrimaryServer)) )
                    {
                        if ($nsDNSDKIM[0].PrimaryServer)
                        {
                            $paramsDNS.Server = $nsDNSDKIM[0].PrimaryServer
                        }
                        else
                        {
                            $paramsDNS.Server = $nsDNSDKIM[0].NameHost
                        }
                    }
                    else
                    {
                        Write-Warning 'Could not find SOA or NS!'
                    }
                }

                $dkim = Resolve-DnsName -Name $dkimResponse.NameHost -Type TXT @paramsDNS | select @{l='Record';e={$_.Name +'|TTL='+ $_.TTL + '|' + $_.Strings }}
            }
            else
            {
                $dkim = $dkimResonse | select @{l='Record';e={$_.Name +'|TTL='+ $_.TTL + '|' + $_.Strings }}
            }

            $output | Add-Member -MemberType NoteProperty -Name Domain -Value $Domain
            $output | Add-Member -MemberType NoteProperty -Name MXRecords -Value $mx
            $output | Add-Member -MemberType NoteProperty -Name DMARCRecord -Value $dmarc.Record
            $output | Add-Member -MemberType NoteProperty -Name SPFRecord -Value $spf.Record
            $output | Add-Member -MemberType NoteProperty -Name DKIMRecord -Value $dkim.Record
        }
        catch
        {
            $_
        }
    }

    end
    {
        $output
    }

}

function global:Get-JunkConfiguration
{
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [System.String[]]
        $Identity,

        [System.Management.Automation.SwitchParameter]
        $ShowProgress,

        [System.Management.Automation.SwitchParameter]
        $AllProperties

    )

    begin
    {
        # initiate variables
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $collection = [System.Collections.ArrayList]@()
        [System.Int32]$j = '1'
    }

    process
    {
        foreach ($Id in $Identity)
        {
            if ($ShowProgress)
            {
                $progressParams = @{
                    Activity = "Processing user - $($Id)"
                    PercentComplete = $j / $Identity.count * 100
                    Status = "Remaining: $($Identity.count - $j) out of $($Identity.count)"
                }
                Write-Progress @progressParams
            }

            $j++

            $setting = New-Object -TypeName psobject
            $setting | Add-Member -MemberType NoteProperty -Name Identity -Value $Id

            try
            {
                $junkFolder = Get-MailboxJunkEmailConfiguration -Identity $Id -ErrorAction Stop

                $setting | Add-Member -MemberType NoteProperty -Name Status -Value $junkFolder.Enabled
                
                if ($AllProperties)
                {
                    $setting | Add-Member -MemberType NoteProperty -Name TrustedListsOnly -Value $junkFolder.TrustedListsOnly
                    $setting | Add-Member -MemberType NoteProperty -Name ContactsTrusted -Value $junkFolder.ContactsTrusted
                    $setting | Add-Member -MemberType NoteProperty -Name TrustedSendersAndDomains -Value $junkFolder.TrustedSendersAndDomains
                    $setting | Add-Member -MemberType NoteProperty -Name BlockedSendersAndDomains -Value $junkFolder.BlockedSendersAndDomains
                    $setting | Add-Member -MemberType NoteProperty -Name TrustedRecipientsAndDomains -Value $junkFolder.TrustedRecipientsAndDomains
                }
            }
            catch
            {
                $setting | Add-Member -MemberType NoteProperty -Name Status -Value $_
            }

            $collection += $setting

        }

    }

    end
    {
        if ($ShowProgress)
        {
            $progressParams.Status = "Ready"
            $progressParams.Add('Completed',$true)
            Write-Progress @progressParams
        }
        $collection
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Get-AadProvisioningErrors
{
    [CmdLetBinding()]
    param (
        [System.String]
        $User
    )

    if (-not [System.String]::IsNullOrEmpty($User))
    {
        try {
            (Get-AzureADUser -SearchString $User).ProvisioningErrors | Where-Object -FilterScript {-not [System.String]::IsNullOrEmpty($_.ErrorDetail)} | select Timestamp,ErrorDetail
        }
        catch {
            $_
        }
    }
}

function global:Get-AadLastDirsyncTime
{
    (Get-AzureADTenantDetail).CompanyLastDirSyncTime
}

function global:ConvertFrom-SizeString
{
param(
    [Parameter(
        Mandatory=$true,
        ValueFromPipelineByPropertyName=$false,
        Position=0)]
    [System.String]
    $SizeString,

    [Parameter(
        Mandatory=$false,
        ValueFromPipelineByPropertyName=$false,
        Position=1)]
    [System.Management.Automation.SwitchParameter]
    $ToMB,

    [Parameter(
        Mandatory=$false,
        ValueFromPipelineByPropertyName=$false,
        Position=2)]
    [System.Management.Automation.SwitchParameter]
    $ToGB
)
    $result = $SizeString -replace '(.*\()|,| [a-z]*\)', ''

    if ($ToMB)
    {
        $result = [System.Math]::Round($result/1MB,2)
    }
    elseIf ($ToGB)
    {
        $result = [System.Math]::Round($result/1GB,2)
    }

    return $result
}

function global:Update-BlockSenderList
{
    [CmdLetBinding()]
    param (
        [System.String]
        $CSV = 'C:\temp\BlockList.csv'
    )

    try
    {
        $senderList = Import-Csv -Encoding utf8 -Path $CSV
        $SenderList01 = $senderList[0..199]
        $SenderList02 = $senderList[200..380]
        $SenderList03 = $senderList[381..599]

        $commandRule01 = 'Set-TransportRule -Identity 11e63de1-9b71-41aa-8ddc-4fc61ef2610e -FromAddressMatchesPatterns "' + $($senderList01.SMTPAddress -join '","') + '"'
        Invoke-Expression $commandRule01

        if (-not [System.String]::IsNullOrEmpty($SenderList02))
        {
            Write-Verbose 'Processing rule #2...'
            $commandRule02 = 'Set-TransportRule -Identity 070bdc77-d095-40df-b8cc-0c6213c13802 -FromAddressMatchesPatterns "' + $($senderList02.SMTPAddress -join '","') + '"'
            Invoke-Expression $commandRule02
        }
        if (-not [System.String]::IsNullOrEmpty($SenderList03))
        {
            Write-Verbose 'Processing rule #3...'
            $commandRule03 = 'Set-TransportRule -Identity e5a656f4-7ec5-4064-84cb-be425214f24b -FromAddressMatchesPatterns "' + $($senderList03.SMTPAddress -join '","') + '"'
            Invoke-Expression $commandRule03
        }
    }
    catch
    {
        $_
    }
}

function global:Add-MailboxFolderPermissionRecursive
{
    [CmdletBinding()]
    param(
        [System.String]
        $Identity,

        [System.String]
        $User,

        [System.String]
        [ValidateSet("Author","Contributor","Editor","NonEditingAuthor","Owner","PublishingAuthor","PublishingEditor","Reviewer")]
        $AccessRights,

        [System.String]
        $FilterFolderPath
    )

    # get recipient
    $trustee = Get-Recipient -Identity $User -ErrorAction SilentlyContinue
    if ([System.String]::IsNullOrEmpty($trustee))
    {
        Write-Error "User $($User) couldn't be found!"
        break
    }

    # retrieve folders with scope Inbox
    $folderSet = Get-MailboxFolderStatistics -Identity $Identity -FolderScope Inbox

    if ($FilterFolderPath)
    {
        $folderSet = $folderSet | Where-Object {$_.FolderPath -Match $FilterFolderPath}
        Write-Verbose "Found the following folders for filter $($FilterFolderPath):"
        $folderSet.FolderPath
    }

    if (-not [System.String]::IsNullOrEmpty($folderSet) )
    {
        Write-Verbose "Found $(($folderSet | measure).Count) folders..."
        foreach ($folder in $folderSet)
        {
            $progressParams = @{
                Activity = "Processing folder - $($folder.Name)"
                PercentComplete = $j / ($folderSet | measure).Count * 100
                Status = "Remaining: $(($folderSet | measure).Count - $j) out of $(($folderSet | measure).Count)"
            }
            Write-Progress @progressParams
            $j++

            Write-Verbose "Processing folder:$($folder.Name)..."
            
            $params = @{
                Identity = $Identity + ":" + $folder.FolderId
                User = $trustee.Identity
                AccessRights = $AccessRights
                ErrorAction = 'Stop'
            }
            try
            {
                Add-MailboxFolderPermission @params
            }
            catch
            {
                if ('UserAlreadyExistsInPermissionEntryException' -eq $_.CategoryInfo.Reason)
                {
                    Write-Verbose "Existing permission found. Will replace..."
                    Set-MailboxFolderPermission @params
                }
                else
                {
                    Write-Host "Error setting permission on folder $($folder):$_"
                }
            }
        }

        $progressParams.Status = "Ready"
        Write-Progress @progressParams
    }
}

function global:Remove-MailboxFolderPermissionRecursive
{
    [CmdletBinding()]
    param(
        [System.String]
        $Identity,

        [System.String]
        $User,

        [System.String]
        $FilterFolderPath
    )

    # get recipient
    $trustee = Get-Recipient -Identity $User -ErrorAction SilentlyContinue
    if ([System.String]::IsNullOrEmpty($trustee))
    {
        Write-Error "User $($User) couldn't be found!"
        break
    }

    # retrieve folders with scope Inbox
    $folderSet = Get-MailboxFolderStatistics -Identity $Identity -FolderScope Inbox

    if ($FilterFolderPath)
    {
        $folderSet = $folderSet | Where-Object {$_.FolderPath -Match $FilterFolderPath}
        Write-Verbose "Found the following folders for filter $($FilterFolderPath):"
        $folderSet.FolderPath
    }

    if (-not [System.String]::IsNullOrEmpty($folderSet) )
    {
        Write-Verbose "Found $(($folderSet | measure).Count) folders..."
        foreach ($folder in $folderSet)
        {
            $progressParams = @{
                Activity = "Processing folder - $($folder.Name)"
                PercentComplete = $j / ($folderSet | measure).Count * 100
                Status = "Remaining: $(($folderSet | measure).Count - $j) out of $(($folderSet | measure).Count)"
            }
            Write-Progress @progressParams
            $j++

            Write-Verbose "Processing folder:$($folder.Name)..."
            $params = @{
                Identity = $Identity + ":" + $folder.FolderId
                User = $trustee.Identity
                ErrorAction = 'Stop'
                Confirm = $false
            }

            try
            {
                Remove-MailboxFolderPermission @params
            }
            catch
            {
                if ('UserNotFoundInPermissionEntryException' -eq $_.CategoryInfo.Reason)
                {
                     Write-Verbose 'No existing permission for this user to remove...'
                }
                else
                {
                    Write-Host "Error removing permission on folder $($folder.Name):$_"
                }
            }
        }

        $progressParams.Status = "Ready"
        Write-Progress @progressParams
    }
}

function global:Get-AADServicePrincipalEXOReport
{
    [CmdletBinding()]
    param(
        [System.Management.Automation.SwitchParameter]
        $IncludePolicyCheck
    )
    begin
    {
        # start timer
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        # initiate collection
        $collection = [System.Collections.ArrayList]@()
        # retrieve Microsoft Graph and Exchange Online servicePrincipals
        $MSGraphEXOSPN = Get-MgServicePrincipal -Filter "(AppId eq '00000002-0000-0ff1-ce00-000000000000') or (AppId eq '00000003-0000-0000-c000-000000000000')"
        # extract AppRoles
        $AppRoles = $MSGraphEXOSPN.AppRoles | Where-Object { $_.Value -match 'Mail\.|MailboxSettings\.|Calendars\.|Contacts\.|full_access_as_app'} | Sort-Object -Property Value -Unique
        # retrieve all servicePrincipals
        $AllSPN = Get-MgServicePrincipal -ExpandProperty AppRoleAssignments -All
        Write-Verbose "Found $($AllSPN.Count) serviceprincipals..."
        # filter for serviceprincipals with AppRoleAssignments
        $SPNwithAppRoles = $AllSPN | Where-Object { -not [System.String]::IsNullOrEmpty($_.AppRoleAssignments)}
        Write-Verbose "Found $($SPNwithAppRoles.Count) serviceprincipals with AppRoleAssignments..."
        Write-Verbose "Retrieved all ServicePrincipals processing time:$($timer.Elapsed.ToString())"
        
        if ($IncludePolicyCheck)
        {
            # retreieve all application access policies from EXO
            $AppPolicies = Get-ApplicationAccessPolicy -ErrorAction SilentlyContinue
        }

        function Test-ServicePrincipal
        {
            [OutputType([System.Boolean])]
            param(
                [Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal]
                $ServicePrincipal,

                [Microsoft.Graph.PowerShell.Models.MicrosoftGraphAppRole[]]
                $AppRoles
            )

            foreach ($AppRoleID in $ServicePrincipal.AppRoleAssignments.AppRoleId)
            {
                if ($AppRoles.Id.Contains($AppRoleID))
                {
                    return $true
                    break
                }
                else
                {
                    return $false
                }
            }
        }

    }

    process
    {
        foreach ($SPN in $SPNwithAppRoles)
        {
            # create object for current serviceprincipal
            $sPNObject = New-Object -TypeName psobject
            # check for relevant permissions
            if (Test-ServicePrincipal -ServicePrincipal $SPN -AppRoles $AppRoles )
            {
                $sPNObject | Add-Member -MemberType NoteProperty -Name AppDisplayName -Value $SPN.AppDisplayName
                $sPNObject | Add-Member -MemberType NoteProperty -Name ObjectId -Value $SPN.Id
                $sPNObject | Add-Member -MemberType NoteProperty -Name AppId -Value $SPN.AppId
                $sPNObject | Add-Member -MemberType NoteProperty -Name AppOwnerOrganizationId -Value $SPN.AppOwnerOrganizationId

                if ($IncludePolicyCheck)
                {
                    if ([System.String]::IsNullOrEmpty($AppPolicies))
                    {
                        $sPNObject | Add-Member -MemberType NoteProperty -Name EXOPolicyExists -Value 'None found!'
                    }
                    else
                    {
                        $sPNObject | Add-Member -MemberType NoteProperty -Name EXOPolicyExists -Value $($AppPolicies.AppId.Contains($SPN.AppId))
                    }
                    # initiate variable
                    [System.String]$Permissions = ''
                    # build permission string from AppRoleId
                    $Permissions = ($SPN.AppRoleAssignments.AppRoleId | ForEach-Object{$perm=$_ ; $AppRoles | Where-Object {$_.Id -eq $perm}}).Value -join '|'
                    $sPNObject | Add-Member -MemberType NoteProperty -Name Permissions -Value $Permissions
                }

                # add to collection for output
                $collection += $sPNObject
            }
        }
    }

    end
    {
        if ([System.String]::IsNullOrEmpty($collection))
        {
            Write-Host 'No enterprise application with EXO related permissions found!'
        }
        else
        {
            $collection
        }
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Format-AuditLog
{
    Param (
        [Parameter( Mandatory=$true, Position=0)]
        [System.Object[]]
        $AuditLog
    )

    [System.Collections.ArrayList]$script:selectProperties = @(
    "LastAccessed",
    "Operation",
    "OperationResult",
    "LogonType",
    "LogonUserDisplayName",
    "DestFolderPathName",
    "FolderPathName",
    "FolderName",
    "ClientInfoString",
    "ClientIPAddress",
    "ClientIP",
    "ClientProcessName",
    "ClientVersion",
    "InternalLogonType",
    "MailboxOwnerUPN",
    "ItemInternetMessageId",
    "ItemSubject",
    "SourceItemSubjectsList",
    "MailboxGuid",
    "MailboxResolvedOwnerName",
    "AuditOperationsCountInAggregatedRecord",
    "DestMailboxOwnerUPN",
    "DestMailboxGuid",
    "CrossMailboxOperation",
    "SourceItems",
    "SourceFolders",
    "SourceItemIdsList",
    "SourceItemAttachmentsList",
    "SourceItemFolderPathNamesList",
    "SourceFolderPathNamesList",
    "SourceItemInternetMessageIdsList",
    "ItemId",
    "ItemAttachments",
    "DirtyProperties",
    "OriginatingServer",
    "SessionId",
    "OperationProperties",
    "AggregatedRecordFoldersData",
    "AppId",
    "ClientAppId",
    "ItemIsRecord",
    "ItemComplianceLabel"
    #"Identity"
    )

    $cmd = '$AuditLog |' + "select $($selectProperties -join ',')"
    Invoke-Expression $cmd

}

function global:Set-Oauth2PermissionGrantforMG
{
    <#
        .SYNOPSIS
            This function creates Oauth2PermissionGrant for a user for Microsoft Graph permissions.
        .DESCRIPTION
            This function creates Oauth2PermissionGrant for a user for Microsoft Graph permissions in case a user cannot grant by himself e.g.: missing publisher verification. This is more specific than granting admin consent on enterprise application level for any permission.
        .PARAMETER ObjectID
            The parameter ObjectID defines AAD object id of the service principal in your tenant.
        .PARAMETER Scopes
            This parameter Scopes defines comma seperated scopes you want to grant.
        .PARAMETER RemoveAllGrants
            The parameter RemoveAllGrants removes all consent for give user.
        .EXAMPLE
            # create Oauth2PermissionGrant for John Doe to read and write Intune configuration on the service principal 919c3642-68e5-43a1-b427-38e44fce5d4f
            Set-Oauth2PermissionGrantforMG -User john.doe@contoso.com -Verbose -ObjectID 919c3642-68e5-43a1-b427-38e44fce5d4f -Scopes DeviceManagementServiceConfig.ReadWrite.All
            # removes all Oauth2PermissionGrant for the user John Doe on the service principal 919c3642-68e5-43a1-b427-38e44fce5d4f
            Set-Oauth2PermissionGrantforMG -User john.doe@contoso.com -Verbose -ObjectID 919c3642-68e5-43a1-b427-38e44fce5d4f -RemoveAllGrants 
        .NOTES
            If you want to leverage all functionality you will need to provide an access token with the following claims:
                User.ReadBasic.All
                Application.ReadWrite.All
                DelegatedPermissionGrant.ReadWrite.All
            Connect-MgGraph -Scopes User.ReadBasic.All,Application.ReadWrite.All,DelegatedPermissionGrant.ReadWrite.All
        .LINK
            https://docs.microsoft.com/en-us/azure/active-directory/manage-apps/grant-consent-single-user
    #>
    [CmdletBinding()]
    param(
        [parameter(
            Mandatory=$true,
            Position=0
        )]
        [System.String]
        $ObjectID,

        [parameter(
            Mandatory=$true,
            Position=1
        )]
        [System.String[]]
        $User,

        [parameter(
            Mandatory=$false,
            Position=2,
            ParameterSetName='scope'
        )]
        [System.String[]]
        [ValidateSet('AccessReview.Read.All','AccessReview.ReadWrite.All','AccessReview.ReadWrite.Membership','AdministrativeUnit.Read.All',
        'AdministrativeUnit.ReadWrite.All','Agreement.Read.All','Agreement.ReadWrite.All','AgreementAcceptance.Read','AgreementAcceptance.Read.All',
        'Analytics.Read','APIConnectors.Read.All','APIConnectors.ReadWrite.All','AppCatalog.Read.All','AppCatalog.ReadWrite.All','AppCatalog.Submit',
        'Application.Read.All','Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All','Approval.Read.All','Approval.ReadWrite.All',
        'AttackSimulation.Read.All','AuditLog.Read.All','AuthenticationContext.Read.All','AuthenticationContext.ReadWrite.All','BitlockerKey.Read.All',
        'BitlockerKey.ReadBasic.All','Bookings.Manage.All','Bookings.Read.All','Bookings.ReadWrite.All','BookingsAppointment.ReadWrite.All','Calendars.Read',
        'Calendars.Read.Shared','Calendars.ReadWrite','Calendars.ReadWrite.Shared','Channel.Create','Channel.Delete.All','Channel.ReadBasic.All',
        'ChannelMember.Read.All','ChannelMember.ReadWrite.All','ChannelMessage.Edit','ChannelMessage.Read.All','ChannelMessage.ReadWrite','ChannelMessage.Send',
        'ChannelSettings.Read.All','ChannelSettings.ReadWrite.All','Chat.Create','Chat.Read','Chat.ReadBasic','Chat.ReadWrite','ChatMember.Read',
        'ChatMember.ReadWrite','ChatMessage.Read','ChatMessage.Send','CloudPC.Read.All','CloudPC.ReadWrite.All','ConsentRequest.Read.All',
        'ConsentRequest.ReadWrite.All','Contacts.Read','Contacts.Read.Shared','Contacts.ReadWrite','Contacts.ReadWrite.Shared','CrossTenantInformation.ReadBasic.All',
        'CrossTenantUserProfileSharing.Read','CrossTenantUserProfileSharing.Read.All','CrossTenantUserProfileSharing.ReadWrite','CrossTenantUserProfileSharing.ReadWrite.All',
        'CustomAuthenticationExtension.Read.All','CustomAuthenticationExtension.ReadWrite.All','CustomSecAttributeAssignment.Read.All',
        'CustomSecAttributeAssignment.ReadWrite.All','CustomSecAttributeDefinition.Read.All','CustomSecAttributeDefinition.ReadWrite.All',
        'DelegatedAdminRelationship.Read.All','DelegatedAdminRelationship.ReadWrite.All','DelegatedPermissionGrant.ReadWrite.All','Device.Command',
        'Device.Read','Device.Read.All','DeviceManagementApps.Read.All','DeviceManagementApps.ReadWrite.All','DeviceManagementConfiguration.Read.All',
        'DeviceManagementConfiguration.ReadWrite.All','DeviceManagementManagedDevices.PrivilegedOperations.All','DeviceManagementManagedDevices.Read.All',
        'DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementRBAC.Read.All','DeviceManagementRBAC.ReadWrite.All','DeviceManagementServiceConfig.Read.All',
        'DeviceManagementServiceConfig.ReadWrite.All','Directory.AccessAsUser.All','Directory.Read.All','Directory.ReadWrite.All','Directory.Write.Restricted',
        'DirectoryRecommendations.Read.All','DirectoryRecommendations.ReadWrite.All','Domain.Read.All','Domain.ReadWrite.All','EAS.AccessAsUser.All',
        'eDiscovery.Read.All','eDiscovery.ReadWrite.All','EduAdministration.Read','EduAdministration.ReadWrite','EduAssignments.Read','EduAssignments.ReadBasic',
        'EduAssignments.ReadWrite','EduAssignments.ReadWriteBasic','EduRoster.Read','EduRoster.ReadBasic','EduRoster.ReadWrite','email','EntitlementManagement.Read.All',
        'EntitlementManagement.ReadWrite.All','EventListener.Read.All','EventListener.ReadWrite.All','EWS.AccessAsUser.All','ExternalConnection.Read.All',
        'ExternalConnection.ReadWrite.All','ExternalConnection.ReadWrite.OwnedBy','ExternalItem.Read.All','ExternalItem.ReadWrite.All','ExternalItem.ReadWrite.OwnedBy',
        'Family.Read','Files.Read','Files.Read.All','Files.Read.Selected','Files.ReadWrite','Files.ReadWrite.All','Files.ReadWrite.AppFolder','Files.ReadWrite.Selected',
        'Financials.ReadWrite.All','Group.Read.All','Group.ReadWrite.All','GroupMember.Read.All','GroupMember.ReadWrite.All','IdentityProvider.Read.All',
        'IdentityProvider.ReadWrite.All','IdentityRiskEvent.Read.All','IdentityRiskEvent.ReadWrite.All','IdentityRiskyServicePrincipal.Read.All',
        'IdentityRiskyServicePrincipal.ReadWrite.All','IdentityRiskyUser.Read.All','IdentityRiskyUser.ReadWrite.All','IdentityUserFlow.Read.All',
        'IdentityUserFlow.ReadWrite.All','IMAP.AccessAsUser.All','InformationProtectionPolicy.Read','LearningContent.Read.All','LearningContent.ReadWrite.All',
        'LearningProvider.Read','LearningProvider.ReadWrite','LicenseAssignment.ReadWrite.All','Mail.Read','Mail.Read.Shared','Mail.ReadBasic','Mail.ReadWrite',
        'Mail.ReadWrite.Shared','Mail.Send','Mail.Send.Shared','MailboxSettings.Read','MailboxSettings.ReadWrite','ManagedTenants.Read.All','ManagedTenants.ReadWrite.All',
        'Member.Read.Hidden','Notes.Create','Notes.Read','Notes.Read.All','Notes.ReadWrite','Notes.ReadWrite.All','Notes.ReadWrite.CreatedByApp',
        'Notifications.ReadWrite.CreatedByApp','offline_access','OnlineMeetingArtifact.Read.All','OnlineMeetingRecording.Read.All','OnlineMeetings.Read',
        'OnlineMeetings.ReadWrite','OnlineMeetingTranscript.Read.All','OnPremisesPublishingProfiles.ReadWrite.All','openid','Organization.Read.All',
        'Organization.ReadWrite.All','OrgContact.Read.All','People.Read','People.Read.All','Place.Read.All','Place.ReadWrite.All','Policy.Read.All',
        'Policy.Read.ConditionalAccess','Policy.Read.PermissionGrant','Policy.ReadWrite.AccessReview','Policy.ReadWrite.ApplicationConfiguration',
        'Policy.ReadWrite.AuthenticationFlows','Policy.ReadWrite.AuthenticationMethod','Policy.ReadWrite.Authorization','Policy.ReadWrite.ConditionalAccess',
        'Policy.ReadWrite.ConsentRequest','Policy.ReadWrite.CrossTenantAccess','Policy.ReadWrite.DeviceConfiguration','Policy.ReadWrite.FeatureRollout',
        'Policy.ReadWrite.MobilityManagement','Policy.ReadWrite.PermissionGrant','Policy.ReadWrite.TrustFramework','POP.AccessAsUser.All','Presence.Read',
        'Presence.Read.All','Presence.ReadWrite','PrintConnector.Read.All','PrintConnector.ReadWrite.All','Printer.Create','Printer.FullControl.All',
        'Printer.Read.All','Printer.ReadWrite.All','PrinterShare.Read.All','PrinterShare.ReadBasic.All','PrinterShare.ReadWrite.All','PrintJob.Create',
        'PrintJob.Read','PrintJob.Read.All','PrintJob.ReadBasic','PrintJob.ReadBasic.All','PrintJob.ReadWrite','PrintJob.ReadWrite.All','PrintJob.ReadWriteBasic',
        'PrintJob.ReadWriteBasic.All','PrintSettings.Read.All','PrintSettings.ReadWrite.All','PrivilegedAccess.Read.AzureAD','PrivilegedAccess.Read.AzureADGroup',
        'PrivilegedAccess.Read.AzureResources','PrivilegedAccess.ReadWrite.AzureAD','PrivilegedAccess.ReadWrite.AzureADGroup','PrivilegedAccess.ReadWrite.AzureResources',
        'profile','ProgramControl.Read.All','ProgramControl.ReadWrite.All','RecordsManagement.Read.All','RecordsManagement.ReadWrite.All','Reports.Read.All',
        'ReportSettings.Read.All','ReportSettings.ReadWrite.All','RoleAssignmentSchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory',
        'RoleEligibilitySchedule.Read.Directory','RoleEligibilitySchedule.ReadWrite.Directory','RoleManagement.Read.All','RoleManagement.Read.CloudPC',
        'RoleManagement.Read.Directory','RoleManagement.ReadWrite.CloudPC','RoleManagement.ReadWrite.Directory','RoleManagementPolicy.Read.Directory',
        'RoleManagementPolicy.ReadWrite.Directory','Schedule.Read.All','Schedule.ReadWrite.All','SearchConfiguration.Read.All','SearchConfiguration.ReadWrite.All',
        'SecurityActions.Read.All','SecurityActions.ReadWrite.All','SecurityAlert.Read.All','SecurityAlert.ReadWrite.All','SecurityEvents.Read.All',
        'SecurityEvents.ReadWrite.All','SecurityIncident.Read.All','SecurityIncident.ReadWrite.All','ServiceHealth.Read.All','ServiceMessage.Read.All',
        'ServiceMessageViewpoint.Write','ServicePrincipalEndpoint.Read.All','ServicePrincipalEndpoint.ReadWrite.All','SharePointTenantSettings.Read.All',
        'SharePointTenantSettings.ReadWrite.All','ShortNotes.Read','ShortNotes.ReadWrite','Sites.FullControl.All','Sites.Manage.All','Sites.Read.All','Sites.ReadWrite.All',
        'SMTP.Send','SubjectRightsRequest.Read.All','SubjectRightsRequest.ReadWrite.All','Subscription.Read.All','Tasks.Read','Tasks.Read.Shared','Tasks.ReadWrite',
        'Tasks.ReadWrite.Shared','Team.Create','Team.ReadBasic.All','TeamMember.Read.All','TeamMember.ReadWrite.All','TeamMember.ReadWriteNonOwnerRole.All',
        'TeamsActivity.Read','TeamsActivity.Send','TeamsAppInstallation.ReadForChat','TeamsAppInstallation.ReadForTeam','TeamsAppInstallation.ReadForUser',
        'TeamsAppInstallation.ReadWriteForChat','TeamsAppInstallation.ReadWriteForTeam','TeamsAppInstallation.ReadWriteForUser','TeamsAppInstallation.ReadWriteSelfForChat',
        'TeamsAppInstallation.ReadWriteSelfForTeam','TeamsAppInstallation.ReadWriteSelfForUser','TeamSettings.Read.All','TeamSettings.ReadWrite.All','TeamsTab.Create',
        'TeamsTab.Read.All','TeamsTab.ReadWrite.All','TeamsTab.ReadWriteForChat','TeamsTab.ReadWriteForTeam','TeamsTab.ReadWriteForUser','TeamsTab.ReadWriteSelfForChat',
        'TeamsTab.ReadWriteSelfForTeam','TeamsTab.ReadWriteSelfForUser','TeamworkDevice.Read.All','TeamworkDevice.ReadWrite.All','TeamworkTag.Read','TeamworkTag.ReadWrite',
        'TermStore.Read.All','TermStore.ReadWrite.All','ThreatAssessment.ReadWrite.All','ThreatHunting.Read.All','ThreatIndicators.Read.All',
        'ThreatIndicators.ReadWrite.OwnedBy','ThreatSubmission.Read','ThreatSubmission.Read.All','ThreatSubmission.ReadWrite','ThreatSubmission.ReadWrite.All',
        'ThreatSubmissionPolicy.ReadWrite.All','TrustFrameworkKeySet.Read.All','TrustFrameworkKeySet.ReadWrite.All','UnifiedGroupMember.Read.AsGuest','User.Export.All',
        'User.Invite.All','User.ManageIdentities.All','User.Read','User.Read.All','User.ReadBasic.All','User.ReadWrite','User.ReadWrite.All',
        'UserActivity.ReadWrite.CreatedByApp','UserAuthenticationMethod.Read','UserAuthenticationMethod.Read.All','UserAuthenticationMethod.ReadWrite',
        'UserAuthenticationMethod.ReadWrite.All','UserNotification.ReadWrite.CreatedByApp','UserTimelineActivity.Write.CreatedByApp','WindowsUpdates.ReadWrite.All',
        'WorkforceIntegration.Read.All','WorkforceIntegration.ReadWrite.All')]
        $Scopes,

        [parameter(
            Mandatory=$false,
            Position=3,
            ParameterSetName='cleanup'
        )]
        [System.Management.Automation.SwitchParameter]
        $RemoveAllGrants
    )

    begin
    {
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        # check for required permissions
        $requiredMGPermissions = @('User.ReadBasic.All','Application.ReadWrite.All','DelegatedPermissionGrant.ReadWrite.All')
        # get current context
        $currentMGContext = Get-MgContext
        if (-not [System.String]::IsNullOrEmpty($currentMGContext))
        {
            foreach ($permission in $requiredMGPermissions)
            {
                if (-not $currentMGContext.Scopes.Contains($permission))
                {
                    Write-Error "Required permission missing:$($permission)"
                    [System.Boolean]$insufficientPerms = $true
                }
            }
            if ($insufficientPerms)
            {
                break
            }
        }
        else
        {
            Write-Warning 'No existing connection. Please connect to MS Graph first! e.g.:Connect-MgGraph -Scopes User.ReadBasic.All,Application.ReadWrite.All,DelegatedPermissionGrant.ReadWrite.All'
            break
        }
        # retrieve serviceprincipal
        try
        {
            $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $ObjectID -ErrorAction Stop
        }
        catch
        {
            $_.Exception
            break
        }
        
        # initiate variable
        $userCollection = [System.Collections.ArrayList]@()
        foreach ($object in $User)
        {
            try
            {
                $userCollection += Get-MgUser -UserId $object -ErrorAction Stop
            }
            catch
            {
                $_.Exception
            }
        }
        # retrieve MS Graph
        $resourceMG = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
        # retrieve existing grants
        $grantsFilter = @(
            "clientId eq '$($servicePrincipal.Id)'"
            "resourceId eq '$($resourceMG.Id)'"
            "consentType eq 'Principal'"
        ) -join " and "
        $oAuth2Grants = Get-MgOauth2PermissionGrant -Filter $grantsFilter -All
    }

    process
    {
        if (-not [System.String]::IsNullOrEmpty($userCollection))
        {
            foreach ($userMG in $userCollection)
            {
                Write-Verbose "Processing user:$($userMG.UserPrincipalName)"

                # checking for existing grants
                $existingGrants = $oAuth2Grants | Where-Object {$_.PrincipalId -eq $userMG.Id}

                if ($RemoveAllGrants)
                {
                    if ([System.String]::IsNullOrEmpty($existingGrants))
                    {
                        Write-Host 'No existing grants found!'
                    }
                    else
                    {
                        $paramsMgOauth2 = @{
                            OAuth2PermissionGrantId = $existingGrants.Id
                            ErrorAction = 'Stop'
                            #Verbose = $VerbosePreference
                        }

                        Remove-MgOauth2PermissionGrant @paramsMgOauth2
                        break
                    }
                }

                if (-not [System.String]::IsNullOrEmpty($existingGrants))
                {
                    # check for matching grants
                    foreach ($scope in $Scopes)
                    {
                        if ($existingGrants.Scope.Contains($scope))
                        {
                            Write-Verbose "Existing scope found:$($scope)"
                        }
                    }
                }

                if (-not [System.String]::IsNullOrEmpty($Scopes))
                {
                    $paramsMgOauth2 = @{
                        ResourceId = $resourceMG.Id
                        Scope = $($Scopes -join " ")
                        ClientId = $servicePrincipal.Id
                        ConsentType = 'Principal'
                        PrincipalId = $userMG.Id
                        #Verbose = $VerbosePreference
                    }

                    if ([System.String]::IsNullOrEmpty($existingGrants))
                    {
                        $newMGGrant = New-MgOauth2PermissionGrant @paramsMgOauth2
                    }
                    else
                    {
                        $paramsMgOauth2.Add('OAuth2PermissionGrantId',$existingGrants.Id)
                        Update-MgOauth2PermissionGrant @paramsMgOauth2
                    }
                }
            }
        }
    }

    end
    {
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Set-AppRoleAssignmentforMG
{
    <#
        .SYNOPSIS
            This function creates AppRoleAssignment on serviceprincipal for Microsoft Graph and Exchange Online permissions.
        .DESCRIPTION
            This function creates AppRoleAssignment on serviceprincipal for Microsoft Graph and Exchange Online permissions, which are supported by ApplicationAccessPolicy.
        .PARAMETER ObjectID
            The parameter ObjectID defines AAD object id of the service principal in your tenant.
        .PARAMETER Roles
            This parameter Roles defines comma seperated roles you want to grant.
        .PARAMETER AddApplicationPermissions
            The parameter AddApplicationPermissions will add the permissins to the application (not serviceprincipal). This is only possible, when the application is in the same tenant and the user has the required permissions.
        .PARAMETER RemoveAppRoleAssignment
            The parameter RemoveAppRoleAssignment will remove existing grants from serviceprincipal defined in the parameter Roles.
        .EXAMPLE
            # add admin consent for permissions Mail.ReadWrite,Mail.Send on serviceprincipal with objectId 4462f870-4605-4ad6-bf6b-ccc61769bee8
            Set-AppRoleAssignmentforMG -ObjectID 4462f870-4605-4ad6-bf6b-ccc61769bee8 -Verbose -Roles Mail.ReadWrite,Mail.Send
            # remove the consent for permissions Calendars.ReadWrite,Contacts.Read on serviceprincipal with objectId 4462f870-4605-4ad6-bf6b-ccc61769bee8
            Set-AppRoleAssignmentforMG -ObjectID 4462f870-4605-4ad6-bf6b-ccc61769bee8 -Verbose -Roles Calendars.ReadWrite,Contacts.Read -RemoveAppRoleAssignment
            # add admin consent for permissions Mail.ReadWrite on serviceprincipal with objectId 4462f870-4605-4ad6-bf6b-ccc61769bee8 and add permission to application
            Set-AppRoleAssignmentforMG -ObjectID 4462f870-4605-4ad6-bf6b-ccc61769bee8 -Verbose -Roles Mail.ReadWrite -AddApplicationPermissions
        .NOTES
            If you want to leverage all functionality you will need to provide an access token with the following claims:
                AppRoleAssignment.ReadWrite.All
                Application.ReadWrite.All
            Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All,Application.ReadWrite.All
        .LINK
            https://docs.microsoft.com/en-us/graph/api/serviceprincipal-post-approleassignments?view=graph-rest-1.0&tabs=http
            https://docs.microsoft.com/en-us/graph/api/serviceprincipal-delete-approleassignments?view=graph-rest-1.0&tabs=http
            https://docs.microsoft.com/en-us/graph/migrate-azure-ad-graph-configure-permissions#request-3
    #>
    [CmdletBinding(DefaultParameterSetName = "AddAdminConsent")]
    param(
        [parameter(
            ParameterSetName="AddAdminConsent",
            Mandatory=$true,
            Position=0
        )]
        [parameter(
            ParameterSetName="AddAppPermissions",
            Mandatory=$true,
            Position=0
        )]
        [parameter(
            ParameterSetName="RemoveAppRoleAssignments",
            Mandatory=$true,
            Position=0
        )]
        [System.String]
        $ObjectID,

        [parameter(
            ParameterSetName="AddAdminConsent",
            Mandatory=$true,
            Position=1
        )]
        [parameter(
            ParameterSetName="AddAppPermissions",
            Mandatory=$true,
            Position=1
        )]
        [parameter(
            ParameterSetName="RemoveAppRoleAssignments",
            Mandatory=$true,
            Position=1
        )]
        [System.String[]]
        [ValidateSet("Calendars.Read","Calendars.ReadWrite","Contacts.Read",
                    "Contacts.ReadWrite","full_access_as_app","IMAP.AccessAsApp",
                    "Mail.Read","Mail.ReadBasic","Mail.ReadBasic.All",
                    "Mail.ReadWrite","Mail.Send","MailboxSettings.Read",
                    "MailboxSettings.ReadWrite")]
        $Roles,

        [parameter(
            ParameterSetName="AddAppPermissions",
            Mandatory=$true,
            Position=2
        )]
        [System.Management.Automation.SwitchParameter]
        $AddApplicationPermissions,

        [parameter(
            ParameterSetName="RemoveAppRoleAssignments",
            Mandatory=$true,
            Position=2
        )]
        [System.Management.Automation.SwitchParameter]
        $RemoveAppRoleAssignment
    )

    begin
    {
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        # initiate variable
        $missingPermissions = [System.Collections.ArrayList]@()
        $existingAppRoleAssignment = [System.Collections.ArrayList]@()
        # check for required permissions
        $requiredMGPermissions = @('AppRoleAssignment.ReadWrite.All','Application.ReadWrite.All')
        # get current context
        $currentMGContext = Get-MgContext
        if (-not [System.String]::IsNullOrEmpty($currentMGContext))
        {
            foreach ($permission in $requiredMGPermissions)
            {
                if (-not $currentMGContext.Scopes.Contains($permission))
                {
                    Write-Error "Required permission missing:$($permission)"
                    [System.Boolean]$insufficientPerms = $true
                }
            }
            if ($insufficientPerms)
            {
                break
            }
        }
        else
        {
            Write-Warning 'No existing connection. Please connect to MS Graph first! e.g.:Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All,Application.ReadWrite.All'
            break
        }

        # retrieve serviceprincipal
        try
        {
            $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $ObjectID -ExpandProperty AppRoleAssignments -ErrorAction Stop
        }
        catch
        {
            $_.Exception
            break
        }
        # retrieve Exchange Online and Microsoft Graph servicePrincipals
        $MSGraphEXOSPN = Get-MgServicePrincipal -Filter "(AppId eq '00000002-0000-0ff1-ce00-000000000000') or (AppId eq '00000003-0000-0000-c000-000000000000')"
        # extract AppRoles
        $EXOAppRoles = ( $MSGraphEXOSPN | Where-Object { $_.AppID -eq '00000002-0000-0ff1-ce00-000000000000'} ).AppRoles | Where-Object { $_.Value -match "^(full_access_as_app|IMAP.AccessAsApp)$"}
        $MSGraphAppRoles = ( $MSGraphEXOSPN | Where-Object { $_.AppID -eq '00000003-0000-0000-c000-000000000000'} ).AppRoles | Where-Object { $_.Value -match 'Mail\.|MailboxSettings\.|Calendars\.|Contacts\.'}
        # get Ids
        $roleDetails = [System.Collections.ArrayList]@()
        $roleDetails += $EXOAppRoles | Where-Object {$_.Value -match "^($($Roles -join '|'))$"}
        $roleDetails += $MSGraphAppRoles | Where-Object {$_.Value -match "^($($Roles -join '|'))$"}
        
        if ($AddApplicationPermissions)
        {
            # get application
            $appFilter = "AppId eq '" + $servicePrincipal.AppId +"'"
            $app = Get-MgApplication -Filter $appFilter
            if (-not [System.String]::IsNullOrEmpty($app))
            {
                $EXOPerms = $EXOAppRoles | Where-Object {$_.Value -match "^($($Roles -join '|'))$"}
                $MSGraphPerms = $MSGraphAppRoles | Where-Object {$_.Value -match "^($($Roles -join '|'))$"}
            }
        }
    }

    process
    {
        # grant admin consent using AppRoleAssignments
        if ([System.String]::IsNullOrEmpty($servicePrincipal.AppRoleAssignments))
        {
            $missingPermissions = $roleDetails
        }
        else
        {
            # checking for existing permissions
            foreach ($role in $roleDetails)
            {
                if ($servicePrincipal.AppRoleAssignments.AppRoleId.Contains($role.Id))
                {
                    Write-Verbose "Admin consent exists:$($role.Value)"
                    if ($RemoveAppRoleAssignment)
                    {
                        # get all assignments to be removed
                        $existingAppRoleAssignment += $servicePrincipal.AppRoleAssignments | Where-Object {$_.AppRoleId -eq $role.Id}
                    }
                }
                else
                {
                    # add missing permissions for processing
                    [System.Void]$missingPermissions.Add($role)
                }
            }
        }

        if ($RemoveAppRoleAssignment)
        {
            if ([System.String]::IsNullOrEmpty($existingAppRoleAssignment))
            {
                Write-Host 'No admin consent exists. Nothing to do!'
            }
            else
            {
                foreach ($appRoleAssignment in $existingAppRoleAssignment)
                {
                    $paramsRemove = @{
                        AppRoleAssignmentId = $appRoleAssignment.Id
                        ServicePrincipalId = $ObjectID
                    }
                    Remove-MgServicePrincipalAppRoleAssignment @paramsRemove
                }
            }
        }
        else
        {

        if ([System.String]::IsNullOrEmpty($missingPermissions))
        {
            Write-Host 'All admin consent already exists. Nothing to do!'
        }
        else
        {
            Write-Verbose "The following admin consent is missing: $($missingPermissions.Value)"
            foreach ($permission in $missingPermissions)
            {
                # build body
                # checking for either MS Graph or EXO resource
                if ($permission.Value -match 'full_access_as_app|IMAP.AccessAsApp')
                {
                    Write-Verbose "Use EXO as resourceID as following permission is only there available:$($permission.Value)"
                    $resourceID = ($MSGraphEXOSPN | Where-Object {$_.AppID -eq '00000002-0000-0ff1-ce00-000000000000'}).Id
                }
                else
                {
                    Write-Verbose "Use MS Graph as resourceID as following permission is only there available:$($permission.Value)"
                    $resourceID = ($MSGraphEXOSPN | Where-Object {$_.AppID -eq '00000003-0000-0000-c000-000000000000'}).Id
                }

                $appRoleAssignment = @{
                    "principalId"= $ObjectID
                    "resourceId"= $resourceID
                    "appRoleId"= $permission.Id
                }

                try
                {
                    $newAppRoleAssignment = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ObjectID -BodyParameter $appRoleAssignment -ErrorAction Stop
                }
                catch
                {
                    $_.Exception
                }
            }
        }

        # add permission to application
        if ($AddApplicationPermissions)
        {
            if ([System.String]::IsNullOrEmpty($app))
            {
                Write-Host "App couldn't be found in tenant and therefore not updated. Please check manually!"
                break
            }

            if (-not [System.String]::IsNullOrEmpty($EXOPerms))
            {
                $EXOPermsExisting = $app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq '00000002-0000-0ff1-ce00-000000000000' }
                if ([System.String]::IsNullOrEmpty($EXOPermsExisting))
                {
                    # add EXO resource
                    $app.RequiredResourceAccess += @{
                        ResourceAppId = '00000002-0000-0ff1-ce00-000000000000'; 
                        ResourceAccess = @()
                    }

                    foreach ($role in $EXOPerms)
                    {
                        ($app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq '00000002-0000-0ff1-ce00-000000000000' }).ResourceAccess += @{
                                    Id = $role.Id
                                    Type ='Role'
                                }
                    }
                }
                else
                {
                    foreach ($role in $EXOPerms)
                    {
                        if (-not $EXOPermsExisting.ResourceAccess.Id.Contains($role.Id))
                        {
                            ($app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq '00000002-0000-0ff1-ce00-000000000000' }).ResourceAccess += @{
                                    Id = $role.Id
                                    Type ='Role'
                                }
                        }
                        else
                        {
                            Write-Verbose "Permission already exists:$($role.Id)"
                        }
                    }
                }
            }

            # get existing MS Graph permissions and add missing
            if (-not [System.String]::IsNullOrEmpty($MSGraphPerms))
            {
                $MSGraphPermsExisting = $app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq '00000003-0000-0000-c000-000000000000' }
                if ([System.String]::IsNullOrEmpty($MSGraphPermsExisting))
                {
                    # add MS Graph resource
                    $app.RequiredResourceAccess += @{
                        ResourceAppId = '00000003-0000-0000-c000-000000000000'; 
                        ResourceAccess = @()
                    }

                    foreach ($role in $MSGraphPerms)
                    {
                        ($app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq '00000003-0000-0000-c000-000000000000' }).ResourceAccess += @{
                                    Id = $role.Id
                                    Type ='Role'
                                }
                    }
                }
                else
                {
                    foreach ($role in $MSGraphPerms)
                    {
                        if (-not $MSGraphPermsExisting.ResourceAccess.Id.Contains($role.Id))
                        {
                            ($app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq '00000003-0000-0000-c000-000000000000' }).ResourceAccess += @{
                                    Id = $role.Id
                                    Type ='Role'
                                }
                        }
                        else
                        {
                            Write-Verbose "Permission already exists:$($role.Id)"
                        }
                    }
                }
            }

            # update existing application
            Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $app.RequiredResourceAccess 
        }
        
        }

    }

    end
    {
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }
}

function global:Get-EXOAllConnectionContexts
{
    Import-Module -Name ExchangeOnlineManagement -Force
    [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetAllConnectionContexts()
}

function global:Remove-EXO
{
    Disconnect-ExchangeOnline -Confirm:$false
}

function global:Format-MessageTraceDetail
{
    Param (
        [Parameter( Mandatory=$true, Position=0)]
        [System.Object[]]
        $MessageTraceDetail
    )

    foreach ($trace in $MessageTraceDetail)
    {
        # initiate custom objects
         $detailData = New-Object -TypeName psobject
         $detailObject = New-Object -TypeName psobject
         $detailObject | Add-Member -MemberType NoteProperty -Name Date -Value $trace.Date
         $detailObject | Add-Member -MemberType NoteProperty -Name Event -Value $trace.Event
         $detailObject | Add-Member -MemberType NoteProperty -Name Action -Value $trace.Action
         $detailObject | Add-Member -MemberType NoteProperty -Name MessageId -Value $trace.MessageId
         $detailObject | Add-Member -MemberType NoteProperty -Name MessageTraceId -Value $trace.MessageTraceId

        # create XML from data
        [System.Xml.XmlDocument]$xmlData = ''
        $xmlData = $trace.Data
        # get all attributes and add to custom object
        foreach ($node in $xmlData.root.MEP)
        {
            $detailData | Add-Member -MemberType NoteProperty -Name $($node.get_attributes().'#text'[0]) -Value $($node.get_attributes().'#text'[1])
        }

        $detailObject | Add-Member -MemberType NoteProperty -Name Data -Value $detailData
        $detailObject
    }
}

function global:Get-MgOauth2PermissionScopesOfMG
{
    $resourceMG = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
    $resourceMG.Oauth2PermissionScopes | Sort-Object -Property Value
}

