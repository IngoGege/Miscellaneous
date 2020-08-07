function global:Search-UnifiedLog
{
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
        $PageCount = '1',

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
        [ValidateSet('None', 'Failed', 'Pending', 'Delivered', 'Expanded')]
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
    if ((Get-PSSession).ComputerName -match 'compliance')
    {
        $ConnectedTo = 'SCC'
    }
    else
    {
        $ConnectedTo = 'EXO'
    }

    $Host.UI.RawUI.WindowTitle = (Get-Date -UFormat '%y/%m/%d %R').Tostring() + " Connected to $($ConnectedTo) as $((Get-PSSession ).Runspace.ConnectionInfo.Credential.UserName)"
    Write-Host '[' -NoNewline
    Write-Host (Get-Date -UFormat '%T')-NoNewline
    Write-Host ']:' -NoNewline
    Write-Host (Split-Path (Get-Location) -Leaf) -NoNewline
    return "> "
}
Prompt

function global:Get-ManagedFolderAssistantLog
{
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
            $data | add-member -type NoteProperty -Name Ecl -Value $(([xml](Export-MailboxDiagnosticLogs -Identity $ID -ExtendedProperties).MailboxLog).Properties.MailboxTable.Property | ? name -Like 'elc*')
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
        [ValidateSet('Bulk', 'Phish', 'Spam', 'Malware', 'TransportRule')]
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
        [ValidateSet('Bulk', 'Phish', 'Spam', 'TransportRule')]
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
    [CmdletBinding()]
    Param
    (
        [System.String]
        $UserPrincipalName,

        [System.String]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Search Administrator","External ID User Flow Attribute Administrator","Guest User","Power Platform Administrator","Cloud Application Administrator","Compliance Administrator","Security Administrator","Exchange Service Administrator","Restricted Guest User","Device Managers","Office Apps Administrator","Insights Business Leader","Desktop Analytics Administrator","Intune Service Administrator","B2C IEF Policy Administrator","CRM Service Administrator","Reports Reader","Partner Tier1 Support","License Administrator","Customer LockBox Access Approver","Security Reader","Security Operator","Global Administrator","Printer Administrator","Teams Service Administrator","External ID User Flow Administrator","Helpdesk Administrator","Azure Information Protection Administrator","Kaizala Administrator","Lync Service Administrator","Cloud Device Administrator","Message Center Reader","Privileged Authentication Administrator","Search Editor","Directory Readers","Hybrid Identity Administrator","Directory Writers","Guest Inviter","Password Administrator","Application Administrator","Device Join","Device Administrators","User","Power BI Service Administrator","B2C IEF Keyset Administrator","Message Center Privacy Reader","Billing Administrator","Conditional Access Administrator","Teams Communications Administrator","External Identity Provider Administrator","Workplace Device Join","Authentication Administrator","Application Developer","Directory Synchronization Accounts","Network Administrator","Device Users","Partner Tier2 Support","Azure DevOps Administrator","Compliance Data Administrator","Privileged Role Administrator","Printer Technician","Insights Administrator","Service Support Administrator","SharePoint Service Administrator","Global Reader","Teams Communications Support Engineer","Teams Communications Support Specialist","Groups Administrator","User Account Administrator")]
        $Role,

        [System.Int16]
        [ValidateRange(1,10)]
        $Hours = '10',

        [System.String]
        [ValidateNotNullOrEmpty()]
        $Reason = 'Daily work'
    )

    begin
    {
        $Error.Clear()
        Write-Verbose 'Remove existing "old" AzureAD module and load AzureADPreview'
        Remove-Module Azuread -Force -ErrorAction silentlycontinue
        Import-Module AzureADPreview -Verbose:$false
    }

    process
    {

        try {
            $AAD=Connect-AzureAD -AccountId $UserPrincipalname
            $resource = Get-AzureADMSPrivilegedResource -ProviderId AadRoles
            $roleDefinition = Get-AzureADMSPrivilegedRoleDefinition -ProviderId AadRoles -ResourceId $resource.Id -Filter "DisplayName eq '$Role'"
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
        catch {
            $Error[0].Exception
        }
    }
}

function global:Get-MSGraphGroup
{
[CmdletBinding()]
    param(
        [parameter( Position=0)]
        [System.String[]]$Group,

        [parameter( Position=1)]
        [System.String]
        $AccessToken,

        [parameter( Position=2)]
        [System.Management.Automation.SwitchParameter]
        $ByMail,

        [parameter( Position=3)]
        [System.String]
        $Filter,

        [parameter( Position=5)]
        [ValidateSet("login","select_account","consent","admin_consent","none")]
        [System.String]
        $PromptBehaviour = 'select_account'
    )

    begin
    {

        $Error.Clear()

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

            Begin
            {
                Add-Type -AssemblyName System.Web

                If ($V2)
                {
                    $OAuthSub = '/oauth2/v2.0/authorize?'
                }
                Else
                {
                    $OAuthSub = '/oauth2/authorize?'
                }

                #create autorithy Url
                $AuthUrl = $Authority.AbsoluteUri + $Tenant + $OAuthSub
                Write-Verbose -Message "AuthUrl:$($AuthUrl)"

                #create empty body variable
                $Body = @{}
                $Url_String = ''

                Function Show-OAuthWindow
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

                    $global:form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
                    $global:web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url ) }
                    $DocComp  = {
                        $Global:uri = $web.Url.AbsoluteUri
                        if ($Global:Uri -match "error=[^&]*|code=[^&]*|code=[^#]*|#access_token=*")
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
                    $global:result = $web
                    $output = @{}
                    foreach($key in $queryOutput.Keys){
                        $output["$key"] = $queryOutput[$key]
                    }

                    $output
                }
            }

            Process
            {
                $Params = $PSBoundParameters.GetEnumerator() | Where-Object -FilterScript {$_.key -inotmatch 'Verbose|v2|authority|tenant|Redirect_Uri'}
                foreach ($Param in $Params)
                {
                    Write-Verbose -Message "$($Param.Key)=$($Param.Value)"
                    $Url_String += "&" + $Param.Key + '=' + [System.Web.HttpUtility]::UrlEncode($Param.Value)
                }

                If ($Redirect_Uri)
                {
                    $Url_String += "&Redirect_Uri=$Redirect_Uri"
                }
                $Url_String = $Url_String.TrimStart("&")
                Write-Verbose "RedirectURI:$($Redirect_Uri)"
                Write-Verbose "URL:$($Url_String)"
                $Response = Show-OAuthWindow -Url $($AuthUrl + $Url_String) -Response_Mode $Response_Mode
            }

            End
            {
                If ($Response.Count -gt 0)
                {
                    $Response
                }
                Else
                {
                    Write-Verbose "Error occured"
                    Add-Type -AssemblyName System.Web
                    [System.Web.HttpUtility]::UrlDecode($result.Url.OriginalString)
                }
            }
        }

        $timer = [System.Diagnostics.Stopwatch]::StartNew()

        if ($AccessToken)
        {
            $script:token = $AccessToken
        }
        else
        {
            try {
                # get code
                $authParams = @{
                    Authority = 'https://login.microsoftonline.com/'
                    Tenant = 'common'
                    Client_ID = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
                    Redirect_Uri = 'urn:ietf:wg:oauth:2.0:oob'
                    Resource = 'https://graph.microsoft.com'
                    Prompt = $PromptBehaviour
                    Response_Mode = 'query'
                    Response_Type = 'code'
                }

                $script:authCode = Get-AADAuth @authParams

                if ( [System.String]::IsNullOrEmpty($authCode.code) )
                {
                    Write-Host "AuthCode is NULL! Stopping..."
                    break
                }

                # create body
                $body = @{
                    client_id = $authParams.Client_ID
                    code = $($authCode['code'])
                    redirect_uri = $authParams.Redirect_URI
                    grant_type = "authorization_code"
                }

                $params = @{
                    ContentType = 'application/x-www-form-urlencoded'
                    Method = 'POST'
                    Uri = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
                    Body = $body
                }

                $script:token = (Invoke-RestMethod @params).access_token
            }

            catch {
                $Error[0].Exception
            }
        }

        $collection = [System.Collections.ArrayList]@()

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
            Write-Verbose 'Found custom Filter. Will try to find user based on...'
            $filterParams = @{
                ContentType = 'application/json'
                Method = 'GET'
                Headers = @{ Authorization = "Bearer $($token)"}
                Uri = 'https://graph.microsoft.com/beta/groups?$filter=' + $Filter
                ErrorAction = 'Stop'
            }

            try {
                $Group = (Invoke-RestMethod @filterParams).value.id
            }
            catch
            {
                $_
            }

            if ($Group.count -eq 0)
            {
                Write-Verbose $('No group found for filter "' + $Filter + '"! Terminate now...')
                break
            }
            else
            {
                Write-Verbose "Found $($Group.count) groups..."
            }
        }

    }

    process
    {

        foreach($object in $Group)
        {
            try {
                # get group id
                if ($ByMail)
                {
                    Write-Verbose 'Get group by email...'
                    $byMailParams = @{
                        Uri = "https://graph.microsoft.com/beta/groups?filter=mail eq '$($object)'"
                        Method = 'GET'
                        Headers = @{ Authorization = "Bearer $($token)"}
                        ErrorAction = 'Stop'
                    }

                    #$responseByMail = Invoke-RestMethod @byMailParams
                    $id = (Invoke-RestMethod @byMailParams).value.id
                    #$id = $responseByMail.value.id
                }
                else
                {
                    Write-Verbose 'Get group by id...'
                    $id = $object #(Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/$group" -Method GET -Headers @{ Authorization = "Bearer $($token.access_token)"}).value.id
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
                            url = "/groups/$id/members"
                            method = 'GET'
                            id = '3'
                        },
                        @{
                            url = "/groups/$id/sites/root"
                            method = 'GET'
                            id = '4'
                        }
                    )
                }

                $restParams = @{
                    ContentType = 'application/json'
                    Method = 'POST'
                    Headers = @{ Authorization = "Bearer $($token)"}
                    Body = $body | ConvertTo-Json -Depth 4
                    Uri = 'https://graph.microsoft.com/beta/$batch'
                    ErrorAction = 'Stop'
                }

                $global:data = Invoke-RestMethod @restParams

                # create custom object
                $groupInfo = $null
                # check for error
                $groupResponse = $data.responses | Where-Object -FilterScript { $_.id -eq 1}
                $groupInfo = $groupResponse.Body | Select-Object * -ExcludeProperty "@odata.context"

                if (($groupResponse.status -ne 200) -and ('MailboxNotEnabledForRESTAPI' -eq $groupResponse.body.error.code))
                {
                    Write-Verbose "Error MailboxNotEnabledForRESTAPI thrown. WIll try again without certain properties..."

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
                $ownerResponse = $data.responses | Where-Object -FilterScript {$_.id -eq 2}
                
                if ('200' -eq $ownerResponse.status)
                {
                    $groupObject | Add-Member -MemberType NoteProperty -Name Owners -Value @( $($ownerResponse.body.value | Select-Object * -ExcludeProperty "@odata.type"))
                }
                else
                {
                    Write-Verbose "Error found in response for owners..."
                    $groupObject | Add-Member -MemberType NoteProperty -Name Owners -Value @( $($ownerResponse.body.error))
                }

                # add members to object
                $memberResponse = $data.responses | Where-Object -FilterScript {$_.id -eq 3}

                if ('200' -eq $memberResponse.status)
                {
                    $groupObject | Add-Member -MemberType NoteProperty -Name Members -Value @( $($memberResponse.body.value | Select-Object * -ExcludeProperty "@odata.type"))
                }
                else
                {
                    Write-Verbose "Error found in response for members..."
                    $groupObject | Add-Member -MemberType NoteProperty -Name Members -Value @( $($memberResponse.body.error))
                }

                # add root site to object
                $siteResponse = $data.responses | Where-Object -FilterScript {$_.id -eq 4}
                
                if ('200' -eq $siteResponse.status)
                {
                    $groupObject | Add-Member -MemberType NoteProperty -Name Sites -Value @( $($siteResponse.body.value | Select-Object * -ExcludeProperty "@odata.type"))
                }
                else
                {
                    Write-Verbose "Error found in response for sites..."
                    $groupObject | Add-Member -MemberType NoteProperty -Name Sites -Value @( $($siteResponse.body.error))
                }

                $collection += $groupObject

            }
            catch
            {
                $_
                Write-Verbose "Error occured for group $($object)..."
                break
            }

        }
    }

    end
    {
        $collection
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }

}

function global:Get-MSGraphUser
{
    #https://docs.microsoft.com/en-us/graph/query-parameters#filter-parameter
    #https://docs.oasis-open.org/odata/odata/v4.0/errata03/os/complete/part2-url-conventions/odata-v4.0-errata03-os-part2-url-conventions-complete.html#_Toc453752358
    [CmdletBinding()]
    param(
        [parameter( Position=0)]
        [System.String[]]
        $User,

        [parameter( Position=1)]
        [System.String]
        $AccessToken = $MSGraphToken[0].AccessToken,

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
        $GetAuthMethods

    )

    begin
    {

        $Error.Clear()

        $collection = [System.Collections.ArrayList]@()

        if ($Filter)
        {
            Write-Verbose 'Found custom Filter. Will try to find user based on...'
            $filterParams = @{
                ContentType = 'application/json'
                Method = 'GET'
                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                Uri = 'https://graph.microsoft.com/beta/users?$filter=' + $Filter
                ErrorAction = 'Stop'
            }

            try {
                $user = (Invoke-RestMethod @filterParams).value.id
            }
            catch
            {
                $_
            }

            if ($user.count -eq 0)
            {
                Write-Verbose 'No user found for filter $($Filter)! Terminate now...'
                break
            }
            else
            {
                Write-Verbose "Found $($user.count) user..."
            }
        }
    }

    process
    {
        foreach ($account in $User)
        {

            try {
                $body = @{
                    requests = @(
                        @{
                            url = "/users/$($account)" + '?$select=*'
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
                        }
                    )
                }

                if ($GetMailboxSettings)
                {
                    $mailboxsettings = @{
                            url = "/users/$($account)" + '?$select=mailboxSettings'
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
                }

                $restParams = @{
                    ContentType = 'application/json'
                    Method = 'POST'
                    Headers = @{ Authorization = "Bearer $($AccessToken)"}
                    Body = $body | ConvertTo-Json -Depth 4
                    Uri = 'https://graph.microsoft.com/beta/$batch'
                }

                $data = Invoke-RestMethod @restParams

                # create custom object
                $userObject = New-Object -TypeName psobject
                $userInfo = $null
                $userInfo = ($data.responses | Where-Object -FilterScript { $_.id -eq 1}).Body | Select-Object * -ExcludeProperty "@odata.context"
                $userProperties = $userInfo | Get-Member -MemberType NoteProperty

                foreach ($property in $userProperties)
                {
                    $userObject | Add-Member -MemberType NoteProperty -Name $( $property.Name ) -Value $( $userInfo.$( $property.Name ) )
                }

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
                        [System.Int16]$counter = '1'
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

                            Write-Verbose "Loopcount:$($counter)..."
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
                    # retrieve joined teams
                    $teamsParams = @{
                                ContentType = 'application/json'
                                Method = 'GET'
                                Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                Uri = "https://graph.microsoft.com/beta/users/$($userInfo.id)/joinedTeams"
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
                    
                    if ($GetDeltaToken)
                    {
                        Write-Verbose "Get delta for $($userInfo.userPrincipalName)"
                        $deltaParams = @{
                                ContentType = 'application/json'
                                Method = 'GET'
                                Headers = @{ Authorization = "Bearer $($AccessToken)"; prefer = "return=minimal"}
                                #Headers = @{ Authorization = "Bearer $($AccessToken)"}
                                Uri = 'https://graph.microsoft.com/beta/users/delta?' + '$filter=id eq ' + "'$($userInfo.id)'" + '&$deltaToken=latest'
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
                }

                if ($GetMailboxSettings)
                {
                    $responseMailboxsettings = $data.responses | Where-Object -FilterScript {$_.id -eq 5}
                    
                    if ('200' -eq $responseMailboxsettings.status)
                    {
                        $userObject | Add-Member -MemberType NoteProperty -Name MailboxSettings -Value  @( $($responseMailboxsettings.body.mailboxSettings) )
                    }
                    else
                    {
                        $userObject | Add-Member -MemberType NoteProperty -Name MailboxSettings -Value  @( $($responseMailboxsettings.body.error) )
                    }
                }

                if ($GetAuthMethods)
                {
                    $userObject | Add-Member -MemberType NoteProperty -Name AuthenticationMethods -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 6) -and ($_.status -eq 200)}).body.value | Select-Object * -ExcludeProperty "@odata.type" ))
                    $userObject | Add-Member -MemberType NoteProperty -Name PasswordMethods -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 7) -and ($_.status -eq 200)}).body.value ))
                    $userObject | Add-Member -MemberType NoteProperty -Name PhoneMethods -Value @($( ($data.responses | Where-Object -FilterScript { ($_.id -eq 8) -and ($_.status -eq 200)}).body.value ))
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
                        }

                        $responseApp = Invoke-RestMethod @appParams

                        $applicationCollection += $responseApp | Select-Object * -ExcludeProperty "@odata.context"
                    }

                    $userObject | Add-Member -MemberType NoteProperty -Name ApplicationDetails -Value @( $applicationCollection )
                
                }

                $collection += $userObject

            }
            catch
            {
                $_
                Write-Verbose "Error occured for account $($account)..."
                break
            }
        }
    }

    end
    {
        $collection
    }
}

function global:Get-RESTAzKeyVaultSecret
{
#https://docs.microsoft.com/rest/api/keyvault/getsecrets/getsecrets
#https://docs.microsoft.com/rest/api/keyvault/getsecret/getsecret
#https://docs.microsoft.com/rest/api/keyvault/getcertificate/getcertificate

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

