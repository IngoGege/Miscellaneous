function global:Test-SendMail
{
    [CmdletBinding()]
    param (
        [Parameter(
            ValueFromPipelineByPropertyName=$true,
            Mandatory=$true,
            Position=0)]
        [Alias('PrimarySmtpAddress')]
        [System.String[]]
        $EmailAddress,

        [Parameter(
            ValueFromPipelineByPropertyName=$false,
            Mandatory=$false,
            Position=1)]
        [System.String]
        $FromAddress,

        [Parameter(
            ValueFromPipelineByPropertyName=$false,
            Mandatory=$false,
            Position=2)]
        [System.String[]]
        $Recipients,

        [Parameter(
            Mandatory=$false,
            Position=3)]
        [System.String]
        $AccessToken,

        [Parameter(
            Mandatory=$false,
            Position=4)]
        [System.Management.Automation.SwitchParameter]
        $UseRest,

        [Parameter(
            Mandatory=$false,
            Position=5)]
        [System.Management.Automation.SwitchParameter]
        $Impersonate,

        [Parameter(
            Mandatory=$false,
            Position=6)]
        [System.String]
        $Subject = "Testmail $(Get-Date -Format 'yyyyMMdd HHmmssfff')",

        [Parameter(
            Mandatory=$false,
            Position=7)]
        [System.String]
        $Server,

        [Parameter(
            Mandatory=$false,
            Position=8)]
        [System.Management.Automation.SwitchParameter]
        $TrustAnySSL,

        [Parameter(
            Mandatory=$false,
            Position=9)]
        [ValidateScript({if (Test-Path $_ -PathType leaf){$True} else {Throw "WebServices DLL could not be found!"}})]
        [System.String]
        $WebServicesDLL

    )

    begin
    {

        function Get-AutoDV2
        {
            [CmdletBinding()]
            param (
                [Parameter(
                    Mandatory=$true,
                    Position=0)]
                [System.String]
                $EmailAddress,

                [Parameter(
                    Mandatory=$false,
                    Position=1)]
                [System.String]
                $Server,

                [Parameter(
                    Mandatory=$true,
                    Position=2)]
                [ValidateSet("AutodiscoverV1","ActiveSync","Ews","Rest","Substrate","SubstrateNotificationService","SubstrateSearchService","OutlookMeetingScheduler")]
                [System.String]
                $Protocol

            )

            try
            {
                if ($Server)
                {
                    #$Domain = $EmailAddress.Split("@")[1]
                    #$Server = "autodiscover." + $Domain
                    $URL = "https://$server/autodiscover/autodiscover.json?Email=$EmailAddress&Protocol=$Protocol"
                }
                else
                {
                    $URL = "https://autodiscover-s.outlook.com//autodiscover/autodiscover.json?Email=$EmailAddress&Protocol=$Protocol"
                }
                Write-Verbose "URL=$($Url)"
                Invoke-RestMethod -Uri $Url
            }
            catch
            {
                #create object
                $returnValue = New-Object -TypeName PSObject
                #get all properties from last error
                $ErrorProperties =$Error[0] | Get-Member -MemberType Property
                #add existing properties to object
                foreach ($Property in $ErrorProperties)
                {
                    if ($Property.Name -eq 'InvocationInfo')
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                    }
                    else
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                    }
                }
                #return object
                $returnValue
                break
            }
        }

        $timer = [System.Diagnostics.Stopwatch]::StartNew()

}

    process
    {

        try
        {
            foreach ($MailboxName in $EmailAddress)
            {

                if ($UseRest)
                {
                    Write-Verbose "Sending e-mail using REST..."
                    # create body
                    $i = 0
                    $message = '
                        {
                        "message":{
                            "subject": "",
                            "body": {
                            "contentType": "",
                            "content": ""
                            },
                            "toRecipients": [
                    '

                    foreach ($Recipient in $Recipients)
                    {
                        $i++
                        $message += @"
                                {
                                    "emailAddress": {
                                    "address": "$Recipient"
                                }
                                }
"@
                        if ($i -lt $Recipients.Count)
                        {
                            $message += ","
                        }
                    }

                    $message += '],
                            "from": {
                            "emailAddress": {
                                "address": ""
                            }
                            }
                        },
                        "saveToSentItems": "true"
                    }'

                    $message = $message | ConvertFrom-Json
                    $message.message.subject = $Subject
                    $message.message.body.contentType = 'HTML'
                    $message.message.body.content = 'Test 123'

                    if (-not [System.String]::IsNullOrWhiteSpace($FromAddress))
                    {
                        $message.message.from.emailAddress.address = $FromAddress
                        $senderAddress = $FromAddress
                    }
                    else
                    {
                        $message.message.from.emailAddress.address = $MailboxName
                        $senderAddress = $MailboxName
                    }

                    $param = @{
                        Method = 'POST'
                        Uri = "https://graph.microsoft.com/v1.0/users/$($senderAddress)/sendMail"
                        Headers = @{'Authorization'="$($AccessToken)"; 'Content-type'="application/json";'X-AnchorMailbox'=$($senderAddress)}
                        Body = ($message | ConvertTo-Json -Depth 4 | Out-String)
                    }

                    Invoke-RestMethod @param
                }
                else
                {
                    Write-Verbose "Sending e-mail using EWS..."
                    [System.String]$RootFolder="MsgFolderRoot"

                    if ($WebServicesDLL)
                    {
                        try
                        {
                            $EWSDLL = $WebServicesDLL
                            Import-Module -Name $EWSDLL
                        }
                        catch
                        {
                            $Error[0].Exception
                            exit
                        }
                    }
                    else
                    {
                        ## Load Managed API dll
                        ###CHECK FOR EWS MANAGED API, if PRESENT IMPORT THE HIGHEST VERSION EWS DLL, else EXIT
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
                            exit
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


                    if ($TrustAnySSL)
                    {
                        ## Choose to ignore any SSL Warning issues caused by Self Signed Certificates
                        ## Code From http://poshcode.org/624
                        ## Create a compilation environment
                        $Provider=New-Object -TypeName Microsoft.CSharp.CSharpCodeProvider
                        $Compiler=$Provider.CreateCompiler()
                        $Params=New-Object -TypeName System.CodeDom.Compiler.CompilerParameters
                        $Params.GenerateExecutable=$False
                        $Params.GenerateInMemory=$True
                        $Params.IncludeDebugInformation=$False
                        $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
                        $TASource=@'
namespace Local.ToolkitExtensions.Net.CertificatePolicy{
public class TrustAll : System.Net.ICertificatePolicy {
public TrustAll(){
}
public bool CheckValidationResult(System.Net.ServicePoint sp,
System.Security.Cryptography.X509Certificates.X509Certificate cert,
System.Net.WebRequest req, int problem){
return true;
}
}
}
'@
                        $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
                        $TAAssembly=$TAResults.CompiledAssembly
                        ## We now create an instance of the TrustAll and attach it to the ServicePointManager
                        $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
                        [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
                        ## end code from http://poshcode.org/624
                    }

                    ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use
                    if ($Server)
                    {
                        #CAS URL Option 2 Hardcoded
                        $uri=[system.URI] "https://$server/ews/exchange.asmx"
                        $service.Url = $uri
                    }
                    else
                    {
                        if (-not [System.String]::IsNullOrEmpty($FromAddress))
                        {
                            $targetAddress = $FromAddress
                        }
                        else
                        {
                            $targetAddress = $MailboxName
                        }
                        $service.Url = $(Get-AutoDV2 -EmailAddress $targetAddress -Protocol EWS).Url
                    }

                    ## Optional section for Exchange Impersonation
                    if ($Impersonate){
                        $Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
                    }

                    #increase performance by adding headers
                    if ($Service.HttpHeaders.keys.Contains("X-AnchorMailbox"))
                    {
                        $Service.HttpHeaders.Remove("X-AnchorMailbox") | Out-Null
                    }
                    $Service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName)

                    # create message object
                    $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($Service)
                    $message.Subject = $Subject
                    $message.Body = "Test 123"

                    foreach ($Recipient in $Recipients)
                    {
                        $message.ToRecipients.Add($Recipient) | Out-Null
                    }

                    if (-not [System.String]::IsNullOrWhiteSpace($FromAddress))
                    {
                        $message.From = $FromAddress
                    }
                    else
                    {
                        $message.From = $MailboxName
                    }

                    #$message.SendAndSaveCopy([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems)
                    $message.Send()
                    }
                }
            }

        catch
        {
            #create object
            $returnValue = New-Object -TypeName PSObject
            #get all properties from last error
            $ErrorProperties = $Error[0] | Get-Member -MemberType Property
            #add existing properties to object
            foreach ($Property in $ErrorProperties)
            {
                if ($Property.Name -eq 'InvocationInfo')
                {
                    $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                }
                else
                {
                    $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                }
            }
            #return object
            $returnValue
        }
    }

    end
    {
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
    }

}

function global:Get-CalendarEvents
{
    [CmdletBinding()]
    param(
        [Parameter(
            ValueFromPipeline=$true,
            Mandatory=$true,
            Position=0)]
        [ValidateNotNull()]
        [Alias('PrimarySmtpAddress')]
        [System.String[]]
        $EmailAddress,

        [Parameter(
            ValueFromPipeline=$true,
            Mandatory=$true,
            Position=1)]
        [ValidateNotNull()]
        [System.String]
        $AccessToken,

        [Parameter(
            Mandatory=$false,
            Position=2)]
        [System.Management.Automation.SwitchParameter]
        $UseEWS,

        [Parameter(
            Mandatory=$false,
            Position=3)]
        [System.Management.Automation.SwitchParameter]
        $UseMSGraph,

        [Parameter(
            Mandatory=$false,
            Position=4)]
        [System.Management.Automation.SwitchParameter]
        $Impersonate,

        [Parameter(
            Mandatory=$false,
            Position=5)]
        [System.String]
        $Server,

        [Parameter(
            Mandatory=$false,
            Position=6)]
        [ValidateScript({If (Test-Path $_ -PathType leaf){$True} Else {Throw "WebServices DLL could not be found!"}})]
        [System.String]
        $WebServicesDLL

    )

    begin
    {
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $objcol = [System.Collections.ArrayList]@()
        if($UseMSGraph)
        {
            $baseURI = 'https://graph.microsoft.com/v1.0/users/'
        }
        else
        {
            $baseURI = 'https://outlook.office.com/api/v2.0/users/'
        }

        Write-Verbose "BaseURI:$($baseURI)"

        function Get-AutoDV2
        {
            [CmdletBinding()]
            Param (
                [Parameter(
                    Mandatory=$true,
                    Position=0)]
                [System.String]
                $EmailAddress,

                [Parameter(
                    Mandatory=$false,
                    Position=1)]
                [System.String]
                $Server,
    
                [Parameter(
                    Mandatory=$true,
                    Position=2)]
                [ValidateSet("AutodiscoverV1","ActiveSync","Ews","Rest","Substrate","SubstrateNotificationService","SubstrateSearchService","OutlookMeetingScheduler")]
                [System.String]
                $Protocol
        
            )
            try
            {
                If ($Server)
                {
                    #$Domain = $EmailAddress.Split("@")[1]
                    #$Server = "autodiscover." + $Domain
                    $URL = "https://$server/autodiscover/autodiscover.json?Email=$EmailAddress&Protocol=$Protocol"
                }
                Else
                {
                    $URL = "https://autodiscover-s.outlook.com//autodiscover/autodiscover.json?Email=$EmailAddress&Protocol=$Protocol"
                }
                Write-Verbose "URL=$($Url)"
                Invoke-RestMethod -Uri $Url
            }
            catch
            {
                #create object
                $returnValue = New-Object -TypeName PSObject
                #get all properties from last error
                $ErrorProperties =$Error[0] | Get-Member -MemberType Property
                #add existing properties to object
                foreach ($Property in $ErrorProperties)
                {
                    if ($Property.Name -eq 'InvocationInfo')
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                    }
                    else
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                    }
                }
                #return object
                $returnValue
                break
            }
        }

    }

    process
    {
        foreach ($Mailbox in $EmailAddress)
        {
            try
            {
                if($UseEWS)
                {
                    Write-Verbose 'Using EWS for retrieving Calendar events...'
                    [System.String]$RootFolder="MsgFolderRoot"

                    if ($WebServicesDLL)
                    {
                        try
                        {
                            $EWSDLL = $WebServicesDLL
                            Import-Module -Name $EWSDLL
                        }
                        catch
                        {
                            $Error[0].Exception
                            exit
                        }
                    }
                    else
                    {
                        ## Load Managed API dll
                        ###CHECK FOR EWS MANAGED API, if PRESENT IMPORT THE HIGHEST VERSION EWS DLL, else EXIT
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
                            exit
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
                    ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use
                    if ($Server)
                    {
                        #CAS URL Option 2 Hardcoded
                        $uri=[system.URI] "https://$server/ews/exchange.asmx"
                        $service.Url = $uri
                    }
                    else
                    {
                        $service.Url = $(Get-AutoDV2 -EmailAddress $Mailbox -Protocol EWS).Url
                    }

                    ## Optional section for Exchange Impersonation
                    if ($Impersonate){
                        $Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox)
                    }

                    #increase performance by adding headers
                    if ($Service.HttpHeaders.keys.Contains("X-AnchorMailbox"))
                    {
                        $Service.HttpHeaders.Remove("X-AnchorMailbox") | Out-Null
                    }

                    $Service.HttpHeaders.Add("X-AnchorMailbox", $Mailbox)

                    $calFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$Mailbox) 
                    $calFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$calFolderID)
                    $ItemPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                    $ItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(20,0,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
                    $ItemView.PropertySet = $ItemPropset
                    $searchResult = $calFolder.FindItems($ItemView)
                    $objcol.Add($searchResult.Items) | Out-Null

                }
                else
                {
                    Write-Verbose 'Using Rest for retrieving Calendar events...'
                    # create parameterset
                    $param = @{
                        Method = 'GET'
                        Uri = $baseURI + $mailbox + '/events?$top=20&$select=subject,start,end,organizer'
                        Headers = @{'Authorization'="$($AccessToken)";'X-AnchorMailbox'=$($mailbox)}
                    }

                    $result = Invoke-RestMethod @param
                    foreach($event in $result.value)
                    {
                        $objcol.Add($event) | Out-Null
                    }
                }
            }
            catch
            {
                #create object
                $returnValue = New-Object -TypeName PSObject
                #get all properties from last error
                $ErrorProperties = $Error[0] | Get-Member -MemberType Property
                #add existing properties to object
                foreach ($Property in $ErrorProperties)
                {
                    if ($Property.Name -eq 'InvocationInfo')
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                    }
                    else
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                    }
                }
                #return object
                $returnValue
            }

        }
    }

    end
    {
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
        $objcol
    }

}

function global:Create-CalendarEvent
{
    [CmdletBinding()]
    param(
        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=0)]
        [ValidateNotNull()]
        [Alias('PrimarySmtpAddress')]
        [System.String[]]
        $EmailAddress,

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$true,
            Position=1)]
        [ValidateNotNull()]
        [System.String]
        $Attendee,

        [Parameter(
            ValueFromPipeline=$true,
            Mandatory=$true,
            Position=2)]
        [ValidateNotNull()]
        [System.String]
        $AccessToken,

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=3)]
        [ValidateNotNull()]
        [System.String]
        $TimeZone = 'W. Europe Standard Time',

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=4)]
        [ValidateNotNull()]
        [System.DateTime]
        $StartDate = $((Get-Date).AddHours(1)),

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=5)]
        [ValidateNotNull()]
        [System.DateTime]
        $EndDate = $((Get-Date).AddHours(1)),

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=6)]
        [ValidateNotNull()]
        [System.String]
        $Subject = "Test Event $(Get-Date -Format s)",

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=6)]
        [ValidateNotNull()]
        [System.String]
        $Content = "Event created via MS Graph.",

        [Parameter(
            Mandatory=$false,
            Position=6)]
        [System.Management.Automation.SwitchParameter]
        $UseOutlook,

        [Parameter(
            Mandatory=$false,
            Position=7)]
        [System.String]
        $O365GroupID

    )

    begin
    {
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $objcol = [System.Collections.ArrayList]@()
        if (-not [System.String]::IsNullOrEmpty($O365GroupID) )
        {
            $objectType = 'groups'
            $EmailAddress = $O365GroupID
        }
        else
        {
            $objectType = 'users'
        }
        if($UseOutlook)
        {
            $baseURI = "https://outlook.office.com/api/v2.0/$objectType/"
        }
        else
        {
            $baseURI = "https://graph.microsoft.com/v1.0/$objectType/"
        }

        Write-Verbose "BaseURI:$($baseURI)"

    }

    process
    {
        foreach ($Mailbox in $EmailAddress)
        {
            try
            {

                # create parameterset
                $global:event = @"
                    {
                    "subject": "$($Subject)",
                    "body": {
                        "contentType": "HTML",
                        "content": "$($Content)"
                    },
                    "start": {
                        "dateTime": "$(Get-Date $StartDate -Format s)",
                        "timeZone": "$($TimeZone)"
                    },
                    "end": {
                        "dateTime": "$(Get-Date $EndDate -Format s)",
                        "timeZone": "$($TimeZone)"
                    },
                    "location":{
                        "displayName":"TBD"
                    },
                    "attendees": [
                        {
                        "emailAddress": {
                            "address":"$($Attendee)",
                            "name":"bla"
                        },
                        "type": "required"
                        }
                    ]
                    }
"@

                $body = [System.Text.Encoding]::UTF8.GetBytes($event)

                $param = @{
                    Method = 'Post'
                    Uri = $baseURI + $mailbox + '/calendar/events'
                    Headers = @{
                        'Authorization' = "$($AccessToken)";
                        'X-AnchorMailbox' = $($mailbox);
                        'Content-type' = 'application/json; charset=utf-8'
                        }
                }

                $result = Invoke-RestMethod @param -Body $body

            }
            catch
            {
                #create object
                $returnValue = New-Object -TypeName PSObject
                #get all properties from last error
                $ErrorProperties = $Error[0] | Get-Member -MemberType Property
                #add existing properties to object
                foreach ($Property in $ErrorProperties)
                {
                    if ($Property.Name -eq 'InvocationInfo')
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                    }
                    else
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                    }
                }
                #return object
                $returnValue
            }

        }
    }

    end
    {
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
        $result
    }

}

function global:Create-OnlineMeeting
{
    [CmdletBinding()]
    param(
        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=0)]
        [ValidateNotNull()]
        [System.String]
        $OrganizerAddress,

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=0)]
        [ValidateNotNull()]
        [System.String[]]
        $Attendee,

        [Parameter(
            ValueFromPipeline=$true,
            Mandatory=$true,
            Position=1)]
        [ValidateNotNull()]
        [System.String]
        $AccessToken,

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=2)]
        [ValidateNotNull()]
        [System.String]
        $TimeZone = 'W. Europe Standard Time',

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=3)]
        [ValidateNotNull()]
        [System.DateTime]
        $StartDate = $((Get-Date).AddHours(1)),

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=4)]
        [ValidateNotNull()]
        [System.DateTime]
        $EndDate = $((Get-Date).AddHours(1)),

        [Parameter(
            ValueFromPipeline=$false,
            Mandatory=$false,
            Position=5)]
        [ValidateNotNull()]
        [System.String]
        $Subject = "Test Event $(Get-Date -Format s)"

    )

    begin
    {
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $objcol = [System.Collections.ArrayList]@()

    }

    process
    {
        foreach ($Organizer in $OrganizerAddress)
        {
            try
            {

                # create parameterset
                $global:OnlineMeeting = @"
                    {
                        "startDateTime":"$(Get-Date $StartDate -Format o)",
                        "endDateTime":"$(Get-Date $EndDate -Format o)",
                        "subject":"$($Subject)",
                        "participants":{ 
                            "organizer":{"upn":"$($Organizer)"},
                            "attendees":[
"@
                $i = 0
                Write-Verbose $Attendee.Count
                foreach ($A in $Attendee)
                {
                    $i++
                    Write-Verbose "Adding attendee:$($A)"
                    $global:OnlineMeeting += @"
                    {"upn":"$A"}
"@
                    if ($i -lt $Attendee.Count)
                    {
                        $global:OnlineMeeting += ","
                    }
                }
                $global:OnlineMeeting += @"
                        ]
                    }
                }
"@

                $param = @{
                    Method = 'Post'
                    Uri = 'https://graph.microsoft.com/v1.0/me/onlineMeetings'
                    Headers = @{
                        'Authorization' = "$($AccessToken)";
                        'Content-type' = 'application/json'
                        }
                }

                $result = Invoke-RestMethod @param -Body $OnlineMeeting

            }
            catch
            {
                #create object
                $returnValue = New-Object -TypeName PSObject
                #get all properties from last error
                $ErrorProperties = $Error[0] | Get-Member -MemberType Property
                #add existing properties to object
                foreach ($Property in $ErrorProperties)
                {
                    if ($Property.Name -eq 'InvocationInfo')
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                    }
                    else
                    {
                        $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                    }
                }
                #return object
                $returnValue
            }

        }
    }

    end
    {
        $timer.Stop()
        Write-Verbose "ScriptRuntime:$($timer.Elapsed.ToString())"
        $result
    }

}

