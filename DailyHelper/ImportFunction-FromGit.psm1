function global:ImportFunction-FromGit
{
    <#
    .SYNOPSIS
        This script is intended to import scripts from GitHub.
    .DESCRIPTION
        This script is intended to import scripts from GitHub as function into PowerShell
    .PARAMETER Url
        The parameter Url specifies the link to the raw GitHub content.
    .PARAMETER FunctionName
        The parameter FunctionName can be used to specify the name of the function.
    .PARAMETER AlreadyFunction
        The parameter AlreadyFunction shall be used, when the content contains functions.
    .EXAMPLE
        # import Get-Autodiscover function from PowerShell script
        ImportFunction-FromGit -Url 'https://raw.githubusercontent.com/IngoGege/Get-Autodiscover/main/Get-Autodiscover.ps1'
        # import Get-AccessToken function from PowerShell script with name Get-AADToken
        ImportFunction-FromGit -Url 'https://raw.githubusercontent.com/IngoGege/Get-AccessToken/master/Get-AccessToken.ps1' -FunctionName Get-AADToken
    .NOTES

    .LINK
        https://ingogegenwarth.wordpress.com/
    #>
    [CmdletBinding()]
    param(
        [parameter(
            mandatory = $true,
            Position = 0)]
        [System.Uri]
        $Url,

        [parameter(
            mandatory = $false,
            Position = 1)]
        [System.String]
        $FunctionName,

        [parameter(
            mandatory = $false,
            Position = 2)]
        [System.Management.Automation.SwitchParameter]
        $AlreadyFunction

    )

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        # retrieve code
        $code = Invoke-RestMethod -Method GET -Uri $Url

        if ([System.String]::IsNullOrEmpty($FunctionName))
        {
            Write-Verbose "FunctionName not given..."
            $FunctionName = ($Url.AbsoluteUri.ToString().Split('/')[ -1]).Split('.')[0]
            Write-Verbose "Using:$($FunctionName)..."
        }

        if ($AlreadyFunction)
        {
            Invoke-Expression $code
        }
        else
        {
            Invoke-Expression "function global:$($FunctionName) { $($code) }"
        }
    }
    catch {
        $_
    }

}
