<#

WorldJournal.Email.psm1

    2017-??-?? Initial creation
    2018-05-11 Port to PowerShell module
    2018-05-11 Add 'pass' parameter to functions 
    2018-05-11 Fixed typo
    2018-05-14 Update Subject date format
    2018-05-22 'ScriptPath' parameter in Emailv2 is now mandatory
    2018-05-23 'Body' and 'ScriptPath' accepts empty string
    2018-05-24 Add 'Get-WJEmail' function, reads info from xml file in _DoNotRepository

'Emailv5' is not set to be exported in manifest, use v2 instead.
v5 uses 'Send-MailMessage' cmdlet, but it's only available in higher version of PowerShell.
v2 and v3 uses System.Net.Mail.MailMessage object which is compatiable with any version.

#>

$xmlPath = (Split-Path (Split-Path (Split-Path ($MyInvocation.MyCommand.Path) -Parent) -Parent) -Parent)+"\_DoNotRepository\"+(($MyInvocation.MyCommand.Name) -replace '.psm1', '.xml')
[xml]$xml = Get-Content $xmlPath -Encoding UTF8
$lyu = ($xml.Root.Email | Where-Object{$_.Name -eq 'lyu'}).MailAddress

function Get-WJEmail() {
    [CmdletBinding()]
    Param ()
    DynamicParam {

        $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary

        $attributes = New-Object System.Management.Automation.ParameterAttribute
        $attributes.Mandatory = $false
        $attributes.ParameterSetName = '__AllParameterSets'
        $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $attributeCollection.Add($attributes)
        $values = $xml.Root.Email.Name | Select-Object -Unique
        $validateSet = New-Object System.Management.Automation.ValidateSetAttribute($values)    
        $attributeCollection.Add($validateSet)
        $dynamicParam = New-Object -Type System.Management.Automation.RuntimeDefinedParameter(
            "Name", [string], $attributeCollection
        )

        $paramDictionary.Add("Name", $dynamicParam)

        return $paramDictionary

    }

    begin {}
    process {

        $Name = $PSBoundParameters.Name
    
        $whereArray = @()
        if ($Name -ne $null) { $whereArray += '$_.Name -eq $Name' }
        $whereString = $whereArray -Join " -and "  
        $whereBlock = [scriptblock]::Create( $whereString )

        if ($PSBoundParameters.Count -ne 0) {
            $xml.Root.Email | Where-Object -FilterScript $whereBlock | Select-Object Name, MailAddress, Password
        }
        else {
            $xml.Root.Email | Select-Object Name, MailAddress, Password
        }

    }
    end {}
}

Function Emailv2 {

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$From,
        [Parameter(Mandatory = $true)][string]$Pass,
        [Parameter(Mandatory = $true)][string]$To,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Body,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$ScriptPath,
        [Parameter(Mandatory = $false)][array]$Attachments
    )

    $User = $From

    $SMTPServer = "smtp.gmail.com"
    $SMTPPort = "587"

    $Mail = New-Object System.Net.Mail.MailMessage
    $Mail.From = $From
    $Mail.To.Add($To)
    $Mail.Bcc.Add($lyu)
    $Mail.Subject = $Subject + " " + (Get-Date).ToString("yyyy-MM-dd")
    $Mail.Body = $Body + "`n`n`n" +
    "--`n" +
    "Email Time: " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + "`n" + 
    "Script Path: " + $ScriptPath + "`n" + 
    "Computer Name: " + $env:computername

    If ($Attachments.Count -gt 0) {
        Foreach ($Attachment in $Attachments) {
            $Mail.Attachments.Add((New-Object Net.Mail.Attachment($Attachment)))        
        }
    }

    $SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($User, $Pass);
    #[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
    $SMTPClient.Send($Mail)
}

Function Emailv3 {

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$From,
        [Parameter(Mandatory = $true)][string]$Pass,
        [Parameter(Mandatory = $true)][string]$To,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Body,
        [Parameter(Mandatory = $false)][array]$Attachments
    )

    $User = $From

    $SMTPServer = "smtp.gmail.com"
    $SMTPPort = "587"

    $Mail = New-Object System.Net.Mail.MailMessage
    $Mail.From = $From
    $Mail.To.Add($To)
    $Mail.Bcc.Add($lyu)
    $Mail.Subject = $Subject
    $Mail.Body = $Body + "`n`n`n" +
    "--`n" +
    "Sent from World Journal MIS department"

    If ($Attachments.Count -gt 0) {
        Foreach ($Attachment in $Attachments) {
            $Mail.Attachments.Add((New-Object Net.Mail.Attachment($Attachment)))        
        }
    }

    $SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($User, $Pass);
    $SMTPClient.Send($Mail)
}



Function Emailv5 {

    # # # # # # # # # # # # # # #
    #                           #
    #    Use Emailv2 instead    #
    #                           #
    # # # # # # # # # # # # # # #

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$From,
        [Parameter(Mandatory = $true)][string]$Pass,
        [Parameter(Mandatory = $true)][string]$To,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$Body,
        [Parameter(Mandatory = $false)][array]$Attachments
    )
                    
    $getFrom = $From
    $getTo = $To
    $getSubject = $Subject + " " + (Get-Date).ToString("yyyy-MM-dd")
    $getBody = $Body + "`n`n`n" +
    "--`n" +
    "Email Time: " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + "`n" + 
    "Script Name: " + (Split-Path $MyInvocation.PSCommandPath -Leaf) + "`n" + 
    "Computer Name: " + $env:computername
    $getAttachments = $Attachments

    $SMTPserver = "smtp.gmail.com"
    $SMTPport = "587"
    $Pass = ConvertTo-SecureString $Pass -AsPlainText -Force
    $Cred = New-Object System.Management.Automation.PSCredential($getFrom, $Pass)
    $Bcc = $lyu

    If ($getAttachments.Count -gt 0) {
        Send-MailMessage -From $getFrom -to $getTo -Bcc $Bcc -Subject $getSubject -Body $getBody -Attachments $getAttachments `
            -SmtpServer $SMTPserver -port $SMTPport -UseSsl -Credential $Cred        
    }
    Else {
        Send-MailMessage -From $getFrom -to $getTo -Bcc $Bcc -Subject $getSubject -Body $getBody `
            -SmtpServer $SMTPserver -port $SMTPport -UseSsl -Credential $Cred    
    }
}