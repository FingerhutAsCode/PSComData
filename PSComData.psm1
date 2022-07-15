<#
.DESCRIPTION

.NOTES
0.3.0 - Updated Username and Password variables to use Module Variable as described in https://thedavecarroll.com/powershell/how-i-implement-module-variables/

#>

$ComDataSession = [ordered]@{
    WebServicesURL  = "https://w6.iconnectdata.com/FleetCreditWS/services/FleetCreditWS0200"
    ComDataUsername = $null
    ComDataPassword = $null
}
New-Variable -Name ComDataSession -Value $ComDataSession -Scope Script -Force


function Set-UnsafeHeaderParsing {
    $netAssembly = [Reflection.Assembly]::GetAssembly([System.Net.Configuration.SettingsSection])
    if ($netAssembly) {
        $BindingFlags = [Reflection.BindingFlags] "Static,GetProperty,NonPublic"
        $SettingsType = $netAssembly.GetType("System.Net.Configuration.SettingsSectionInternal")
        $Instance = $SettingsType.InvokeMember("Section", $BindingFlags, $null, $null, @())
        if ($Instance) {
            $BindingFlags = "NonPublic", "Instance"
            $UseUnsafeHeaderParsingField = $SettingsType.GetField("useUnsafeHeaderParsing", $BindingFlags)
            if ($UseUnsafeHeaderParsingField) {
                $UseUnsafeHeaderParsingField.SetValue($Instance, $true)
            }
        }
    }
}

# Set the ComData Session Credentials
# The ComData Session Credentials must be set before any other functions can be called
function Set-ComDataSessionCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [String] $Username,
        [Parameter(Mandatory = $true)]
        [String] $Password
    )
    $ComDataSession.ComDataUsername = $Username
    $ComDataSession.ComDataPassword = $Password
}

function Set-ComDataSessionEnvironment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Certification', 'Production')]
        [String] $Environment
    )
    switch ($Environment) {
        Certification {
            $ComDataSession.WebServicesURL = 'https://w8cert.iconnectdata.com/FleetCreditWS/services/FleetCreditWS0200'
        }
        Production {
            $ComDataSession.WebServicesURL = 'https://api.iconnectdata.com/FleetCreditWS/services/FleetCreditWS0200'
        }
    }  
}

function Get-ComDataSessionVariables {
    $ComDataSession | ConvertTo-Json | ConvertFrom-Json | Format-List
}

function Test-ComDataSessionVariables {
    $Score = 0
    if ($null -ne $ComDataSession.ComDataUsername) {
        Write-Verbose "ComData Username confirmed to be not null"
        $Score++
    }
    if ($null -ne $ComDataSession.ComDataPassword) {
        Write-Verbose "ComData Password confirmed to be not null"
        $Score++
    }
    if ($null -ne $ComDataSession.WebServicesURL) {
        Write-Verbose "WebServicesURL confirmed to be not null"
        $Score++
    }
    if ($Score -eq 3) {
        return $true
    }
    else {
        return $false
    }
}

function Get-ComDataDriverList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [String] $AccountNumber
    )

    begin {
        if (-not(Test-ComDataSessionVariables)) {
            throw "ComData Session Variables have not been set properly!  Please call Set-ComDataSessionCredentials first."
        }
    }

    process {
        $SOAPActionHeader = @{"SOAPAction" = "http://fleetCredit02.comdata.com/FleetCreditWS0200/inquireDriverId" }

        [xml]$SOAP = '
            <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:main="http://fleetCredit02.comdata.com/maintenance/">
                <soapenv:Header>
                    <wsse:Security soapenv:mustUnderstand="1" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
                        <wsse:UsernameToken wsu:Id="UsernameToken-12" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                            <wsse:Username></wsse:Username>
                            <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"></wsse:Password>
                        </wsse:UsernameToken>
                    </wsse:Security>
                </soapenv:Header>
                <soapenv:Body>
                    <main:DriverIdInquireRequest>
                        <criteria>
                            <driverId></driverId>
                            <firstName></firstName>
                            <lastName></lastName>
                            <driverLicNbr></driverLicNbr>
                            <driverLicState></driverLicState>
                            <custId></custId>
                            <acctNbr></acctNbr>
                            <misc1></misc1>
                            <misc2></misc2>
                            <driverEmail></driverEmail>
                        </criteria>
                        <pageNbr>1</pageNbr>
                    </main:DriverIdInquireRequest>
                </soapenv:Body>
            </soapenv:Envelope>'
        
        Set-UnsafeHeaderParsing

        $SOAP.Envelope.Header.Security.UsernameToken.Username = $ComDataSession.ComDataUsername
        $SOAP.Envelope.Header.Security.UsernameToken.Password.InnerText = $ComDataSession.ComDataPassword
        $SOAP.Envelope.Body.DriverIdInquireRequest.criteria.acctNbr = $AccountNumber

        Write-Verbose $SOAP.OuterXml

        $Response = Invoke-WebRequest -Uri $ComDataSession.WebServicesURL -Method post -ContentType 'text/xml' -Body $SOAP -Headers $SOAPActionHeader

        $ResponseContent = [xml]$Response.Content

        $PageCount = $ResponseContent.Envelope.Body.DriverIdInquireResponse.pageCount
        $PageNumber = 1
        $DriverList = @()
        do {
            $SOAP.Envelope.Body.DriverIdInquireRequest.pageNbr = "$PageNumber"
            $Response = Invoke-WebRequest $ComDataSession.WebServicesURL -Method post -ContentType 'text/xml' -Body $SOAP -Headers $SOAPActionHeader
            $ResponseContent = [xml]$Response.Content
            $DriverList += $ResponseContent.Envelope.Body.DriverIdInquireResponse.records.driverIdSearchRecord
            $PageNumber++
        } while ($PageNumber -le $PageCount)
    }

    end {
        return $DriverList
    }
}

function Get-ComDataDriver {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [String] $AccountNumber,
        [Parameter(Mandatory = $true)]
        [String] $Username,
        [Parameter(Mandatory = $true)]
        [String] $Password
    )

    $SOAPActionHeader = @{"SOAPAction" = "http://fleetCredit02.comdata.com/FleetCreditWS0200/inquireDriverId" }
    
    [xml]$SOAP = '
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:main="http://fleetCredit02.comdata.com/maintenance/">
            <soapenv:Header>
                <wsse:Security soapenv:mustUnderstand="1" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
                    <wsse:UsernameToken wsu:Id="UsernameToken-12" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                        <wsse:Username></wsse:Username>
                        <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"></wsse:Password>
                    </wsse:UsernameToken>
                </wsse:Security>
            </soapenv:Header>
            <soapenv:Body>
                <main:DriverIdInquireRequest>
                    <criteria>
                        <driverId></driverId>
                        <firstName></firstName>
                        <lastName></lastName>
                        <driverLicNbr></driverLicNbr>
                        <driverLicState></driverLicState>
                        <custId></custId>
                        <acctNbr></acctNbr>
                        <misc1></misc1>
                        <misc2></misc2>
                        <driverEmail></driverEmail>
                    </criteria>
                    <pageNbr>1</pageNbr>
                </main:DriverIdInquireRequest>
            </soapenv:Body>
        </soapenv:Envelope>'
    
    $SOAP.Envelope.Header.Security.UsernameToken.Username = $Username
    $SOAP.Envelope.Header.Security.UsernameToken.Password = $Password
    $SOAP.Envelope.Body.DriverIdInquireRequest.criteria.acctNbr = $AccountNumber
    $SOAP.Envelope.Body.DriverIdInquireRequest.criteria.driverId = $DriverID

    Set-UnsafeHeaderParsing

    $Response = Invoke-WebRequest $ComDataWebServicesURL -Method post -ContentType 'text/xml' -Body $SOAP -Headers $SOAPActionHeader

    $ResponseContent = [xml]$Response.Content
    return $ResponseContent.Envelope.Body.DriverIdInquireResponse
}

function Remove-ComDataDriver {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [String] $AccountNumber,
        [Parameter(Mandatory = $true)]
        [String] $CustomerID,
        [Parameter(Mandatory = $true)]
        [String] $Username,
        [Parameter(Mandatory = $true)]
        [String] $Password
    )

    $SOAPActionHeader = @{"SOAPAction" = "http://fleetCredit02.comdata.com/FleetCreditWS0200/deleteDriverId" }
    
    [xml]$SOAP = '
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:main="http://fleetCredit02.comdata.com/maintenance/">
            <soapenv:Header>
                <wsse:Security soapenv:mustUnderstand="1" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
                    <wsse:UsernameToken wsu:Id="UsernameToken-12" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                        <wsse:Username></wsse:Username>
                        <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"></wsse:Password>
                    </wsse:UsernameToken>
                </wsse:Security>
            </soapenv:Header>
            <soapenv:Body>
                <main:DriverIdDeleteRequest>
                    <driverId></driverId>
                    <custId></custId>
                    <acctNbr></acctNbr>
                </main:DriverIdDeleteRequest>
            </soapenv:Body>
        </soapenv:Envelope>'
       
    $SOAP.Envelope.Header.Security.UsernameToken.Username = $Username
    $SOAP.Envelope.Header.Security.UsernameToken.Password = $Password
    $SOAP.Envelope.Body.DriverIdInquireRequest.criteria.acctNbr = $AccountNumber
    $SOAP.Envelope.Body.DriverIdInquireRequest.criteria.custId = $CustomerID

    $DriverQueryResponse = Get-ComDataDriver -DriverID $DriverID -LegalEntity $LegalEntity
    if ($DriverQueryResponse.responseCode -eq 0) {
        Set-UnsafeHeaderParsing
        $SOAP.Envelope.Body.DriverIdDeleteRequest.driverId = "$DriverID"
        $Response = Invoke-WebRequest $ComDataWebServicesURL -Method post -ContentType 'text/xml' -Body $SOAP -Headers $SOAPActionHeader
        $ResponseContent = [xml]$Response.Content
        if ($ResponseContent.Envelope.Body.DriverIdDeleteResponse.responseCode -eq "0") {
            Write-Host "Driver Removed Sucessfully"
        }
    }
    else {
        Write-Host "Error: DriverID $DriverID does not exists"
    } 
}

function New-ComDataDriver {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [String] $AccountNumber,
        [Parameter(Mandatory = $true)]
        [String] $CustomerID,
        [Parameter(Mandatory = $true)]
        [String] $Username,
        [Parameter(Mandatory = $true)]
        [String] $Password,
        [Parameter(Mandatory = $true)]
        [String] $DriverID,
        [Parameter(Mandatory = $true)]
        [String] $FirstName,
        [Parameter(Mandatory = $true)]
        [String] $LastName,
        [Parameter(Mandatory = $true)]
        [String] $EmployeeID,
        [Parameter(Mandatory = $true)]
        [String] $Email
    )
    
    $SOAPActionHeader = @{"SOAPAction" = "http://fleetCredit02.comdata.com/FleetCreditWS0200/addDriverId" }

    [xml]$SOAP = '
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:main="http://fleetCredit02.comdata.com/maintenance/">
            <soapenv:Header>
                <wsse:Security soapenv:mustUnderstand="1" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
                    <wsse:UsernameToken wsu:Id="UsernameToken-12" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                        <wsse:Username></wsse:Username>
                        <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"></wsse:Password>
                    </wsse:UsernameToken>
                </wsse:Security>
            </soapenv:Header>
            <soapenv:Body>
                <main:DriverIdAddRequest>
                    <record>
                        <driverId></driverId>
                        <firstName></firstName>
                        <lastName></lastName>
                        <driverLicNbr></driverLicNbr>
                        <driverLicState></driverLicState>
                        <custId></custId>
                        <acctNbr></acctNbr>
                        <misc1></misc1>
                        <misc2></misc2>
                        <driverEmail></driverEmail>
                    </record>
                </main:DriverIdAddRequest>
            </soapenv:Body>
        </soapenv:Envelope>'

    $DriverQueryResponse = Get-ComDataDriver -DriverID $DriverID -LegalEntity $LegalEntity 
    if ($DriverQueryResponse.responseCode -eq 1) {
        Set-UnsafeHeaderParsing

        $SOAP.Envelope.Header.Security.UsernameToken.Username = $Username
        $SOAP.Envelope.Header.Security.UsernameToken.Password = $Password
        $SOAP.Envelope.Body.DriverIdInquireRequest.criteria.acctNbr = $AccountNumber
        $SOAP.Envelope.Body.DriverIdInquireRequest.criteria.custId = $CustomerID
        $SOAP.Envelope.Body.DriverIdAddRequest.record.driverId = "$DriverID"
        $SOAP.Envelope.Body.DriverIdAddRequest.record.firstName = "$FirstName"
        $SOAP.Envelope.Body.DriverIdAddRequest.record.lastName = "$LastName"
        $SOAP.Envelope.Body.DriverIdAddRequest.record.misc1 = "$EmployeeID"
        
        $Response = Invoke-WebRequest $ComDataWebServicesURL -Method post -ContentType 'text/xml' -Body $SOAP -Headers $SOAPActionHeader
    
        $ResponseContent = [xml]$Response.Content
    
        if ($ResponseContent.Envelope.Body.DriverIdAddResponse.responseCode -eq "0") {
            Write-Host "Driver Added Sucessfully"
        }
    
    }
    else {
        Write-Host "Error: DriverID $DriverID already exists"
    } 
}

Export-ModuleMember -Function Set-ComDataSessionCredentials
Export-ModuleMember -Function Set-ComDataSessionEnvironment
Export-ModuleMember -Function Get-ComDataSessionVariables
Export-ModuleMember -Function Get-ComDataDriver
Export-ModuleMember -Function Get-ComDataDriverList
Export-ModuleMember -Function Remove-ComDataDriver
Export-ModuleMember -Function New-ComDataDriver