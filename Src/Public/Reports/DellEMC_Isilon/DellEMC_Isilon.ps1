#requires -Module @{ModuleName="PScribo";ModuleVersion="0.7.23"}

<#
.SYNOPSIS  
    PowerShell script to document the configuration of Dell EMC Isilon cluster in Word/HTML/XML/Text formats using PScribo.
.DESCRIPTION
    This PowerShell script has been tested with OneFS version 8.0 and later. 
.NOTES
    Version:        0.3
    Author:         Martin Cooper
    Twitter:        @mc1903
    Github:         mc1903
    Credits:        Iain Brighton (@iainbrighton) - PScribo
					Tim Carman (@tpcarman) - As Built Report

.LINK
    https://github.com/iainbrighton/PScribo
    https://github.com/tpcarman/As-Built-Report
#>

#region Configuration Settings
###############################################################################################
#                                    CONFIG SETTINGS                                          #
###############################################################################################

# If custom style not set, use Dell EMC Isilon style
if (!$StyleName) {
    .\Styles\DellEMC_Isilon.ps1
}

#endregion Configuration Settings


#region Script Functions
###############################################################################################
#                                    SCRIPT FUNCTIONS                                         #
###############################################################################################

## PlaceHolder

#endregion Script Functions

#region Script Body
###############################################################################################
#                                      SCRIPT BODY                                            #
###############################################################################################

#region Disable Certificate Validation
if ($Options.DisableCertificateValidation) {

add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

} 
#endregion Disable Certificate Validation


#region Use TLS 1.2 for OneFS 8.1.x & Later
if ($Options.UseTLSv12) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} else {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls
        }
#endregion region Use TLS 1.2 for OneFS 8.1.x & Later


#region Connect to each Isilon Cluster in turn
$IsiArrays = $Target.split(",")
foreach ($IsiArray in $IsiArrays) {

    #region Establish REST Session
    $Script:IsiPreAuthBody = ConvertTo-Json @{username= $Credentials.UserName; password = $Credentials.GetNetworkCredential().Password; services = ('platform','namespace')}

    $Script:IsiBaseURL = "https://$IsiArray`:8080"

    $Script:IsiNewSession = Invoke-RestMethod -Uri "$IsiBaseURL/session/1/session" -Body $IsiPreAuthBody -ContentType "application/json" -Method POST -SessionVariable IsiSession -TimeoutSec 180

    $Script:IsiCookies = $IsiSession.cookies.GetCookies($IsiBaseURL)
    try {
        $IsiCSRFCookie = [string]$IsiCookies['isicsrf']
        $IsiCSRFToken = $IsiCSRFCookie.split('=')[1]
        $IsiSession.Headers.Add('X-CSRF-Token', $IsiCSRFToken)
        $IsiSession.Headers.Add('Referer', $IsiBaseURL)
        } catch {
        }

    try {
        $Script:IsiApiVer = Invoke-Restmethod -method GET -Uri "$IsiBaseURL/platform/latest" -WebSession $IsiSession
        } catch { $IsiApiVer = [PSCustomObject]@{'latest' = 1}
        }

    $Script:IsiLatestApiBaseURL = "$IsiBaseURL/platform/$($IsiApiVer.latest)"

    #endregion Establish REST Session



    #region Get Isilon Cluster Information
	$Script:IsiClusterConfig = Invoke-Restmethod -Method GET -Uri "$IsiLatestApiBaseURL/cluster/config" -WebSession $IsiSession
    $Script:IsiClusterOnlineNodes = $IsiClusterConfig.devices | Where-Object {($_.is_up -eq "True")}
    $Script:IsiClusterIdentity = Invoke-Restmethod -Method GET -Uri "$IsiLatestApiBaseURL/cluster/identity" -WebSession $IsiSession
    $Script:IsiClusterOwner = Invoke-Restmethod -Method GET -Uri "$IsiLatestApiBaseURL/cluster/owner" -WebSession $IsiSession
    $Script:IsiClusterEmail = Invoke-Restmethod -Method GET -Uri "$IsiLatestApiBaseURL/cluster/email" -WebSession $IsiSession



    
    #endregion Get Isilon Cluster Information     
    
    Section -Style Heading1 "$($IsiClusterConfig.name) - Cluster Management" {

        #region Cluster Overview
        Section -Style Heading2 'Cluster Overview' {
            $ClusterOverview = [PSCustomObject]@{
                'Name' = $IsiClusterConfig.name
                'Description' = $IsiClusterConfig.description
                'GUID' = $IsiClusterConfig.guid
                'Total Nodes' = $IsiClusterConfig.devices.Count
                'Online Nodes' = $IsiClusterOnlineNodes.Count
                'Has Quorum' = $IsiClusterConfig.has_quorum
                'Join Mode' = $IsiClusterConfig.join_mode
                'Encoding' = $IsiClusterConfig.encoding
                'Operating System' = $IsiClusterConfig.onefs_version.type
                'Operating System Release' = $IsiClusterConfig.onefs_version.release
                'Operating System Build' = $IsiClusterConfig.onefs_version.build
            }
            $ClusterOverview | Table -Name 'Cluster Overview' -List -ColumnWidths 50, 50 
        }
        #endregion Cluster Overview 

        BlankLine
        
            #region Login Message
            Section -Style Heading3 'Login Message' {
                $LoginMessage = [PSCustomObject]@{
                   'Title' = $IsiClusterIdentity.logon.motd_header
                   'Description' = $IsiClusterIdentity.logon.motd
                }
                $LoginMessage | Table -Name 'Login Message' -List -ColumnWidths 20, 80 
			}
            #endregion Login Message
        
        BlankLine

            #region Contact Information
            Section -Style Heading3 'Contact Information' {
                $ClusterOwner = [PSCustomObject]@{
                    'Company' = $IsiClusterOwner.company
                    'Location' = $IsiClusterOwner.location
                    'Primary Contact Name' = $IsiClusterOwner.primary_name
                    'Primary Contact Email' = $IsiClusterOwner.primary_email
                    'Primary Contact Phone (Main)' = $IsiClusterOwner.primary_phone1
                    'Primary Contact Phone (Alt)' = $IsiClusterOwner.primary_phone2
                    'Secondary Contact Name' = $IsiClusterOwner.secondary_name
                    'Secondary Contact Email' = $IsiClusterOwner.secondary_email
                    'Secondary Contact Phone (Main)' = $IsiClusterOwner.secondary_phone1
                    'Secondary Contact Phone (Alt)' = $IsiClusterOwner.secondary_phone2
                } 
                $ClusterOwner | Table -Name 'Contact Information' -List -ColumnWidths 50, 50
			}
            #endregion Contact Information

            BlankLine

            #region Email Settings
            Section -Style Heading3 'Email Settings' {
                $ClusterEmail = [PSCustomObject]@{
                    'SMTP Relay Address' = $IsiClusterEmail.settings.mail_relay
                    'SMTP Relay Port' = $IsiClusterEmail.settings.smtp_port
                    'Send Email As' = $IsiClusterEmail.settings.mail_sender
                    'Subject Line' = $IsiClusterEmail.settings.mail_subject
                    'Notification Batch Mode' = $IsiClusterEmail.settings.batch_mode
                    'Use SMTP Authentication' = $IsiClusterEmail.settings.use_smtp_auth
                    'SMTP Authentication User Name' = $IsiClusterEmail.settings.smtp_auth_username
                    'SMTP Authentication Password Has Been Set' = $IsiClusterEmail.settings.smtp_auth_passwd_set
                    'SMTP Authentication Connection Security' = $IsiClusterEmail.settings.smtp_auth_security
                } 
                $ClusterEmail | Table -Name 'Email Settings' -List -ColumnWidths 50, 50
			}
            #endregion Email Settings

            BlankLine























    }
    PageBreak
    $Null = Invoke-Restmethod -Method DELETE -Uri "$IsiBaseURL/session/1/session" -WebSession $IsiSession
    }

#endregion Connect to each Isilon Cluster in turn
#endregion Script Body

<#

$xx = [PSCustomObject]@{
    'xx' = $IsiClusterIdentity.
    'xx' = $IsiClusterIdentity.
    'xx' = $IsiClusterIdentity.
    'xx' = $IsiClusterIdentity.
    'xx' = $IsiClusterIdentity.
    'xx' = $IsiClusterIdentity.
    'xx' = $IsiClusterIdentity.
    'xx' = $IsiClusterIdentity.
    'xx' = $IsiClusterIdentity.
}


#>